"""Security questionnaire agent built on Box AI.

When a prospect's questionnaire (.xlsx) lands in a deal's Security Review
folder, the agent:
  1. parses the questions with a spreadsheet parser (deterministic, no LLM)
  2. asks Box AI to answer each one from the approved Trust Center content
  3. flags weak answers: no citation, "NOT_FOUND", or a source past its review date
  4. writes a DRAFT copy with Response, Source and Confidence columns
  5. sets securityReview metadata and assigns a review task to Security
When the reviewer approves the task, it sends the NDA with Box Sign (if set up)
and, once signed, releases the SOC 2 report through an expiring, view-only link.

Usage:
  python agent.py process <file_id>    run the pipeline on one file
  python agent.py watch                poll the Deals folder (no public URL needed)
  python agent.py serve                webhook receiver (use with ngrok)
  python agent.py approve <draft_id>   run the approval path by hand
  python agent.py release <draft_id>   create the expiring SOC 2 link by hand
  python agent.py ask "question"       ask the Trust Center one question
  python agent.py reset                forget processed files (demo rehearsal)
"""
import argparse
import io
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timedelta, timezone

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from box_sdk_gen.box.errors import BoxAPIError
from box_sdk_gen.managers.ai import CreateAiAskMode
from box_sdk_gen.managers.file_metadata import (
    CreateFileMetadataByIdScope,
    DeleteFileMetadataByIdScope,
    GetFileMetadataByIdScope,
)
from box_sdk_gen.managers.shared_links_files import (
    AddShareLinkToFileSharedLink,
    AddShareLinkToFileSharedLinkAccessField,
    AddShareLinkToFileSharedLinkPermissionsField,
)
from box_sdk_gen.managers.task_assignments import CreateTaskAssignmentAssignTo, CreateTaskAssignmentTask
from box_sdk_gen.managers.comments import CreateCommentItem, CreateCommentItemTypeField
from box_sdk_gen.managers.tasks import CreateTaskAction, CreateTaskItem, CreateTaskItemTypeField
from box_sdk_gen.managers.uploads import (
    UploadFileAttributes,
    UploadFileAttributesParentField,
    UploadFileVersionAttributes,
)
from box_sdk_gen.schemas.ai_item_ask import AiItemAsk, AiItemAskTypeField
from box_sdk_gen.schemas.file_base import FileBase
from box_sdk_gen.schemas.folder_mini import FolderMini
from box_sdk_gen.schemas.sign_request_create_signer import SignRequestCreateSigner

from box_common import TEMPLATE_KEY, env, get_client, load_state, log, save_state

DRAFT_SUFFIX = " - DRAFT"
PROMPT = (
    "You are drafting an answer to a customer's security questionnaire on behalf of the vendor. "
    "Answer ONLY from the provided documents, in the vendor's voice, in at most three sentences. "
    "If the documents do not clearly answer the question, reply with exactly NOT_FOUND.\n\n"
    "Question: {question}"
)
FLAG_FILL = PatternFill("solid", fgColor="FDE7D4")
OK_FILL = PatternFill("solid", fgColor="E3F4E8")
REVIEW_BY = re.compile(r"review-by:\s*(\d{4}-\d{2}-\d{2})", re.I)


# ---------------------------------------------------------------- knowledge
def trust_sources(client):
    """Return (items for Box AI, {file_id: (name, review_by_date)})."""
    folder_id = env("TRUST_FOLDER_ID", required=True)
    entries = client.folders.get_folder_items(folder_id, fields=["name", "description", "type"], limit=100).entries
    files = {
        e.id: (e.name, _review_date(getattr(e, "description", "") or ""))
        for e in entries
        if e.type == "file"
    }
    hub_id = env("TRUST_HUB_ID")
    if hub_id:  # ask over the whole Hub when one exists
        return [AiItemAsk(id=hub_id, type=AiItemAskTypeField.HUBS)], files
    ids = list(files)[:25]  # Box AI multi-item Q&A takes up to 25 files
    return [AiItemAsk(id=i, type=AiItemAskTypeField.FILE) for i in ids], files


def _review_date(description):
    m = REVIEW_BY.search(description)
    return date.fromisoformat(m.group(1)) if m else None


def ask_box_ai(client, items, question):
    mode = CreateAiAskMode.SINGLE_ITEM_QA if len(items) == 1 else CreateAiAskMode.MULTIPLE_ITEM_QA
    resp = client.ai.create_ai_ask(mode, PROMPT.format(question=question), items, include_citations=True)
    answer = (resp.answer or "").strip() if resp else ""
    cited = []
    for c in (resp.citations or []) if resp else []:
        if c.id and c.id not in [x[0] for x in cited]:
            cited.append((c.id, c.name or c.id))
    return answer, cited


def judge(answer, cited, sources):
    """Deterministic confidence rules: the part a reviewer can audit."""
    if not answer or "NOT_FOUND" in answer:
        return "LOW", "No approved source answers this"
    if not cited:
        return "LOW", "Answer has no citation"
    today = date.today()
    stale = [sources[i][0] for i, _ in cited if i in sources and sources[i][1] and sources[i][1] < today]
    if stale:
        return "LOW", f"Source past review date: {stale[0]}"
    return "HIGH", ""


# ---------------------------------------------------------------- pipeline
def process_file(client, file_id, state=None):
    state = state or load_state()
    info = client.files.get_file_by_id(file_id, fields=["name", "parent", "file_version", "path_collection"])
    version = info.file_version.id if info.file_version else "0"
    if not info.name.lower().endswith(".xlsx") or DRAFT_SUFFIX in info.name:
        log("skip", f"{info.name} is not a questionnaire")
        return
    if state["processed"].get(file_id) == version:
        log("skip", f"{info.name} version {version} already processed")
        return
    account = _account_name(info)
    log("received", f"{info.name} for {account}")

    # 1. parse questions deterministically
    wb = load_workbook(io.BytesIO(client.downloads.download_file(file_id).read()))
    ws = wb.active
    header_row, cols = _find_header(ws)
    questions = [
        (r, ws.cell(r, cols["id"]).value if cols.get("id") else f"row {r}", str(ws.cell(r, cols["question"]).value).strip())
        for r in range(header_row + 1, ws.max_row + 1)
        if ws.cell(r, cols["question"]).value
    ]
    log("parsed", f"{len(questions)} questions (header on row {header_row})")

    # 2. ask Box AI, a few questions at a time
    items, sources = trust_sources(client)
    log("sources", f"{len(sources)} approved Trust Center files" + (" via Hub" if env("TRUST_HUB_ID") else ""))

    def work(q):
        row, qid, text = q
        try:
            answer, cited = ask_box_ai(client, items, text)
        except BoxAPIError as e:
            return row, qid, "", [], "LOW", f"Box AI error {e.response_info.status_code}"
        conf, reason = judge(answer, cited, sources)
        return row, qid, answer, cited, conf, reason

    with ThreadPoolExecutor(max_workers=int(env("AI_WORKERS", "4"))) as pool:
        results = list(pool.map(work, questions))
    flagged = [r for r in results if r[4] == "LOW"]
    for _, qid, _, cited, conf, reason in results:
        src = ", ".join(n for _, n in cited)[:60] or "-"
        log("answer", f"{qid}  {conf:<4} {reason or src}")

    # 3. write the draft workbook
    resp_col = cols.get("response") or ws.max_column + 1
    base = ws.max_column + 1
    for i, title in enumerate(["Source (Box)", "Confidence", "Reviewer note"]):
        c = ws.cell(header_row, base + i, title)
        c.font = Font(bold=True)
    for row, qid, answer, cited, conf, reason in results:
        ws.cell(row, resp_col, "" if conf == "LOW" and "NOT_FOUND" in answer else answer)
        ws.cell(row, base, "; ".join(f"{n} (file {i})" for i, n in cited))
        ws.cell(row, base + 1, conf).fill = FLAG_FILL if conf == "LOW" else OK_FILL
        ws.cell(row, base + 2, reason)
        for c in (ws.cell(row, resp_col), ws.cell(row, base), ws.cell(row, base + 2)):
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for col, width in ((base, 40), (base + 1, 12), (base + 2, 36)):
        ws.column_dimensions[ws.cell(1, col).column_letter].width = width
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    draft_name = info.name[:-5] + DRAFT_SUFFIX + ".xlsx"
    draft_id = _upload(client, info.parent.id, draft_name, buf)
    log("draft", f"uploaded {draft_name} (file {draft_id})")

    # 4. metadata the whole team can see and sort by
    auto = len(results) - len(flagged)
    due = datetime.now(timezone.utc) + timedelta(days=int(env("REVIEW_DAYS", "3")))
    meta = {
        "account": account,
        "status": "In review",
        "questionsTotal": len(results),
        "autoAnswered": round(100 * auto / max(len(results), 1)),
        "flaggedCount": len(flagged),
        "reviewer": env("REVIEWER_LOGIN", ""),
        "dueDate": due.strftime("%Y-%m-%dT00:00:00Z"),
    }
    set_metadata(client, draft_id, meta)
    log("metadata", f"{auto}/{len(results)} drafted, {len(flagged)} flagged, due {due:%b %d}")

    # 5. review task for Security
    task_id = None
    reviewer = env("REVIEWER_LOGIN")
    if env("USE_AUTOMATE", "false").lower() == "true":
        reviewer = None  # Box Automate creates the approval task and sends the NDA
        log("handoff", "status 'In review' set; Box Automate takes over approval and Box Sign")
    if reviewer:
        msg = (f"Security review for {account}: {len(flagged)} answers need your input "
               f"({', '.join(str(r[1]) for r in flagged)}). Approve the task when the draft is ready to send.")
        task = client.tasks.create_task(
            CreateTaskItem(id=draft_id, type=CreateTaskItemTypeField.FILE),
            action=CreateTaskAction.REVIEW, message=msg[:500], due_at=due,
        )
        client.task_assignments.create_task_assignment(
            CreateTaskAssignmentTask(id=task.id), CreateTaskAssignmentAssignTo(login=reviewer)
        )
        task_id = task.id
        log("task", f"review task assigned to {reviewer}")

    # 6. a summary comment, so people see the agent's work right in Box
    lines = [f"Questionnaire agent: drafted {auto} of {len(results)} answers from the Trust Center with Box AI. "
             f"{len(flagged)} need Security review:"]
    for _, qid, _, cited, _, reason in flagged:
        lines.append(f"- {qid}: {reason}")
    lines.append("Every drafted answer cites its source file in the 'Source (Box)' column.")
    post_comment(client, draft_id, "\n".join(lines))

    state["processed"][file_id] = version
    state["drafts"][draft_id] = {"account": account, "folder": info.parent.id, "task": task_id, "status": "In review"}
    save_state(state)
    log("done", f"{account}: draft ready for Security review")
    return draft_id


def on_approved(client, draft_id, state):
    d = state["drafts"][draft_id]
    d["status"] = "Approved"
    set_metadata(client, draft_id, {"status": "Approved"}, merge=True)
    log("approved", f"{d['account']} draft approved by Security")
    nda, prospect = env("NDA_FILE_ID"), env("PROSPECT_EMAIL")
    if nda and prospect:
        try:
            sr = client.sign_requests.create_sign_request(
                [SignRequestCreateSigner(email=prospect)],
                source_files=[FileBase(id=nda)],
                parent_folder=FolderMini(id=d["folder"]),
                name=f"Mutual NDA - {d['account']}",
                email_subject="Mutual NDA before we share our SOC 2 report",
            )
            state["sign_requests"][sr.id] = draft_id
            d["status"] = "NDA sent"
            log("sign", f"NDA sent to {prospect} with Box Sign")
        except BoxAPIError as e:
            log("sign", f"Box Sign unavailable ({e.response_info.status_code}); releasing directly")
            release(client, draft_id, state)
    else:
        release(client, draft_id, state)
    save_state(state)


def release(client, draft_id, state=None):
    state = state or load_state()
    soc2 = env("SOC2_FILE_ID", required=True)
    expires = datetime.now(timezone.utc) + timedelta(days=int(env("LINK_DAYS", "7")))
    link = AddShareLinkToFileSharedLink(
        access=AddShareLinkToFileSharedLinkAccessField.OPEN,
        password=env("LINK_PASSWORD") or None,
        unshared_at=expires,
        permissions=AddShareLinkToFileSharedLinkPermissionsField(can_download=False),
    )
    try:
        f = client.shared_links_files.add_share_link_to_file(soc2, "shared_link", shared_link=link)
    except BoxAPIError as e:  # some plans don't allow passwords or expiry
        log("release", f"plan rejected link settings ({e.response_info.status_code}); retrying view-only link")
        f = client.shared_links_files.add_share_link_to_file(
            soc2, "shared_link",
            shared_link=AddShareLinkToFileSharedLink(
                access=AddShareLinkToFileSharedLinkAccessField.OPEN,
                permissions=AddShareLinkToFileSharedLinkPermissionsField(can_download=False),
            ),
        )
    if draft_id in state["drafts"]:
        state["drafts"][draft_id]["status"] = "Sent"
        save_state(state)
    set_metadata(client, draft_id, {"status": "Sent"}, merge=True)
    post_comment(client, draft_id, f"Questionnaire agent: NDA complete. SOC 2 report released through a view-only link "
                                   f"that expires {expires:%b %d}: {f.shared_link.url}")
    log("release", f"SOC 2 link (view-only, expires {expires:%b %d}): {f.shared_link.url}")


# ---------------------------------------------------------------- helpers
def _find_header(ws):
    for r in range(1, min(ws.max_row, 20) + 1):
        cols = {}
        for c in range(1, ws.max_column + 1):
            v = str(ws.cell(r, c).value or "").strip().lower()
            if len(v) > 30:  # titles and instructions are not column headings
                continue
            if v in ("id", "#", "no", "no.", "ref"):
                cols["id"] = c
            elif v.startswith("question"):
                cols["question"] = c
            elif "response" in v or "answer" in v:
                cols["response"] = c
        if "question" in cols:
            return r, cols
    raise ValueError("No 'Question' column found in the first 20 rows")


def _account_name(info):
    """Deals / <Account> / Security Review  ->  <Account>."""
    path = [e.name for e in (info.path_collection.entries if info.path_collection else [])]
    if len(path) >= 2 and path[-1].lower() == "security review":
        return path[-2]
    return path[-1] if path else "Unknown account"


def _upload(client, folder_id, name, buf):
    existing = [e for e in client.folders.get_folder_items(folder_id, limit=1000).entries if e.name == name]
    if existing:
        return client.uploads.upload_file_version(
            existing[0].id, UploadFileVersionAttributes(name=name), buf
        ).entries[0].id
    return client.uploads.upload_file(
        UploadFileAttributes(name=name, parent=UploadFileAttributesParentField(id=folder_id)), buf
    ).entries[0].id


def set_metadata(client, file_id, values, merge=False):
    """Create the securityReview instance, or replace it (merging old values)."""
    scope = CreateFileMetadataByIdScope.ENTERPRISE
    try:
        client.file_metadata.create_file_metadata_by_id(file_id, scope, TEMPLATE_KEY, values)
        return
    except BoxAPIError as e:
        if e.response_info.status_code != 409:
            log("metadata", f"could not set metadata ({e.response_info.status_code}); is the template created?")
            return
    if merge:
        current = client.file_metadata.get_file_metadata_by_id(file_id, GetFileMetadataByIdScope.ENTERPRISE, TEMPLATE_KEY)
        raw = current.to_dict()
        raw = raw.get("extra_data") or raw  # the SDK keeps template fields under extra_data
        keep = {k: v for k, v in raw.items() if not k.startswith("$") and k not in ("extra_data", "can_edit", "canEdit")}
        values = {**keep, **values}
    client.file_metadata.delete_file_metadata_by_id(file_id, DeleteFileMetadataByIdScope.ENTERPRISE, TEMPLATE_KEY)
    client.file_metadata.create_file_metadata_by_id(file_id, scope, TEMPLATE_KEY, values)


def post_comment(client, file_id, message):
    """Leave a visible note on the file in Box; never fail the run over it."""
    try:
        client.comments.create_comment(message[:9000], CreateCommentItem(id=file_id, type=CreateCommentItemTypeField.FILE))
        log("comment", "summary posted on the draft in Box")
    except BoxAPIError as e:
        log("comment", f"could not post comment ({e.response_info.status_code})")


def _walk_xlsx(client, folder_id, depth=0):
    for e in client.folders.get_folder_items(folder_id, fields=["name", "type"], limit=1000).entries:
        if e.type == "folder" and depth < 3:
            yield from _walk_xlsx(client, e.id, depth + 1)
        elif e.type == "file" and e.name.lower().endswith(".xlsx") and DRAFT_SUFFIX not in e.name:
            yield e


def _check_approvals(client, state):
    for draft_id, d in list(state["drafts"].items()):
        if d.get("status") == "In review" and d.get("task"):
            for a in client.task_assignments.get_task_assignments(d["task"]).entries or []:
                if str(a.resolution_state.value if a.resolution_state else "") == "approved":
                    on_approved(client, draft_id, state)
                    break
    for sr_id, draft_id in list(state["sign_requests"].items()):
        sr = client.sign_requests.get_sign_request_by_id(sr_id)
        if str(sr.status.value if sr.status else "") == "signed":
            log("sign", "prospect signed the NDA")
            state["sign_requests"].pop(sr_id)
            release(client, draft_id, state)


# ---------------------------------------------------------------- entry points
def watch(client):
    deals = env("DEALS_FOLDER_ID", required=True)
    interval = int(env("POLL_SECONDS", "10"))
    log("watch", f"polling Deals folder {deals} every {interval}s (Ctrl+C to stop)")
    while True:
        state = load_state()
        for f in _walk_xlsx(client, deals):
            if f.id not in state["processed"]:
                process_file(client, f.id, state)
        _check_approvals(client, state)
        save_state(state)
        time.sleep(interval)


def serve(client):
    from flask import Flask, request
    from box_sdk_gen.managers.webhooks import WebhooksManager

    app = Flask(__name__)
    primary, secondary = env("WEBHOOK_PRIMARY_KEY", required=True), env("WEBHOOK_SECONDARY_KEY")

    @app.post("/box/webhook")
    def webhook():
        body = request.get_data(as_text=True)
        if not WebhooksManager.validate_message(body, dict(request.headers), primary, secondary_key=secondary):
            log("webhook", "rejected: bad signature")
            return "invalid signature", 403
        event = request.get_json()
        trigger, source = event.get("trigger"), event.get("source", {})
        log("webhook", f"{trigger} on {source.get('type')} {source.get('id')}")
        if trigger == "FILE.UPLOADED":
            threading.Thread(target=process_file, args=(client, source["id"]), daemon=True).start()
        elif trigger in ("TASK_ASSIGNMENT.UPDATED", "SIGN_REQUEST.COMPLETED"):
            threading.Thread(target=_check_approvals, args=(client, load_state()), daemon=True).start()
        return "", 200

    app.run(port=int(env("PORT", "8080")))


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("command", choices=["process", "watch", "serve", "approve", "release", "ask", "reset"])
    p.add_argument("arg", nargs="?")
    a = p.parse_args()
    if a.command == "reset":
        from box_common import STATE_FILE
        STATE_FILE.unlink(missing_ok=True)
        return log("reset", "state cleared; delete old DRAFT files in Box before rehearsing")
    client = get_client()
    if a.command == "process":
        process_file(client, a.arg)
    elif a.command == "watch":
        watch(client)
    elif a.command == "serve":
        serve(client)
    elif a.command == "approve":
        on_approved(client, a.arg, load_state())
    elif a.command == "release":
        release(client, a.arg)
    elif a.command == "ask":
        items, sources = trust_sources(client)
        answer, cited = ask_box_ai(client, items, a.arg)
        conf, reason = judge(answer, cited, sources)
        print(f"\n{answer}\n\nSources: {', '.join(n for _, n in cited) or 'none'}\nConfidence: {conf} {reason}")


if __name__ == "__main__":
    main()
