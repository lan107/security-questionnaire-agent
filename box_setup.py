"""One-time setup of the demo environment in Box.

  python box_setup.py                 folders, metadata template, sample content
  python box_setup.py --webhook URL   also register the webhook (e.g. your ngrok URL)

Safe to re-run: existing folders and files are reused, not duplicated.
"""
import argparse
from pathlib import Path

from box_sdk_gen.box.errors import BoxAPIError
from box_sdk_gen.managers.folders import CreateFolderParent
from box_sdk_gen.managers.metadata_templates import (
    CreateMetadataTemplateFields as F,
    CreateMetadataTemplateFieldsOptionsField as Opt,
    CreateMetadataTemplateFieldsTypeField as T,
)
from box_sdk_gen.managers.uploads import UploadFileAttributes, UploadFileAttributesParentField
from box_sdk_gen.managers.webhooks import CreateWebhookTarget, CreateWebhookTargetTypeField, CreateWebhookTriggers

from box_common import TEMPLATE_KEY, env, get_client, log
from make_content import OUT, TRUST_DOCS, main as make_content

STATUSES = ["Drafting", "In review", "Approved", "NDA sent", "Sent"]


def child(client, parent_id, name, kind):
    for e in client.folders.get_folder_items(parent_id, fields=["name", "type"], limit=1000).entries:
        if e.name == name and e.type == kind:
            return e.id
    return None


def folder(client, parent_id, name):
    fid = child(client, parent_id, name, "folder")
    if fid:
        return fid
    fid = client.folders.create_folder(name, CreateFolderParent(id=parent_id)).id
    log("folder", f"created {name}")
    return fid


def upload(client, parent_id, path: Path, description=None):
    fid = child(client, parent_id, path.name, "file")
    if not fid:
        with path.open("rb") as fh:
            fid = client.uploads.upload_file(
                UploadFileAttributes(name=path.name, parent=UploadFileAttributesParentField(id=parent_id)), fh
            ).entries[0].id
        log("upload", path.name)
    if description:
        client.files.update_file_by_id(fid, description=description)
    return fid


def metadata_template(client):
    fields = [
        F(type=T.STRING, key="account", display_name="Account"),
        F(type=T.ENUM, key="status", display_name="Status", options=[Opt(key=s) for s in STATUSES]),
        F(type=T.DATE, key="dueDate", display_name="Due date"),
        F(type=T.FLOAT, key="questionsTotal", display_name="Questions"),
        F(type=T.FLOAT, key="autoAnswered", display_name="Auto-answered %"),
        F(type=T.FLOAT, key="flaggedCount", display_name="Flagged"),
        F(type=T.STRING, key="reviewer", display_name="Reviewer"),
    ]
    try:
        client.metadata_templates.create_metadata_template(
            "enterprise", "Security Review", template_key=TEMPLATE_KEY, fields=fields
        )
        log("template", "created securityReview metadata template")
    except BoxAPIError as e:
        code = e.response_info.status_code
        if code == 409:
            log("template", "securityReview template already exists")
        else:
            log("template", f"could not create it via API ({code}). Create it in Admin Console > Content > "
                            "Metadata with key 'securityReview' and the fields listed in the README.")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--webhook", help="public URL ending in /box/webhook")
    a = p.parse_args()
    client = get_client()
    make_content()

    trust = env("TRUST_FOLDER_ID", required=True)
    deals = env("DEALS_FOLDER_ID", required=True)
    metadata_template(client)

    for name, _title, review_by, _ in TRUST_DOCS:
        upload(client, trust, OUT / "01 Trust Center Content" / name, description=f"review-by: {review_by}")

    gated = folder(client, trust, "NDA-gated (not used by AI)")
    soc2 = upload(client, gated, OUT / "Legal" / "SOC 2 Type II Report (SAMPLE).docx")
    nda = upload(client, gated, OUT / "Legal" / "Mutual NDA (SAMPLE).docx")

    contoso = folder(client, folder(client, deals, "Contoso Bank"), "Security Review")
    folder(client, folder(client, deals, "Fabrikam Health"), "Security Review")

    if a.webhook:
        hook = client.webhooks.create_webhook(
            CreateWebhookTarget(id=deals, type=CreateWebhookTargetTypeField.FOLDER),
            a.webhook,
            [CreateWebhookTriggers.FILE_UPLOADED, CreateWebhookTriggers.TASK_ASSIGNMENT_UPDATED,
             CreateWebhookTriggers.SIGN_REQUEST_COMPLETED],
        )
        log("webhook", f"registered {hook.id} -> {a.webhook}")

    print("\nAdd these to your .env:")
    print(f"SOC2_FILE_ID={soc2}\nNDA_FILE_ID={nda}")
    print(f"\nLive demo: upload demo_content/Prospect/Contoso_Vendor_Security_Questionnaire.xlsx "
          f"to Deals / Contoso Bank / Security Review (folder {contoso}).")


if __name__ == "__main__":
    main()
