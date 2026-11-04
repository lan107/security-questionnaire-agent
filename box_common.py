"""Shared helpers: configuration, the authenticated Box client, and logging."""
import json
import os
import sys
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from box_sdk_gen import BoxCCGAuth, BoxClient, BoxDeveloperTokenAuth, CCGConfig

ROOT = Path(__file__).parent
load_dotenv(ROOT / ".env")

STATE_FILE = ROOT / "state.json"
TEMPLATE_KEY = "securityReview"


def env(name, default=None, required=False):
    value = os.getenv(name, default)
    if required and not value:
        sys.exit(f"Missing {name} in .env (see .env.example)")
    return value


def get_client() -> BoxClient:
    """Client Credentials Grant by default; a developer token for quick tests.

    With BOX_USER_ID set, the app acts as that user (files show as created by
    them). Without it, the app acts as its own service account, which must be
    invited as an Editor on the demo folders.
    """
    token = env("BOX_DEVELOPER_TOKEN")
    if token:
        return BoxClient(auth=BoxDeveloperTokenAuth(token=token))
    user_id = env("BOX_USER_ID")
    config = CCGConfig(
        client_id=env("BOX_CLIENT_ID", required=True),
        client_secret=env("BOX_CLIENT_SECRET", required=True),
        enterprise_id=None if user_id else env("BOX_ENTERPRISE_ID", required=True),
        user_id=user_id or None,
    )
    return BoxClient(auth=BoxCCGAuth(config=config))


def log(step, message):
    """One readable line per action, so the terminal tells the story in a demo."""
    print(f"{datetime.now():%H:%M:%S}  {step:<10} {message}", flush=True)


def load_state():
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text())
    return {"processed": {}, "drafts": {}, "sign_requests": {}}


def save_state(state):
    STATE_FILE.write_text(json.dumps(state, indent=2))
