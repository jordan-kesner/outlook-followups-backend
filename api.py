# api.py
import os
import requests
import msal
from typing import Dict, List, Optional
from datetime import datetime, timedelta, timezone

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware

# ================== CONFIG ==================
TENANT_ID = "fe62ff8e-1750-452e-b2ff-2d788a3db229"
CLIENT_ID = "98aa19bc-8efd-4a09-8ae8-23a2ad3858d4"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/Mail.Read"]
GRAPH = "https://graph.microsoft.com/v1.0"

# ================== FASTAPI ==================
app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://localhost:3000", "http://localhost:3000"],  # webpack dev server
    allow_methods=["*"],
    allow_headers=["*"],
)

# ================== MSAL (interactive once, silent thereafter) ==================
msal_app = msal.PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY)

def get_access_token() -> str:
    """Acquire a delegated token via MSAL (interactive first time, silent afterwards)."""
    accounts = msal_app.get_accounts()
    result = msal_app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
    if not result:
        # Interactive prompt on first run
        result = msal_app.acquire_token_interactive(scopes=SCOPES)
    if "access_token" not in result:
        raise HTTPException(status_code=401, detail=result.get("error_description", "Auth failed"))
    return result["access_token"]

# ================== GRAPH HELPERS ==================
def graph_get(url: str, token: str, params: Optional[Dict] = None) -> Dict:
    """GET helper with error visibility."""
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, params=params)
    if r.status_code >= 400:
        print("GRAPH ERROR", r.status_code, r.text)  # log full Graph error for debugging
        raise HTTPException(status_code=r.status_code, detail=r.text)
    return r.json()

def list_all_pages(url: str, token: str, params: Optional[Dict] = None) -> List[Dict]:
    """Follow @odata.nextLink to collect all pages."""
    data = graph_get(url, token, params)
    items = data.get("value", [])
    while "@odata.nextLink" in data:
        data = graph_get(data["@odata.nextLink"], token, None)
        items.extend(data.get("value", []))
    return items

def iso_utc(dt: datetime) -> str:
    """Format datetime as Graph-friendly UTC ISO with 'Z' (no microseconds)."""
    return dt.astimezone(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")

def parse_iso_z(s: str) -> datetime:
    """Parse Graph-style 'YYYY-MM-DDTHH:MM:SSZ' into aware UTC datetime."""
    return datetime.strptime(s, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)

# ================== ENDPOINT ==================
@app.get("/unreplied")
def unreplied(days: int = 30, top: int = 100) -> List[Dict]:
    """
    Returns sent emails in the last `days` that have no reply in Inbox.
    Response shape (array of objects):
      { "subject": str, "to": str, "sent": ISO8601 }
    """
    token = get_access_token()

    since = datetime.now(timezone.utc) - timedelta(days=days)
    since_iso = iso_utc(since)

    # 1) Get recent Sent Items (select only needed fields)
    sent_url = f"{GRAPH}/me/mailFolders/SentItems/messages"
    sent_params = {
        "$filter": f"sentDateTime ge {since_iso}",  # DateTime unquoted
        "$select": "id,subject,conversationId,toRecipients,sentDateTime,internetMessageId",
        "$top": max(1, min(top, 1000)),
        "$orderby": "sentDateTime desc",
    }
    sent_items = list_all_pages(sent_url, token, sent_params)

    results: List[Dict] = []

    # 2) For each sent message, fetch Inbox messages by conversationId ONLY
    #    (Graph 400 'InefficientFilter' happens if we also filter by date server-side)
    for m in sent_items:
        cid = m.get("conversationId")
        sent_time = m.get("sentDateTime")
        if not cid or not sent_time:
            continue

        inbox_url = f"{GRAPH}/me/mailFolders/Inbox/messages"
        inbox_params = {
            "$filter": f"conversationId eq '{cid}'",  # quote the conversationId
            "$select": "id,from,receivedDateTime,internetMessageId",
            "$top": 50,
            # "$orderby": "receivedDateTime desc",  # uncomment if needed; remove if it causes 400
        }
        inbox_page = graph_get(inbox_url, token, inbox_params)
        inbox_msgs = inbox_page.get("value", [])

        # Local time comparison to decide if there's a reply AFTER the sent time
        sent_dt = parse_iso_z(sent_time)
        has_reply = False
        for im in inbox_msgs:
            rcv = im.get("receivedDateTime")
            if rcv:
                try:
                    if parse_iso_z(rcv) > sent_dt:
                        has_reply = True
                        break
                except ValueError:
                    # If a different ISO format ever appears, skip that item
                    continue

        if not has_reply:
            to_list = m.get("toRecipients") or []
            to_first = (to_list[0].get("emailAddress", {}).get("address") if to_list else "") or ""
            results.append({
                "subject": m.get("subject") or "(no subject)",
                "to": to_first,
                "sent": sent_time
            })

    # Sort newest first
    results.sort(key=lambda x: x["sent"] or "", reverse=True)
    return results
