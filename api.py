# api.py
import requests
from typing import Dict, List, Optional
from datetime import datetime, timedelta, timezone

from fastapi import FastAPI, HTTPException, Header
from fastapi.middleware.cors import CORSMiddleware

GRAPH = "https://graph.microsoft.com/v1.0"

app = FastAPI()

# CORS: allow your Netlify site (and localhost while developing)
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://godot-outlook-app.netlify.app",
        "https://localhost:3000",  # optional for local dev
    ],
    allow_methods=["*"],
    allow_headers=["*"],
)

def graph_get(url: str, user_token: str, params: Optional[Dict] = None) -> Dict:
    r = requests.get(url, headers={"Authorization": f"Bearer {user_token}"}, params=params, timeout=30)
    if r.status_code >= 400:
        raise HTTPException(status_code=r.status_code, detail=r.text)
    return r.json()

def list_all_pages(url: str, user_token: str, params: Optional[Dict] = None) -> List[Dict]:
    data = graph_get(url, user_token, params)
    items = data.get("value", [])
    while "@odata.nextLink" in data:
        data = graph_get(data["@odata.nextLink"], user_token, None)
        items.extend(data.get("value", []))
    return items

def iso_utc(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")

@app.get("/unreplied")
def unreplied(days: int = 30, authorization: str = Header(...)) -> List[Dict]:
    if not authorization.lower().startswith("bearer "):
        raise HTTPException(status_code=401, detail="Missing bearer token")
    user_token = authorization.split(" ", 1)[1]

    since = datetime.now(timezone.utc) - timedelta(days=days)
    since_iso = iso_utc(since)

    # 1) recent sent items
    sent_url = f"{GRAPH}/me/mailFolders/SentItems/messages"
    sent_params = {
        "$filter": f"sentDateTime ge {since_iso}",
        "$select": "id,subject,conversationId,toRecipients,sentDateTime",
        "$top": 100,
        "$orderby": "sentDateTime desc",
    }
    sent_items = list_all_pages(sent_url, user_token, sent_params)

    results: List[Dict] = []
    for m in sent_items:
        cid = m.get("conversationId"); sent_time = m.get("sentDateTime")
        if not cid or not sent_time: continue

        # 2) replies in Inbox for that conversation (no date filter here; we filter locally)
        inbox_url = f"{GRAPH}/me/mailFolders/Inbox/messages"
        inbox_params = {
            "$filter": f"conversationId eq '{cid}'",
            "$select": "id,from,receivedDateTime",
            "$top": 50,
        }
        inbox = graph_get(inbox_url, user_token, inbox_params)
        inbox_msgs = inbox.get("value", [])

        sent_dt = datetime.strptime(sent_time, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
        has_reply = any(
            (im.get("receivedDateTime") and
             datetime.strptime(im["receivedDateTime"], "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc) > sent_dt)
            for im in inbox_msgs
        )

        if not has_reply:
            to_list = m.get("toRecipients") or []
            to_first = (to_list[0].get("emailAddress", {}).get("address") if to_list else "") or ""
            results.append({
                "subject": m.get("subject") or "(no subject)",
                "to": to_first,
                "sent": sent_time
            })

    results.sort(key=lambda x: x["sent"] or "", reverse=True)
    return results
