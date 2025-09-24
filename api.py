# api.py
import requests
import asyncio
import aiohttp
from typing import Dict, List, Optional
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, HTTPException, Header
from fastapi.middleware.cors import CORSMiddleware

GRAPH = "https://graph.microsoft.com/v1.0"

app = FastAPI()

# CORS: allow your Netlify site (and localhost while developing)
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://godot-outlook-app.netlify.app",
        "https://localhost:3000",  # for local dev
    ],
    allow_methods=["*"],
    allow_headers=["*"],
)

def graph_get(url: str, user_token: str, params: Optional[Dict] = None) -> Dict:
    # Add retry logic and better timeout handling
    max_retries = 3
    for attempt in range(max_retries):
        try:
            r = requests.get(
                url, 
                headers={"Authorization": f"Bearer {user_token}"}, 
                params=params, 
                timeout=15  # Reduced from 30 to 15 seconds
            )
            if r.status_code == 429:  # Rate limited
                # Wait and retry
                import time
                time.sleep(2 ** attempt)  # Exponential backoff
                continue
            elif r.status_code >= 400:
                raise HTTPException(status_code=r.status_code, detail=r.text)
            return r.json()
        except requests.exceptions.Timeout:
            if attempt == max_retries - 1:
                raise HTTPException(status_code=408, detail="Request timeout")
            continue
        except requests.exceptions.RequestException as e:
            if attempt == max_retries - 1:
                raise HTTPException(status_code=500, detail=f"Request failed: {str(e)}")
            continue
    
    raise HTTPException(status_code=500, detail="Max retries exceeded")

def list_all_pages(url: str, user_token: str, params: Optional[Dict] = None, max_pages: int = 5) -> List[Dict]:
    """List pages with a maximum limit to prevent excessive API calls"""
    data = graph_get(url, user_token, params)
    items = data.get("value", [])
    page_count = 1
    
    while "@odata.nextLink" in data and page_count < max_pages:
        data = graph_get(data["@odata.nextLink"], user_token, None)
        items.extend(data.get("value", []))
        page_count += 1
        
    return items

def iso_utc(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")

def is_calendar_invite_or_automated(message: Dict) -> bool:
    """Check if a message is a calendar invite or automated email"""
    subject = message.get("subject", "").lower()
    from_address = message.get("from", {}).get("emailAddress", {}).get("address", "").lower()
    
    # Calendar invite subjects
    calendar_subjects = [
        "calendar invitation", "meeting invitation", "accepted:", "declined:", 
        "tentative:", "meeting reminder", "zoom meeting invite", "invites you to a scheduled zoom meeting",
        "calendly", "meeting request", "meeting update", "meeting canceled"
    ]
    
    # Automated email addresses - be more specific to avoid false positives
    automated_senders = [
        "calendar-noreply@calendly.com", "no-reply@zoom.us", "noreply@microsoft.com",
        "calendar@outlook.com", "noreply@calendly.com", "no-reply@calendly.com",
        "noreply@teams.microsoft.com"
    ]
    
    # Check subject patterns
    if any(pattern in subject for pattern in calendar_subjects):
        print(f"Filtered calendar subject: {subject}")
        return True
    
    # Check sender patterns - be more specific
    for sender in automated_senders:
        if sender == from_address:  # Exact match instead of 'in'
            print(f"Filtered automated sender: {from_address}")
            return True
    
    # Only filter if sender starts with these specific patterns
    if from_address.startswith("calendar-server@") or from_address.startswith("noreply@"):
        print(f"Filtered noreply sender: {from_address}")
        return True
    
    return False

def is_auto_reply(message: Dict) -> bool:
    """Check if a message is an auto-reply"""
    subject = message.get("subject", "").lower()
    
    # Common auto-reply indicators in subject
    auto_reply_subjects = [
        "out of office", "ooo", "automatic reply", "auto-reply", "auto reply",
        "vacation", "away from office", "currently unavailable", "away message",
        "delivery status notification", "undeliverable", "delivery failure",
        "read receipt", "return receipt", "automated response"
    ]
    
    if any(indicator in subject for indicator in auto_reply_subjects):
        return True
    
    # Check if subject starts with common auto-reply prefixes
    auto_prefixes = ["re: out of", "automatic reply:", "auto-reply:", "away:"]
    if any(subject.startswith(prefix) for prefix in auto_prefixes):
        return True
        
    return False

@app.get("/unreplied")
def unreplied(days: int = 30, authorization: str = Header(...)) -> List[Dict]:
    if not authorization.lower().startswith("bearer "):
        raise HTTPException(status_code=401, detail="Missing bearer token")
    user_token = authorization.split(" ", 1)[1]

    since = datetime.now(timezone.utc) - timedelta(days=days)
    since_iso = iso_utc(since)

    # 1) Get sent items with improved filtering
    sent_url = f"{GRAPH}/me/mailFolders/SentItems/messages"
    sent_params = {
        "$filter": f"sentDateTime ge {since_iso}",
        "$select": "id,subject,conversationId,toRecipients,sentDateTime,from",
        "$top": 200,  # Increased to get more emails at once
        "$orderby": "sentDateTime desc",
    }
    sent_items = list_all_pages(sent_url, user_token, sent_params)

    # Filter out calendar invites and automated emails
    print(f"Total sent items before filtering: {len(sent_items)}")
    filtered_sent_items = []
    for item in sent_items:
        if not is_calendar_invite_or_automated(item):
            filtered_sent_items.append(item)
        else:
            print(f"Filtered out: {item.get('subject', 'No subject')} from {item.get('from', {}).get('emailAddress', {}).get('address', 'Unknown sender')}")
    
    print(f"Sent items after filtering: {len(filtered_sent_items)}")

    results: List[Dict] = []
    
    # Process in batches for better performance
    batch_size = 20
    for i in range(0, len(filtered_sent_items), batch_size):
        batch = filtered_sent_items[i:i + batch_size]
        
        for m in batch:
            cid = m.get("conversationId")
            sent_time = m.get("sentDateTime")
            if not cid or not sent_time:
                continue

            # 2) Check for replies in conversation, excluding auto-replies
            inbox_url = f"{GRAPH}/me/mailFolders/Inbox/messages"
            inbox_params = {
                "$filter": f"conversationId eq '{cid}'",
                "$select": "id,from,receivedDateTime,subject",
                "$orderby": "receivedDateTime desc",
                "$top": 20,  # Limit to recent messages in conversation
            }
            
            try:
                inbox = graph_get(inbox_url, user_token, inbox_params)
                inbox_msgs = inbox.get("value", [])

                sent_dt = datetime.strptime(sent_time, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
                
                # Check for real replies (not auto-replies)
                has_real_reply = False
                for im in inbox_msgs:
                    received_time = im.get("receivedDateTime")
                    if not received_time:
                        continue
                        
                    received_dt = datetime.strptime(received_time, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
                    
                    # Only consider messages received after we sent the email
                    if received_dt > sent_dt:
                        # Skip if it's an auto-reply
                        if not is_auto_reply(im):
                            has_real_reply = True
                            break

                if not has_real_reply:
                    to_list = m.get("toRecipients") or []
                    to_first = (to_list[0].get("emailAddress", {}).get("address") if to_list else "") or ""
                    email_result = {
                        "subject": m.get("subject") or "(no subject)",
                        "to": to_first,
                        "sent": sent_time
                    }
                    print(f"Adding unreplied email: {email_result['subject']} to {email_result['to']}")
                    results.append(email_result)
                else:
                    print(f"Email has reply, skipping: {m.get('subject', 'No subject')}")
                    
            except Exception as e:
                # Log error but continue processing other emails
                print(f"Error processing conversation {cid}: {e}")
                continue

    results.sort(key=lambda x: x["sent"] or "", reverse=True)
    return results
