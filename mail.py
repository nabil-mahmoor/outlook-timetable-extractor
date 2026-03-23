import requests
from config import SCOPES, GRAPH_BASE_URL

def get_latest_timetable_email(token)-> dict | None:
    headers = {
        "Authorization": f"Bearer {token}",
        "ConsistencyLevel": "eventual"
    }
    
    params = {
        "$filter": "contains(subject, 'Timetable')",
        "$top": 10,
        "$select": "id, subject, receivedDateTime"
    }
    
    response = requests.get(
        f"{GRAPH_BASE_URL}/me/messages",
        headers=headers,
        params=params
    )
    
    response.raise_for_status()
    messages = response.json().get("value", [])

    if not messages:
        return None
    
    # Sort messages by receivedDateTime
    messages.sort(key=lambda m: m["receivedDateTime"], reverse=True)
    
    return messages[0]

def get_pdf_attachment(token, message_id):
    headers = {
        "Authorization": f"Bearer {token}"
    }
    
    response = requests.get(
        f"{GRAPH_BASE_URL}/me/messages/{message_id}/attachments",
        headers=headers
    )
    
    response.raise_for_status()
    attachments = response.json().get("value", [])

    for attachment in attachments:
        if attachment.get("contentType") == "application/pdf":
            return attachment
        
    return None

def download_pdf(attachment, timestamp=""):
    import base64
    
    download_path = f"timetable{"_" + timestamp}.pdf"
    
    pdf_data = base64.b64decode(attachment.get("contentBytes"))

    with open(download_path, "wb") as f:
        f.write(pdf_data)
    
    print(f"PDF downloaded to {download_path}")