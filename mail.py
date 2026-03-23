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