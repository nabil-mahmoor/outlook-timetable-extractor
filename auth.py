import msal
import json
import os
from config import APPLICATION_CLIENT_ID, AUTHORITY, SCOPES, TOKEN_CACHE_PATH

def load_cache() -> msal.SerializableTokenCache:
    cache: msal.SerializableTokenCache = msal.SerializableTokenCache()

    if os.path.exists(TOKEN_CACHE_PATH):
        with open(TOKEN_CACHE_PATH, "r") as f:
            cache.deserialize(f.read())
    return cache

def save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        with open(TOKEN_CACHE_PATH, "w") as f:
            f.write(cache.serialize())
            
def get_access_token():
    cache: msal.SerializableTokenCache = load_cache()

    app = msal.PublicClientApplication(client_id=APPLICATION_CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    token = None
    
    # Attempt to access token from cache first
    accounts: list[msal.Account] = app.get_accounts()
    if accounts:
        token = app.acquire_token_silent(scopes=SCOPES, account=accounts[0])

    # If no cached token, do the device code login flow
    if not token:
        flow = app.initiate_device_flow(scopes=SCOPES)
        print(flow.get("message"))
        token = app.acquire_token_by_device_flow(flow)

    save_cache(cache)

    if "access_token" in token:
        return token.get("access_token")
    else:
        raise Exception(f"Failed to get token: {token.get("error_description")}")