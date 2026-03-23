import os
from dotenv import load_dotenv
load_dotenv()

APPLICATION_CLIENT_ID = os.getenv("APPLICATION_CLIENT_ID")
DIRECTORY_TENANT_ID = os.getenv("DIRECTORY_TENANT_ID")
AUTHORITY = f"https://login.microsoftonline.com/{DIRECTORY_TENANT_ID}"

# The permission scope we requested on Azure
SCOPES = ["Mail.Read"]

# Token cache file path
TOKEN_CACHE_PATH = "token_cache.json"