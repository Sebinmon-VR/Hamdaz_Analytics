import os
import requests
from flask import session, redirect
from dotenv import load_dotenv

load_dotenv(override=True)

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
REDIRECT_URI = os.getenv("REDIRECT_URI")

# ✅ Always include essential Graph scopes here
SCOPES = os.getenv("SCOPES", "User.Read Files.Read Sites.Read.All offline_access")

GRAPH_API_ENDPOINT = os.getenv("GRAPH_API_ENDPOINT", "https://graph.microsoft.com/v1.0")

AUTH_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"


# ---------------------------------------------------------
# LOGIN REDIRECT
# ---------------------------------------------------------
def login_redirect():
    """Redirect user to Microsoft login page"""
    auth_url = (
        f"{AUTH_URL}?client_id={CLIENT_ID}"
        f"&response_type=code"
        f"&redirect_uri={REDIRECT_URI}"
        f"&response_mode=query"
        f"&scope={SCOPES}"
    )
    return redirect(auth_url)


# ---------------------------------------------------------
# FETCH TOKENS
# ---------------------------------------------------------
def fetch_tokens(code):
    """Exchange authorization code for access & refresh tokens"""
    token_data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "scope": SCOPES,
    }

    response = requests.post(TOKEN_URL, data=token_data)
    if response.status_code != 200:
        print("❌ Token fetch failed:", response.text)
        return False

    tokens = response.json()
    access_token = tokens.get("access_token")
    refresh_token = tokens.get("refresh_token")

    if not access_token:
        print("❌ Missing access token:", tokens)
        return False

    # ✅ Store tokens safely
    session["access_token"] = access_token
    session["refresh_token"] = refresh_token
    return True


# ---------------------------------------------------------
# REFRESH TOKEN
# ---------------------------------------------------------
def refresh_access_token():
    """Refresh expired access token using refresh token"""
    refresh_token = session.get("refresh_token")
    if not refresh_token:
        return None

    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "redirect_uri": REDIRECT_URI,
        "scope": SCOPES,
    }
    response = requests.post(TOKEN_URL, data=data)
    tokens = response.json()

    access_token = tokens.get("access_token")
    new_refresh_token = tokens.get("refresh_token")

    if access_token:
        session["access_token"] = access_token
    if new_refresh_token:
        session["refresh_token"] = new_refresh_token

    return access_token


# ---------------------------------------------------------
# GRAPH HEADERS
# ---------------------------------------------------------
def get_graph_headers():
    """Return headers with valid access token"""
    access_token = session.get("access_token")
    if not access_token:
        access_token = refresh_access_token()

    if access_token:
        return {"Authorization": f"Bearer {access_token}"}
    return None
