import os
import requests
from flask import session, redirect
from urllib.parse import quote
from dotenv import load_dotenv

load_dotenv(override=True)

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
BASE_URL = os.getenv("BASE_URL", "http://localhost:5000")
REDIRECT_URI = f"{BASE_URL}/callback"
SCOPES = os.getenv("SCOPES", "https://graph.microsoft.com/.default offline_access")

AUTH_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

def login_redirect():
    auth_url = (
        f"{AUTH_URL}?client_id={CLIENT_ID}"
        f"&response_type=code"
        f"&redirect_uri={quote(REDIRECT_URI)}"
        f"&response_mode=query"
        f"&scope={quote(SCOPES)}"
    )
    return redirect(auth_url)

def fetch_tokens(code):
    token_data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "scope": SCOPES,
    }
    response = requests.post(TOKEN_URL, data=token_data).json()
    access_token = response.get("access_token")
    refresh_token = response.get("refresh_token")

    if access_token and refresh_token:
        session["access_token"] = access_token
        session["refresh_token"] = refresh_token
        return True
    return False

def refresh_access_token():
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
    response = requests.post(TOKEN_URL, data=data).json()
    access_token = response.get("access_token")
    new_refresh_token = response.get("refresh_token")

    if access_token:
        session["access_token"] = access_token
    if new_refresh_token:
        session["refresh_token"] = new_refresh_token

    return access_token

def get_graph_headers():
    access_token = session.get("access_token")
    if not access_token:
        access_token = refresh_access_token()
    if access_token:
        return {"Authorization": f"Bearer {access_token}"}
    return None
