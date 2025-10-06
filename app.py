import os
from flask import Flask, redirect, url_for, render_template, session, request, jsonify
import requests
from datetime import datetime
import pytz
from collections import defaultdict

from auth import login_redirect, fetch_tokens, get_graph_headers
from functions import *

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "super_secret_key")

GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"
SITE_NAME = os.getenv("SITE_NAME", "ProposalTeam")
LIST_NAME = os.getenv("LIST_NAME", "Proposals")

# ---------------------------------------------------------
# INDEX
# ---------------------------------------------------------
@app.route("/")
def index():
    headers = get_graph_headers()
    if headers:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))

# ---------------------------------------------------------
# LOGIN & CALLBACK
# ---------------------------------------------------------
@app.route("/login")
def login():
    return login_redirect()

@app.route("/callback")
def callback():
    code = request.args.get("code")
    if not code:
        return "Error: No code returned", 400

    if fetch_tokens(code):
        return redirect(url_for("dashboard"))
    return "Error fetching tokens", 400

# ---------------------------------------------------------
# DASHBOARD
# ---------------------------------------------------------
@app.route("/dashboard")
def dashboard():
    structured_items = get_sharepoint_list_data(SITE_NAME, LIST_NAME)
    df = sharepoint_data_to_df(structured_items)
    overall = compute_overall_analytics(df)
    per_user = compute_user_analytics(df)

    user_info = session.get("user_info", {})
    greeting = get_greeting()
    access_token = session["access_token"]
    picture = get_profile_picture(access_token)

    return render_template(
        "dashboard.html",
        overall=overall,
        per_user=per_user,
        user=user_info,
        greeting=greeting, 
        picture=picture,
        org_name = get_graph_data(f"{GRAPH_API_ENDPOINT}/organization", access_token)["value"][0]["displayName"]


    )

# ---------------------------------------------------------
# TEAMS ANALYTICS
# ---------------------------------------------------------
@app.route("/teams")
def teams():
    sp_items = get_sharepoint_list_data(SITE_NAME, LIST_NAME)
    analytics = compute_teams_analytics(sp_items)
    users = list(analytics.get("users", {}).keys())
    return render_template("teams.html", analytics=analytics, users=users)

# ---------------------------------------------------------
# SPECIFIC USER ANALYTICS
# ---------------------------------------------------------
@app.route("/user/<username>")
def user_analytics(username):
    if username.lower() == "dashboard":
        return redirect(url_for("dashboard"))

    sp_items = get_sharepoint_list_data(SITE_NAME, LIST_NAME)
    analytics = compute_user_analytics_specific(sp_items, username)
    return render_template("users_analytics.html", username=username, analytics=analytics)

# ---------------------------------------------------------
# ONEDRIVE FILES
# ---------------------------------------------------------
@app.route("/files")
def files():
    headers = get_graph_headers()
    if not headers:
        return redirect(url_for("login"))

    response = requests.get(f"{GRAPH_API_ENDPOINT}/me/drive/root/children", headers=headers)
    if response.status_code == 401:
        headers = get_graph_headers()
        response = requests.get(f"{GRAPH_API_ENDPOINT}/me/drive/root/children", headers=headers)

    if response.status_code != 200:
        return f"Error fetching files: {response.json()}", response.status_code

    files = response.json().get("value", [])
    files_with_paths = [
        {
            "name": f.get("name"),
            "path": f.get("parentReference", {}).get("path", "/") + "/" + f.get("name")
        } for f in files
    ]

    return render_template("files.html", files=files_with_paths)



# ---------------------------------------------------------
# USER PROFILE
# ---------------------------------------------------------
@app.route("/profile")
def profile():
    headers = get_graph_headers()
    if not headers:
        return jsonify({"error": "User not authenticated"}), 401

    response = requests.get(f"{GRAPH_API_ENDPOINT}/me", headers=headers)
    if response.status_code == 200:
        return jsonify(response.json())
    return jsonify({"error": "Failed to fetch profile", "details": response.json()}), response.status_code

# ---------------------------------------------------------
# USERS WITH PHOTOS
# ---------------------------------------------------------
@app.route("/users-photos")
def users_photos():
    users = get_users_with_photos()
    return render_template("users_photos.html", users=users)

# ---------------------------------------------------------
# PROPOSALS LIST
# ---------------------------------------------------------
@app.route("/proposals")
def proposals():
    items = get_sharepoint_list_data(SITE_NAME, LIST_NAME)
    columns = list(items[0].keys()) if items else []
    return render_template("proposals.html", items=items, columns=columns)

# ---------------------------------------------------------
# LOGOUT
# ---------------------------------------------------------
@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))

# ---------------------------------------------------------
# GREETING FUNCTION
# ---------------------------------------------------------
def get_greeting():
    now = datetime.now()
    hour = now.hour
    if 5 <= hour < 12:
        return "Good Morning"
    elif 12 <= hour < 17:
        return "Good Afternoon"
    elif 17 <= hour < 21:
        return "Good Evening"
    else:
        return "Hello"

# ---------------------------------------------------------
if __name__ == "__main__":
    app.run(debug=True)
