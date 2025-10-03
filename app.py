from flask import Flask, redirect, url_for, render_template, jsonify, request, session
import os
import requests
from auth import *
from functions import *

app = Flask(__name__)
from dotenv import load_dotenv

load_dotenv(override=True)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "super_secret_key")

# ----------------- CONFIG -----------------
# Local dev: http://localhost:5000
# Production: https://your-render-app.onrender.com
BASE_URL = os.getenv("BASE_URL", "http://localhost:5000")

GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"
# -------------------------------------------

# ----------------- ROUTES ------------------

@app.route("/")
def index():
    access_token = get_graph_headers()
    if access_token:
        return redirect(f"{BASE_URL}/files")
    return redirect(f"{BASE_URL}/login")

@app.route("/login")
def login():
    return login_redirect()

@app.route("/callback")
def callback():
    code = request.args.get("code")
    if not code:
        return "Error: No code returned", 400

    if fetch_tokens(code):
        return redirect(f"{BASE_URL}/dashboard")
    return "Error fetching tokens", 400

@app.route("/files")
def files():
    headers = get_graph_headers()
    if not headers:
        return redirect(f"{BASE_URL}/login")

    # Get all files in root
    response = requests.get(f"{GRAPH_API_ENDPOINT}/me/drive/root/children", headers=headers)

    if response.status_code == 401:
        headers = get_graph_headers()  # refresh token
        response = requests.get(f"{GRAPH_API_ENDPOINT}/me/drive/root/children", headers=headers)

    if response.status_code != 200:
        return f"Error fetching files: {response.json()}", response.status_code

    files = response.json().get("value", [])
    files_with_paths = [{"name": f.get("name"), "path": f.get("parentReference", {}).get("path", "/") + "/" + f.get("name")} for f in files]

    return render_template("files.html", files=files_with_paths)

@app.route("/excel-data")
def excel_data():
    user_id = os.getenv("user_id")  # or use get_my_user_id()
    if not user_id:
        return "Error: Cannot get user ID", 400

    file_path = f"/users/{user_id}/drive/root:/Sharepoint Datas.xlsx"
    tables = get_excel_tables(file_path)
    table_data = {table.get("name"): get_table_data(file_path, table.get("name")) for table in tables}
    return jsonify(table_data)

@app.route("/proposals")
def proposals():
    site_name = os.getenv("SITE_NAME")
    list_name = os.getenv("LIST_NAME")

    items = get_sharepoint_list_data(site_name, list_name)
    if not items:
        return "No items found or unable to fetch list", 400

    columns = list(items[0].keys()) if items else []
    return render_template("proposals.html", items=items, columns=columns)

@app.route("/dashboard")
def dashboard():
    structured_items = get_sharepoint_list_data("ProposalTeam", "Proposals")
    df = sharepoint_data_to_df(structured_items)

    overall = compute_overall_analytics(df)
    per_user = compute_user_analytics(df)

    return render_template("dashboard.html", overall=overall, per_user=per_user)

@app.route("/teams")
def teams():
    sp_items = get_sharepoint_list_data("ProposalTeam", "Proposals")
    analytics = compute_teams_analytics(sp_items)
    users = list(analytics["users"].keys())
    return render_template("teams.html", analytics=analytics, users=users)

@app.route("/analytics")
def analytics():
    user_id = get_my_user_id()
    if not user_id:
        return "Error: Cannot get user ID", 400

    file_path = f"/users/{user_id}/drive/root:/Sharepoint Datas.xlsx"
    data = get_users_analytics(file_path)
    return jsonify(data)

@app.route("/user/<username>")
def user_analytics(username):
    if username == "dashboard":
        return redirect(f"{BASE_URL}/dashboard")

    sp_items = get_sharepoint_list_data("ProposalTeam", "Proposals")
    user_analytics_data = compute_user_analytics_specific(sp_items, username)

    return render_template(
        "users_analytics.html",
        username=username,
        analytics=user_analytics_data
    )

@app.route("/logout")
def logout():
    session.clear()
    return redirect(f"{BASE_URL}/")

# ----------------- MAIN --------------------
if __name__ == "__main__":
    # Local development only
    app.run(debug=True, host="0.0.0.0", port=5000)
