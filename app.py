from flask import Flask, redirect, url_for, render_template ,jsonify
import requests
from auth import *
from functions import *
app = Flask(__name__)
app.secret_key = "super_secret_key"  # or use env variable

GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"
# --------------------------------------------------------------------------------------------

# --------------------------------------------------------------------------------------------

@app.route("/")
def index():
    access_token = get_graph_headers()
    if access_token:
        return redirect(url_for("files"))
    return redirect(url_for("login"))

@app.route("/login")
def login():
    return login_redirect()

@app.route("/callback")
def callback():
    from flask import request
    code = request.args.get("code")
    if not code:
        return "Error: No code returned", 400

    if fetch_tokens(code):
        return redirect(url_for("dashboard"))
    return "Error fetching tokens", 400

@app.route("/files")
def files():
    headers = get_graph_headers()
    if not headers:
        return redirect(url_for("login"))

    # Get all files in root
    response = requests.get(f"{GRAPH_API_ENDPOINT}/me/drive/root/children", headers=headers)
    
    if response.status_code == 401:
        # Try refresh token if unauthorized
        headers = get_graph_headers()
        response = requests.get(f"{GRAPH_API_ENDPOINT}/me/drive/root/children", headers=headers)

    if response.status_code != 200:
        return f"Error fetching files: {response.json()}", response.status_code

    files = response.json().get("value", [])

    # Extract name and path
    files_with_paths = []
    for f in files:
        files_with_paths.append({
            "name": f.get("name"),
            "path": f.get("parentReference", {}).get("path", "/") + "/" + f.get("name")
        })

    return render_template("files.html", files=files_with_paths)

@app.route("/excel-data")
def excel_data():
    user_id =os.getenv("user_id") # use get_my_user_id() to get the current logined user_id
    if not user_id:
        return "Error: Cannot get user ID", 400

    file_path = f"/users/{user_id}/drive/root:/Sharepoint Datas.xlsx"

    # Get tables in the file
    tables = get_excel_tables(file_path)

    table_data = {}
    for table in tables:
        table_name = table.get("name")
        rows = get_table_data(file_path, table_name)
        table_data[table_name] = rows

    return jsonify(table_data)



# @app.route("/sites")
# def list_sites():
#     headers = get_graph_headers()
#     if not headers:
#         return "No access token", 400

#     url = "https://graph.microsoft.com/v1.0/sites/root"  # current user's default site
#     response = requests.get(url, headers=headers)

#     if response.status_code != 200:
#         return f"Error fetching site: {response.json()}", response.status_code

#     site = response.json()
#     return render_template("sites.html", sites=[site])
@app.route("/proposals")
def proposals():
    site_name = os.getenv("SITE_NAME")  # Your SharePoint site name
    list_name =  os.getenv("LIST_NAME")    # Your SharePoint list name

    items = get_sharepoint_list_data(site_name, list_name)

    if not items:
        return "No items found or unable to fetch list", 400

    # Dynamically get all column names from the first item
    columns = list(items[0].keys()) if items else []

    return render_template("proposals.html", items=items, columns=columns)

# --------------------------------------------------------------------------------------------



@app.route("/dashboard")
def dashboard():
    structured_items = get_sharepoint_list_data("ProposalTeam", "Proposals")
    df = sharepoint_data_to_df(structured_items)

    overall = compute_overall_analytics(df)
    per_user = compute_user_analytics(df)

    return render_template("dashboard.html", overall=overall, per_user=per_user)


@app.route("/teams")
def teams():
    # Get SharePoint items
    sp_items = get_sharepoint_list_data("ProposalTeam", "Proposals")

    # Compute analytics
    analytics = compute_teams_analytics(sp_items)

    # Get list of users for tabs
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
    if username=="dashboard":
         return redirect(url_for("dashboard"))
    # Fetch SharePoint items
    sp_items = get_sharepoint_list_data("ProposalTeam", "Proposals")
    # Compute analytics
    user_analytics = compute_user_analytics_specific(sp_items, username)

    return render_template(
        "users_analytics.html",
        username=username,
        analytics=user_analytics
    )



# --------------------------------------------------------------------------------------------


@app.route("/logout")
def logout():
    from flask import session
    session.clear()
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(debug=True)
