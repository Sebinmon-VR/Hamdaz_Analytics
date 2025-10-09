import os
from flask import Flask, redirect, url_for, render_template, session, request, jsonify
import requests
from datetime import datetime
import pytz
from io import BytesIO
from apscheduler.schedulers.background import BackgroundScheduler
import pandas as pd

from auth import login_redirect, fetch_tokens, get_graph_headers
from functions import *  # Your existing SharePoint/Excel helper functions

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "super_secret_key")

GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"
SITE_NAME = os.getenv("SITE_NAME", "ProposalTeam")
LIST_NAME = os.getenv("LIST_NAME", "Proposals")
EXCEL_FILE_NAME = "UserAnalytics.xlsx"

# ---------------------------------------------------------
# EXCLUDED USERS
# ---------------------------------------------------------
EXCLUDED_USERS = ["Sebin", "Shamshad", "Jaymon", "Hisham Arackal", "Althaf", "Nidal", "Nayif Muhammed S"]

# ---------------------------------------------------------
# HELPER FUNCTIONS
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
# USER PRIORITY CALCULATION
# ---------------------------------------------------------
def compute_user_priority(df):
    import pytz
    from datetime import datetime, timedelta

    uae_tz = pytz.timezone("Asia/Dubai")
    now_uae = datetime.now(uae_tz)

    if df.empty or 'AssignedTo' not in df.columns:
        return {}

    # Detect Start Date column
    start_col = None
    for col in df.columns:
        if col.lower().replace(" ", "") == "startdate":
            start_col = col
            break

    # Convert date columns to datetime
    if 'BCD' in df.columns:
        df['BCD'] = pd.to_datetime(df['BCD'], errors='coerce', utc=True).dt.tz_convert(uae_tz)
    if start_col:
        df[start_col] = pd.to_datetime(df[start_col], errors='coerce', utc=True).dt.tz_convert(uae_tz)

    # Exclude specific users
    df = df[~df['AssignedTo'].isin(EXCLUDED_USERS)]
    if df.empty:
        return {}

    # Compute active tasks per user
    active_tasks = df[(df['SubmissionStatus'] != 'Submitted') & (df['BCD'] >= now_uae)].groupby('AssignedTo').size()

    # Compute last assigned date per user
    if start_col:
        last_assigned_dates = df.groupby('AssignedTo')[start_col].max()
    else:
        last_assigned_dates = pd.Series({user: now_uae - timedelta(days=3) for user in df['AssignedTo'].unique()})

    # Build user list
    user_list = []
    for user in df['AssignedTo'].unique():
        count = active_tasks.get(user, 0)
        last_date = last_assigned_dates.get(user, now_uae - timedelta(days=3))
        days_since_last = (now_uae - last_date).days
        user_list.append({
            "user": user,
            "active_tasks": count,
            "days_since_last": days_since_last
        })

    # Sort and assign priority
    user_list.sort(key=lambda x: (x['active_tasks'], -x['days_since_last']))
    priorities = {u['user']: idx+1 for idx, u in enumerate(user_list)}
    return priorities

# ---------------------------------------------------------
# USER ANALYTICS
# ---------------------------------------------------------
def compute_user_analytics_with_last_date(df):
    if df.empty or 'AssignedTo' not in df.columns:
        return {}

    # Exclude specific users
    df = df[~df['AssignedTo'].isin(EXCLUDED_USERS)]
    if df.empty:
        return {}

    uae_tz = pytz.timezone("Asia/Dubai")
    now_uae = datetime.now(uae_tz)

    # Detect Start Date column
    start_col = None
    for col in df.columns:
        if col.lower().replace(" ", "") == "startdate":
            start_col = col
            break

    if 'BCD' in df.columns:
        df['BCD'] = pd.to_datetime(df['BCD'], errors='coerce', utc=True).dt.tz_convert(uae_tz)
    if start_col:
        df[start_col] = pd.to_datetime(df[start_col], errors='coerce', utc=True).dt.tz_convert(uae_tz)

    analytics = {}
    for user, user_df in df.groupby('AssignedTo'):
        last_assigned_date = user_df[start_col].max() if start_col else None
        if pd.isna(last_assigned_date):
            last_assigned_date_str = None
        else:
            last_assigned_date_str = last_assigned_date.strftime("%Y-%m-%d %H:%M")

        analytics[user] = {
            "total_tasks": len(user_df),
            "tasks_completed": len(user_df[user_df['SubmissionStatus']=='Submitted']),
            "tasks_pending": len(user_df[(user_df['SubmissionStatus']!='Submitted') & (user_df['BCD']>=now_uae)]),
            "tasks_missed": len(user_df[(user_df['SubmissionStatus']!='Submitted') & (user_df['BCD']<now_uae)]),
            "orders_received": len(user_df[user_df['Status']=='Received']) if 'Status' in df.columns else 0,
            "last_assigned_date": last_assigned_date_str
        }
    return analytics

# ---------------------------------------------------------
# EXCEL FUNCTIONS
# ---------------------------------------------------------
def ensure_excel_file():
    headers = get_graph_headers()
    check_url = f"{GRAPH_API_ENDPOINT}/me/drive/root:/{EXCEL_FILE_NAME}"
    r = requests.get(check_url, headers=headers)
    if r.status_code == 404:
        print("📁 Creating new UserAnalytics.xlsx in OneDrive root...")
        excel_data = BytesIO()
        pd.DataFrame().to_excel(excel_data, index=False)
        excel_data.seek(0)
        create_url = f"{GRAPH_API_ENDPOINT}/me/drive/root:/{EXCEL_FILE_NAME}:/content"
        resp = requests.put(create_url, headers=headers, data=excel_data.read())
        if resp.status_code in [200, 201]:
            print("✅ Created Excel file successfully.")
        else:
            print("❌ Failed to create Excel file:", resp.text)
    else:
        print("✅ Excel file exists in OneDrive root.")
def update_user_analytics_excel(per_user, priorities=None):
    headers = get_graph_headers()
    ensure_excel_file()

    # Filter out excluded users
    filtered_per_user = {user: data for user, data in per_user.items() if user not in EXCLUDED_USERS}
    if not filtered_per_user:
        print("⚠️ No users to update in Excel.")
        return

    # Create DataFrame
    df_per_user = pd.DataFrame.from_dict(filtered_per_user, orient="index").reset_index()
    df_per_user.rename(columns={"index": "User"}, inplace=True)

    # Add Priority if provided
    if priorities:
        filtered_priorities = {user: prio for user, prio in priorities.items() if user not in EXCLUDED_USERS}
        df_per_user["Priority"] = df_per_user["User"].map(filtered_priorities)

    # Ensure proper column order
    columns_order = ["Priority", "User", "total_tasks", "tasks_completed", "tasks_pending",
                     "tasks_missed", "orders_received", "last_assigned_date"]
    df_per_user = df_per_user[[col for col in columns_order if col in df_per_user.columns]]

    # Convert DataFrame to Excel with table formatting
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.worksheet.table import Table, TableStyleInfo

    wb = Workbook()
    ws = wb.active
    ws.title = "UserAnalytics"

    # Write data rows
    for r in dataframe_to_rows(df_per_user, index=False, header=True):
        ws.append(r)

    # Create Excel table
    tab = Table(displayName="UserAnalyticsTable", ref=f"A1:{chr(64+len(df_per_user.columns))}{len(df_per_user)+1}")
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    tab.tableStyleInfo = style
    ws.add_table(tab)

    # Save to BytesIO
    excel_data = BytesIO()
    wb.save(excel_data)
    excel_data.seek(0)

    # Upload to OneDrive
    upload_url = f"{GRAPH_API_ENDPOINT}/me/drive/root:/{EXCEL_FILE_NAME}:/content"
    response = requests.put(upload_url, headers=headers, data=excel_data.read())
    if response.status_code in [200, 201]:
        print("✅ User analytics Excel updated successfully as a table.")
    else:
        print("❌ Failed to update Excel file:", response.text)


# ---------------------------------------------------------
# BACKGROUND SCHEDULER
# ---------------------------------------------------------
scheduler = BackgroundScheduler()
scheduler.start()

def background_analytics_job():
    try:
        structured_items = get_sharepoint_list_data(SITE_NAME, LIST_NAME)
        df = sharepoint_data_to_df(structured_items)
        per_user = compute_user_analytics_with_last_date(df)
        priorities = compute_user_priority(df)
        update_user_analytics_excel(per_user, priorities)
        print(f"[{datetime.now()}] ✅ Analytics and priorities updated.")
    except Exception as e:
        print(f"[{datetime.now()}] ❌ Error in background job: {e}")

# Run every 5 minutes
scheduler.add_job(background_analytics_job, 'interval', minutes=5)

# ---------------------------------------------------------
# FLASK ROUTES
# ---------------------------------------------------------
@app.route("/")
def index():
    headers = get_graph_headers()
    if headers:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))

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

@app.route("/dashboard")
def dashboard():
    structured_items = get_sharepoint_list_data(SITE_NAME, LIST_NAME)
    df = sharepoint_data_to_df(structured_items)
    per_user = compute_user_analytics_with_last_date(df)
    priorities = compute_user_priority(df)
    update_user_analytics_excel(per_user, priorities)
    overall = compute_overall_analytics(df)

    user_info = session.get("user_info", {})
    greeting = get_greeting()
    access_token = session.get("access_token")
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

@app.route("/teams")
def teams():
    sp_items = get_sharepoint_list_data(SITE_NAME, LIST_NAME)
    analytics = compute_teams_analytics(sp_items)
    user_info = session.get("user_info", {})
    user=user_info
    users = list(analytics.get("users", {}).keys())
    return render_template("teams.html", analytics=analytics, users=users , user=user)

@app.route("/user/<username>")
def user_analytics(username):
    if username.lower() == "dashboard":
        return redirect(url_for("dashboard"))
    sp_items = get_sharepoint_list_data(SITE_NAME, LIST_NAME)
    analytics = compute_user_analytics_specific(sp_items, username)
    return render_template("users_analytics.html", username=username, analytics=analytics)

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
        {"name": f.get("name"), "path": f.get("parentReference", {}).get("path", "/") + "/" + f.get("name")}
        for f in files
    ]
    return render_template("files.html", files=files_with_paths)

@app.route("/profile")
def profile():
    headers = get_graph_headers()
    if not headers:
        return jsonify({"error": "User not authenticated"}), 401
    response = requests.get(f"{GRAPH_API_ENDPOINT}/me", headers=headers)
    if response.status_code == 200:
        return jsonify(response.json())
    return jsonify({"error": "Failed to fetch profile", "details": response.json()}), response.status_code

@app.route("/users-photos")
def users_photos():
    users = get_users_with_photos()
    return render_template("users_photos.html", users=users)

@app.route("/proposals")
def proposals():
    items = get_sharepoint_list_data(SITE_NAME, LIST_NAME)
    columns = list(items[0].keys()) if items else []
    return render_template("proposals.html", items=items, columns=columns)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))



# =========================================================
# BUSINESS CARD PAGE ROUTES
# =========================================================

# This route now displays the REAL data from Excel
@app.route('/businesscards')
def business_cards_page():
    # We need user info for the navbar, so we'll check if the user is logged in
    if not session.get("user_info"):
        # If not logged in, we can't show the page. Redirect to login.
        return redirect(url_for("login"))
        
    # Import the REAL data function with your new name
    from functions import get_contacts_from_excel
    
    contacts = get_contacts_from_excel()

    # Get the user info needed to render the navbar correctly
    user_info = session.get("user_info", {})
    access_token = session.get("access_token")
    picture = get_profile_picture(access_token)
    org_name = get_graph_data(f"{GRAPH_API_ENDPOINT}/organization", access_token)["value"][0]["displayName"] if access_token else "Org"

    return render_template(
        'business_cards.html', 
        contacts=contacts, 
        user=user_info, 
        picture=picture, 
        org_name=org_name
    )

# This route now saves the edited data to Excel
@app.route('/update-contact', methods=['POST'])
def update_contact():
    # Import the REAL update function with your new name
    from functions import update_contact_in_excel
    
    edited_data = request.json
    print("Received edited data from UI:", edited_data)
    
    row_index = edited_data.get('id')
    
    # We need to convert the ID from the webpage (which is a string) to an integer
    try:
        row_index = int(row_index)
    except (ValueError, TypeError):
        return jsonify({"status": "error", "message": "Invalid row ID."}), 400

    success = update_contact_in_excel(row_index, edited_data)
    
    if success:
        return jsonify({"status": "success", "message": "Update successful!"})
    else:
        return jsonify({"status": "error", "message": "Failed to update Excel file."}), 500



# ---------------------------------------------------------
if __name__ == "__main__":
    app.run(debug=True)
