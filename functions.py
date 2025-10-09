import os
import requests
import pandas as pd
from datetime import datetime
import pytz
from collections import defaultdict
from auth import get_graph_headers

GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"

def get_graph_data(endpoint, access_token):
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(endpoint, headers=headers)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Graph API error ({response.status_code}): {response.text}")
        return None

def get_my_user_id():
    """
    Returns the current logged-in user's ID from Microsoft Graph.
    """
    headers = get_graph_headers()
    if not headers:
        return None
    resp = requests.get(f"{GRAPH_API_ENDPOINT}/me", headers=headers)
    if resp.status_code == 200:
        return resp.json().get("id")
    return None
# ---------------------------------------------------------
# SHAREPOINT LIST FUNCTIONS
# ---------------------------------------------------------
def get_site_id(site_name):
    headers = get_graph_headers()
    url = f"{GRAPH_API_ENDPOINT}/sites/hamdaz1.sharepoint.com:/sites/{site_name}"
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        return resp.json().get("id")
    return None

def get_list_id(site_id, list_name):
    headers = get_graph_headers()
    url = f"{GRAPH_API_ENDPOINT}/sites/{site_id}/lists"
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        for l in resp.json().get("value", []):
            if l.get("name") == list_name:
                return l.get("id")
    return None

def get_list_items(site_id, list_id):
    headers = get_graph_headers()
    url = f"{GRAPH_API_ENDPOINT}/sites/{site_id}/lists/{list_id}/items?expand=fields($expand=AssignedTo,Author,Editor)"
    items = []
    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code != 200:
            break
        data = resp.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items

def flatten_fields(fields):
    flat = {}
    for k, v in fields.items():
        if isinstance(v, dict):
            flat[k] = v.get("displayName") or v.get("lookupValue") or str(v)
        elif isinstance(v, list):
            flat[k] = ', '.join([i.get('displayName', str(i)) if isinstance(i, dict) else str(i) for i in v])
        else:
            flat[k] = v
    return flat

def get_sharepoint_list_data(site_name, list_name):
    site_id = get_site_id(site_name)
    if not site_id:
        return []
    list_id = get_list_id(site_id, list_name)
    if not list_id:
        return []

    items = get_list_items(site_id, list_id)
    return [flatten_fields(item.get("fields", {})) for item in items]

# ---------------------------------------------------------
# SHAREPOINT DATA TO DF
# ---------------------------------------------------------
def sharepoint_data_to_df(structured_items):
    if not structured_items:
        return pd.DataFrame()
    df = pd.DataFrame(structured_items)
    required_cols = ["AssignedTo", "Priority", "Status", "SubmissionStatus", "BCD", "DueDate", "Title", "id"]
    for col in required_cols:
        if col not in df.columns:
            df[col] = None
    for date_col in ["BCD", "DueDate"]:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce', utc=True)
    return df

# ---------------------------------------------------------
# ANALYTICS FUNCTIONS
# ---------------------------------------------------------
def compute_overall_analytics(df):
    if df.empty:
        return {"total_users":0,"total_tasks":0,"tasks_completed":0,"tasks_pending":0,"tasks_missed":0,"orders_received":0}
    uae_tz = pytz.timezone("Asia/Dubai")
    now_uae = datetime.now(uae_tz)
    df['BCD'] = pd.to_datetime(df['BCD'], errors='coerce', utc=True).dt.tz_convert(uae_tz)
    total_users = df['AssignedTo'].nunique()
    total_tasks = len(df)
    tasks_completed = len(df[df['SubmissionStatus']=='Submitted'])
    tasks_pending = len(df[(df['SubmissionStatus']!='Submitted') & (df['BCD']>=now_uae)])
    tasks_missed = len(df[(df['SubmissionStatus']!='Submitted') & (df['BCD']<now_uae)])
    orders_received = len(df[df['Status']=='Received']) if 'Status' in df.columns else 0
    return {
        "total_users": total_users,
        "total_tasks": total_tasks,
        "tasks_completed": tasks_completed,
        "tasks_pending": tasks_pending,
        "tasks_missed": tasks_missed,
        "orders_received": orders_received
    }

def compute_user_analytics(df):
    if df.empty or 'AssignedTo' not in df.columns:
        return {}

    uae_tz = pytz.timezone("Asia/Dubai")
    now_uae = datetime.now(uae_tz)

    # Convert dates to timezone-aware
    if 'BCD' in df.columns:
        df['BCD'] = pd.to_datetime(df['BCD'], errors='coerce', utc=True).dt.tz_convert(uae_tz)
    if 'Start Date' in df.columns:
        df['Start Date'] = pd.to_datetime(df['Start Date'], errors='coerce', utc=True).dt.tz_convert(uae_tz)

    analytics = {}
    for user, user_df in df.groupby('AssignedTo'):
        # Last assigned date = latest Start Date
        last_assigned_date = None
        if 'Start Date' in user_df.columns and not user_df['Start Date'].isna().all():
            last_assigned_date = user_df['Start Date'].max()
            last_assigned_date = last_assigned_date.strftime("%Y-%m-%d %H:%M")

        analytics[user] = {
            "total_tasks": len(user_df),
            "tasks_completed": len(user_df[user_df['SubmissionStatus']=='Submitted']),
            "tasks_pending": len(user_df[(user_df['SubmissionStatus']!='Submitted') & (user_df['BCD']>=now_uae)]),
            "tasks_missed": len(user_df[(user_df['SubmissionStatus']!='Submitted') & (user_df['BCD']<now_uae)]),
            "orders_received": len(user_df[user_df['Status']=='Received']) if 'Status' in df.columns else 0,
            "last_assigned_date": last_assigned_date
        }

    return analytics


def compute_user_analytics_specific(sp_items, username):
    uae_tz = pytz.timezone("Asia/Dubai")
    now_uae = datetime.now(uae_tz)
    user_items = [i for i in sp_items if i.get("AssignedTo")==username]
    tasks= submissions= pending= missed =0
    for item in user_items:
        submission_status = (item.get("SubmissionStatus") or "").lower()
        bcd_str = item.get("BCD")
        try:
            bcd_dt = datetime.fromisoformat(bcd_str) if bcd_str else None
            if bcd_dt:
                bcd_dt = uae_tz.localize(bcd_dt)
        except: bcd_dt=None
        tasks +=1
        if submission_status=='submitted':
            submissions +=1
        elif bcd_dt and bcd_dt < now_uae:
            missed +=1
        else:
            pending +=1
    return {"tasks":tasks,"submissions":submissions,"pending":pending,"missed":missed}

def compute_teams_analytics(sp_items):
    uae_tz = pytz.timezone("Asia/Dubai")
    now = datetime.now(uae_tz)
    overall = {"total_tasks":0,"total_submissions":0,"total_pending":0,"total_missed":0,"users":{}}
    for item in sp_items:
        overall["total_tasks"] +=1
        user = item.get("AssignedTo") or "Unassigned"
        if user not in overall["users"]:
            overall["users"][user] = {"tasks":0,"submissions":0,"pending":0,"missed":0}
        overall["users"][user]["tasks"] +=1
        bcd_str = item.get("BCD")
        submission_status = (item.get("SubmissionStatus") or "").lower()
        try:
            bcd_dt = datetime.fromisoformat(bcd_str) if bcd_str else None
            if bcd_dt:
                bcd_dt = uae_tz.localize(bcd_dt)
        except: bcd_dt=None
        if submission_status=="submitted":
            overall["total_submissions"] +=1
            overall["users"][user]["submissions"] +=1
        if bcd_dt and bcd_dt >= now:
            overall["total_pending"] +=1
            overall["users"][user]["pending"] +=1
        if bcd_dt and bcd_dt < now and submission_status!="submitted":
            overall["total_missed"] +=1
            overall["users"][user]["missed"] +=1
    return overall

# ---------------------------------------------------------
# EXCEL / ONEDRIVE ANALYTICS
# ---------------------------------------------------------
def get_file_id(file_path):
    headers = get_graph_headers()
    url = f"{GRAPH_API_ENDPOINT}{file_path}"
    resp = requests.get(url, headers=headers)
    return resp.json().get("id") if resp.status_code==200 else None

def get_excel_tables(file_path):
    file_id = get_file_id(file_path)
    if not file_id: return []
    url = f"{GRAPH_API_ENDPOINT}/me/drive/items/{file_id}/workbook/tables"
    headers = get_graph_headers()
    resp = requests.get(url, headers=headers)
    return resp.json().get("value", []) if resp.status_code==200 else []

def get_table_data(file_path, table_name):
    file_id = get_file_id(file_path)
    if not file_id: return []
    url = f"{GRAPH_API_ENDPOINT}/me/drive/items/{file_id}/workbook/tables/{table_name}/rows"
    headers = get_graph_headers()
    resp = requests.get(url, headers=headers)
    return resp.json().get("value", []) if resp.status_code==200 else []

def get_users_analytics(file_path):
    tables = get_excel_tables(file_path)
    analytics = defaultdict(lambda: {"total_tasks":0,"active_tasks":0})
    today = datetime.now().date()
    for table in tables:
        rows = get_table_data(file_path, table.get("name"))
        for row in rows:
            values = row.get("values",[[]])[0]
            if len(values)<4: continue
            user, task_name, due_date, status = values[:4]
            analytics[user]["total_tasks"] +=1
            try:
                due_dt = datetime.strptime(due_date,"%Y-%m-%d").date()
                if due_dt>=today and status.lower()!="completed":
                    analytics[user]["active_tasks"] +=1
            except: pass
    return dict(analytics)

# ---------------------------------------------------------
# USERS WITH PHOTOS
# ---------------------------------------------------------
import base64
from flask import session

def get_users_with_photos():
    access_token = session.get("access_token")
    if not access_token: return []
    headers = {'Authorization':f'Bearer {access_token}'}
    resp = requests.get(f"{GRAPH_API_ENDPOINT}/users?$select=id,displayName,mail", headers=headers)
    if resp.status_code !=200: return []
    users = resp.json().get("value", [])
    for user in users:
        uid = user['id']
        photo_resp = requests.get(f"{GRAPH_API_ENDPOINT}/users/{uid}/photo/$value", headers=headers)
        if photo_resp.status_code==200:
            photo_b64 = base64.b64encode(photo_resp.content).decode('utf-8')
            user['photo'] = f"data:image/jpeg;base64,{photo_b64}"
        else:
            user['photo'] = None
    return users

def get_profile_picture(access_token, user_id=None):
    url = f"https://graph.microsoft.com/v1.0/me/photo/$value" if not user_id else f"https://graph.microsoft.com/v1.0/users/{user_id}/photo/$value"
    headers = {"Authorization": f"Bearer {access_token}"}
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        encoded = base64.b64encode(resp.content).decode("utf-8")
        return f"data:image/jpeg;base64,{encoded}"
    return "/static/default_profile.png"



# ---------------------------------------------------------
# EXCEL / ONEDRIVE READ & UPDATE FUNCTIONS
# ---------------------------------------------------------
def get_excel_file_id(file_path):
    """
    Get the file ID of an Excel file from OneDrive or SharePoint.
    """
    headers = get_graph_headers()
    url = f"{GRAPH_API_ENDPOINT}{file_path}"
    resp = requests.get(url, headers=headers)
    return resp.json().get("id") if resp.status_code == 200 else None

def get_excel_table_rows(file_path, table_name):
    """
    Get all rows of a table from an Excel file.
    Returns a list of lists (each row is a list of cell values).
    """
    file_id = get_excel_file_id(file_path)
    if not file_id:
        return []
    
    url = f"{GRAPH_API_ENDPOINT}/me/drive/items/{file_id}/workbook/tables/{table_name}/rows"
    headers = get_graph_headers()
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        rows = resp.json().get("value", [])
        return [row.get("values", [[]])[0] for row in rows]
    return []

def add_excel_row(file_path, table_name, row_values):
    """
    Add a new row to an Excel table.
    row_values: list of values corresponding to the table columns
    """
    file_id = get_excel_file_id(file_path)
    if not file_id:
        return False

    url = f"{GRAPH_API_ENDPOINT}/me/drive/items/{file_id}/workbook/tables/{table_name}/rows/add"
    headers = get_graph_headers()
    data = {
        "values": [row_values]
    }
    resp = requests.post(url, headers=headers, json=data)
    return resp.status_code == 201 or resp.status_code == 200

def update_excel_row(file_path, table_name, row_index, row_values):
    """
    Update an existing row in an Excel table by index (0-based).
    """
    file_id = get_excel_file_id(file_path)
    if not file_id:
        return False

    url = f"{GRAPH_API_ENDPOINT}/me/drive/items/{file_id}/workbook/tables/{table_name}/rows/{row_index}"
    headers = get_graph_headers()
    data = {"values": [row_values]}
    resp = requests.patch(url, headers=headers, json=data)
    return resp.status_code == 200

# Example usage:
# file_path = "/me/drive/root:/Documents/tasks.xlsx"
# table_name = "Tasks"
# all_rows = get_excel_table_rows(file_path, table_name)
# add_excel_row(file_path, table_name, ["Sebin", "New Task", "2025-10-10", "Pending"])
# update_excel_row(file_path, table_name, 2, ["Sebin", "Updated Task", "2025-10-12", "Completed"])

from io import BytesIO


GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"
EXCEL_FILE_NAME = "UserAnalytics.xlsx"


def ensure_excel_file():
    """Ensure the user analytics Excel file exists in OneDrive root."""
    headers = get_graph_headers()
    check_url = f"{GRAPH_API_ENDPOINT}/me/drive/root:/{EXCEL_FILE_NAME}"
    r = requests.get(check_url, headers=headers)

    if r.status_code == 404:
        print("📁 Creating new User_Analytics.xlsx in OneDrive root...")
        excel_data = BytesIO()
        pd.DataFrame().to_excel(excel_data, index=False)
        excel_data.seek(0)

        create_url = f"{GRAPH_API_ENDPOINT}/me/drive/root:/{EXCEL_FILE_NAME}:/content"
        resp = requests.put(create_url, headers=headers, data=excel_data.read())

        if resp.status_code in [200, 201]:
            print("✅ Created User_Analytics.xlsx successfully.")
        else:
            print("❌ Failed to create Excel file:", resp.text)
    else:
        print("✅ Excel file exists in OneDrive root.")


def update_user_analytics_excel(per_user):
    """Upload user analytics data to the Excel file in OneDrive root."""
    headers = get_graph_headers()
    ensure_excel_file()

    # Convert per_user dict to DataFrame
    df_per_user = pd.DataFrame.from_dict(per_user, orient="index").reset_index()
    df_per_user.rename(columns={"index": "User"}, inplace=True)

    # Save to Excel in memory
    excel_data = BytesIO()
    df_per_user.to_excel(excel_data, index=False)
    excel_data.seek(0)

    # Upload file to OneDrive root
    upload_url = f"{GRAPH_API_ENDPOINT}/me/drive/root:/{EXCEL_FILE_NAME}:/content"
    response = requests.put(upload_url, headers=headers, data=excel_data.read())

    if response.status_code in [200, 201]:
        print("✅ User analytics Excel updated successfully.")
    else:
        print("❌ Failed to update Excel file:", response.text)



# =========================================================
# BUSINESS CARD FUNCTIONS
# =========================================================

import os
import msal
import requests


# --- Configuration for our "Robot User" ---
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

# Using nayif onedrive account for storing the Excel file
ONEDRIVE_USER_ID = os.getenv("ONEDRIVE_USER_ID") 

EXCEL_FILE_NAME = "Contacts.xlsx"
EXCEL_TABLE_NAME = "Table1" 

MS_GRAPH_AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
MS_GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]

# --- Helper function to get an "Application" token for our robot ---
def get_application_token():
    """Gets an access token for the application itself, not a logged-in user."""
    msal_app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=MS_GRAPH_AUTHORITY, client_credential=CLIENT_SECRET
    )
    # Check cache first
    result = msal_app.acquire_token_silent(MS_GRAPH_SCOPE, account=None)
    if not result:
        print("No suitable application token in cache, acquiring a new one...")
        result = msal_app.acquire_token_for_client(scopes=MS_GRAPH_SCOPE)
    
    if "access_token" in result:
        return result["access_token"]
    else:
        print("Failed to get application token:", result.get("error_description"))
        return None

# --- Main function to READ all contacts from Excel ---
def get_contacts_from_excel():
    """
    Connects to MS Graph as the application and reads all rows from the Excel file.
    """
    access_token = get_application_token()
    if not access_token:
        return [] # Return an empty list if authentication fails

    graph_url = f"https://graph.microsoft.com/v1.0/users/{ONEDRIVE_USER_ID}/drive/root:/{EXCEL_FILE_NAME}:/workbook/tables/{EXCEL_TABLE_NAME}/rows"
    headers = {'Authorization': 'Bearer ' + access_token}
    
    try:
        response = requests.get(graph_url, headers=headers)
        response.raise_for_status() # This will raise an error for bad responses (4xx or 5xx)
        
        rows = response.json().get("value", [])
        # Convert the raw data into a list of dictionaries that our HTML can use
        contacts = []
        for i, row in enumerate(rows):
            values = row.get("values", [[]])[0]
            # Ensure the row has enough columns to prevent errors
            if len(values) >= 10: 
                contacts.append({
                    "id": i, # The row's index is its ID for editing
                    "Category": values[0],
                    "Organization": values[1],
                    "Name": values[2],
                    "Designation": values[3],
                    "Contact": values[4],
                    "Email": values[5],
                    "Website": values[6],
                    "Address": values[7],
                    "Remarks": values[8],
                    "Contact Type": values[9]
                })
        return contacts
    except requests.exceptions.RequestException as e:
        print(f"An error occurred while getting Excel data: {e}")
        # In case of an error (like file not found), return an empty list
        return []

# --- Main function to UPDATE a contact in Excel ---
def update_contact_in_excel(row_index, data):
    """
    Connects to MS Graph and updates a specific row in the Excel file by its index.
    """
    access_token = get_application_token()
    if not access_token:
        return False

    # The order of values MUST match the Excel columns exactly
    row_values = [
        data.get('Category', ''),
        data.get('Organization', ''),
        data.get('Name', ''),
        data.get('Designation', ''),
        data.get('Contact', ''),
        data.get('Email', ''),
        data.get('Website', ''),
        data.get('Address', ''),
        data.get('Remarks', ''),
        data.get('Contact Type', '')
    ]

    # The Graph API uses itemAt(index=...) to find a row by its position (0-indexed)
    graph_url = f"https://graph.microsoft.com/v1.0/users/{ONEDRIVE_USER_ID}/drive/root:/{EXCEL_FILE_NAME}:/workbook/tables/{EXCEL_TABLE_NAME}/rows/itemAt(index={row_index})"
    headers = {'Authorization': 'Bearer ' + access_token, 'Content-Type': 'application/json'}
    payload = {"values": [row_values]}

    try:
        response = requests.patch(graph_url, headers=headers, json=payload)
        response.raise_for_status() # Check for errors

        print(f"Successfully updated row at index {row_index}.")
        return True
    except requests.exceptions.RequestException as e:
        print(f"Failed to update row. Error: {e}")
        return False

