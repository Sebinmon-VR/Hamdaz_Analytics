import os
import requests
from flask import session, redirect
from dotenv import load_dotenv
from auth import *
from collections import defaultdict
import pandas as pd
import pytz
from datetime import datetime

load_dotenv(override=True)


def get_all_users():
    """
    Fetch all users in the organization.
    Returns a list of user dictionaries or None if unauthorized.
    """
    headers = get_graph_headers()
    if not headers:
        access_token = refresh_access_token()
        if not access_token:
            return None
        headers = get_graph_headers()

    users = []
    endpoint = f"{GRAPH_API_ENDPOINT}/users?$top=999"  # up to 999 users per request

    while endpoint:
        response = requests.get(endpoint, headers=headers)
        if response.status_code != 200:
            print("Error fetching users:", response.json())
            return None

        data = response.json()
        users.extend(data.get("value", []))
        # Pagination support
        endpoint = data.get("@odata.nextLink", None)

    return users

def get_file_id(file_path):
    """
    Get file ID from OneDrive file path
    Example: /drive/root:/Sharepoint Datas.xlsx
    """
    headers = get_graph_headers()
    url = f"{GRAPH_API_ENDPOINT}{file_path}"
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json().get("id")
    else:
        print("Error:", response.json())
        return None

def get_excel_tables(file_path):
    """
    Fetch all tables from an Excel file in OneDrive
    """
    headers = get_graph_headers()
    file_id = get_file_id(file_path)

    if not file_id:
        return []

    url = f"{GRAPH_API_ENDPOINT}/me/drive/items/{file_id}/workbook/tables"
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json().get("value", [])
    else:
        print("Error fetching tables:", response.json())
        return []

def get_table_data(file_path, table_name):
    """
    Get rows of a specific table in the Excel file
    """
    headers = get_graph_headers()
    file_id = get_file_id(file_path)

    if not file_id:
        return []

    url = f"{GRAPH_API_ENDPOINT}/me/drive/items/{file_id}/workbook/tables/{table_name}/rows"
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json().get("value", [])
    else:
        print("Error fetching table rows:", response.json())
        return []

def get_my_user_id():
    """
    Fetch the signed-in user's ID from Microsoft Graph
    """
    headers = get_graph_headers()
    if not headers:
        return None

    url = f"{GRAPH_API_ENDPOINT}/me"
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json().get("id")
    else:
        print("Error fetching user ID:", response.json())
        return None

def get_users_analytics(file_path):
    """
    Returns analytics per user from the Excel file.
    Example output:
    {
        "John Doe": {"total_tasks": 10, "active_tasks": 3},
        "Jane Smith": {"total_tasks": 7, "active_tasks": 2},
    }
    """
    tables = get_excel_tables(file_path)
    analytics = defaultdict(lambda: {"total_tasks": 0, "active_tasks": 0})

    today = datetime.now().date()

    for table in tables:
        table_name = table.get("name")
        rows = get_table_data(file_path, table_name)

        for row in rows:
            values = row.get("values", [[]])[0]  # Excel API returns values as nested list
            if len(values) < 4:
                continue  # Skip if not enough columns

            user = values[0]
            task_name = values[1]
            due_date = values[2]
            status = values[3]

            analytics[user]["total_tasks"] += 1

            # Count as active if due_date >= today and status is not "Completed"
            try:
                due_date_obj = datetime.strptime(due_date, "%Y-%m-%d").date()
                if due_date_obj >= today and status.lower() != "completed":
                    analytics[user]["active_tasks"] += 1
            except:
                pass  # skip invalid dates

    return dict(analytics)


def sharepoint_data_to_df(structured_items):
    """
    Converts SharePoint list structured items into a pandas DataFrame.
    Ensures necessary columns exist for analytics.
    """
    if not structured_items:
        return pd.DataFrame()
    
    df = pd.DataFrame(structured_items)

    # Make sure required columns exist
    required_cols = ["AssignedTo", "Priority", "Status", "SubmissionStatus", "BCD", "DueDate", "Title", "id"]
    for col in required_cols:
        if col not in df.columns:
            df[col] = None

    # Ensure BCD and DueDate are datetime for calculations
    for date_col in ["BCD", "DueDate"]:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce', utc=True)

    return df
import pytz
from datetime import datetime

def compute_overall_analytics(df):
    if df.empty:
        return {
            "total_users": 0,
            "total_tasks": 0,
            "tasks_completed": 0,
            "tasks_pending": 0,
            "tasks_missed": 0,
            "orders_received": 0
        }

    # UAE timezone
    uae_tz = pytz.timezone("Asia/Dubai")
    now_uae = datetime.now(uae_tz)

    # Ensure BCD is datetime
    if 'BCD' in df.columns:
        df['BCD'] = pd.to_datetime(df['BCD'], errors='coerce', utc=True).dt.tz_convert(uae_tz)

    total_users = df['AssignedTo'].nunique() if 'AssignedTo' in df.columns else 0
    total_tasks = len(df)

    # Completed tasks
    completed_mask = (df['SubmissionStatus'] == 'Submitted')
    tasks_completed = len(df[completed_mask])

    # Pending tasks: not submitted AND BCD >= now
    pending_mask = (df['SubmissionStatus'] != 'Submitted') & (df['BCD'] >= now_uae)
    tasks_pending = len(df[pending_mask])

    # Missed tasks: not submitted AND BCD < now
    missed_mask = (df['SubmissionStatus'] != 'Submitted') & (df['BCD'] < now_uae)
    tasks_missed = len(df[missed_mask])

    orders_received = len(df[df['Status'] == 'Received']) if 'Status' in df.columns else 0

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

    if 'BCD' in df.columns:
        df['BCD'] = pd.to_datetime(df['BCD'], errors='coerce', utc=True).dt.tz_convert(uae_tz)

    user_analytics = {}
    for user, user_df in df.groupby('AssignedTo'):
        total_tasks = len(user_df)
        tasks_completed = len(user_df[user_df['SubmissionStatus'] == 'Submitted'])
        tasks_pending = len(user_df[(user_df['SubmissionStatus'] != 'Submitted') & (user_df['BCD'] >= now_uae)])
        tasks_missed = len(user_df[(user_df['SubmissionStatus'] != 'Submitted') & (user_df['BCD'] < now_uae)])
        orders_received = len(user_df[user_df['Status'] == 'Received']) if 'Status' in df.columns else 0

        user_analytics[user] = {
            "total_tasks": total_tasks,
            "tasks_completed": tasks_completed,
            "tasks_pending": tasks_pending,
            "tasks_missed": tasks_missed,
            "orders_received": orders_received
        }

    return user_analytics

# functions.py

def compute_user_analytics_specific(sp_items, username):
    """
    Computes analytics for a specific user.
    
    sp_items: list of dictionaries representing tasks from SharePoint
    username: string of the user's name
    
    Returns a dictionary with task counts:
    {
        "tasks": total tasks assigned,
        "submissions": completed tasks,
        "pending": pending tasks,
        "missed": missed tasks
    }
    """
    # Timezone for comparison
    uae_tz = pytz.timezone("Asia/Dubai")
    now_uae = datetime.now(uae_tz)

    # Filter items assigned to this user
    user_items = [item for item in sp_items if item.get("AssignedTo") == username]

    total_tasks = len(user_items)
    submissions = 0
    pending = 0
    missed = 0

    for item in user_items:
        # Normalize column names (just in case)
        submission_status = item.get("SubmissionStatus") or item.get("Submission status") or ""
        status = item.get("Status") or ""
        bcd_str = item.get("BCD")  # due date as string

        # Convert BCD to datetime
        try:
            bcd_dt = datetime.strptime(bcd_str, "%Y-%m-%dT%H:%M:%S") if bcd_str else None
            if bcd_dt:
                bcd_dt = uae_tz.localize(bcd_dt)
        except:
            bcd_dt = None

        # Count submissions
        if submission_status.lower() == "submitted" or submission_status.lower() == "done":
            submissions += 1
        # Count missed
        elif bcd_dt and bcd_dt < now_uae:
            missed += 1
        # Pending
        else:
            pending += 1

    return {
        "tasks": total_tasks,
        "submissions": submissions,
        "pending": pending,
        "missed": missed
    }

# -------------------------------
# Example: org users function
# -------------------------------
def get_all_org_users():
    """
    Fetch all users in organization using Microsoft Graph API
    """
    import requests
    from auth import get_graph_headers  # your existing auth function

    url = "https://graph.microsoft.com/v1.0/users?$select=displayName,userPrincipalName"
    headers = get_graph_headers()
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        users = resp.json().get("value", [])
        return [u["displayName"] for u in users]
    else:
        print("Error fetching org users:", resp.json())
        return []


def compute_teams_analytics(sp_items):
    """
    Compute overall and per-user analytics from SharePoint list items.
    sp_items: list of dicts, each dict represents a SharePoint list item
    """

    if not sp_items:
        return {
            "total_tasks": 0,
            "total_submissions": 0,
            "total_pending": 0,
            "total_missed": 0,
            "users": {}
        }

    # Get current UAE datetime
    uae_tz = pytz.timezone("Asia/Dubai")
    now = datetime.now(uae_tz)

    overall = {
        "total_tasks": 0,
        "total_submissions": 0,
        "total_pending": 0,
        "total_missed": 0,
    }
    users = {}

    for item in sp_items:
        overall["total_tasks"] += 1

        username = item.get("AssignedTo") or "Unassigned"
        if username not in users:
            users[username] = {
                "tasks": 0,
                "submissions": 0,
                "pending": 0,
                "missed": 0
            }

        # Parse BCD date
        bcd_str = item.get("BCD")
        submission_status = item.get("SubmissionStatus") or ""
        try:
            bcd_date = datetime.fromisoformat(bcd_str.replace("Z", "+00:00")) if bcd_str else None
        except:
            bcd_date = None

        users[username]["tasks"] += 1

        # Submitted tasks
        if submission_status.lower() == "submitted":
            overall["total_submissions"] += 1
            users[username]["submissions"] += 1

        # Pending tasks (BCD not passed OR not submitted)
        if bcd_date and bcd_date >= now:
            users[username]["pending"] += 1
            overall["total_pending"] += 1

        # Missed tasks (BCD passed and not submitted)
        if bcd_date and bcd_date < now and submission_status.lower() != "submitted":
            users[username]["missed"] += 1
            overall["total_missed"] += 1

    overall["users"] = users
    return overall


def get_site_id(hostname=None, site_path=None):
    """
    Fetch the SharePoint site ID.
    - If hostname and site_path are provided, fetch that specific site.
    - If not, return the root site (default).
    """
    headers = get_graph_headers()
    if not headers:
        return None

    if hostname and site_path:
        # Example: hostname="contoso.sharepoint.com", site_path="/sites/YourSiteName"
        url = f"{GRAPH_API_ENDPOINT}/sites/{hostname}:{site_path}"
    else:
        # Get the root site (default site for logged-in user)
        url = f"{GRAPH_API_ENDPOINT}/sites/root"

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        site = response.json()
        return site.get("id")
    else:
        print("Error fetching site ID:", response.json())
        return None


SHAREPOINT_HOSTNAME = "hamdaz1.sharepoint.com"

import requests
from auth import get_graph_headers  # your existing auth function

GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"

SITE_NAME = "ProposalTeam"
LIST_NAME = "Proposals"

def get_site_id(site_name):
    headers = get_graph_headers()
    url = f"{GRAPH_API_ENDPOINT}/sites/hamdaz1.sharepoint.com:/sites/{site_name}"
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        return resp.json().get("id")
    print("Error fetching site ID:", resp.json())
    return None

def get_list_id(site_id, list_name):
    headers = get_graph_headers()
    url = f"{GRAPH_API_ENDPOINT}/sites/{site_id}/lists"
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        for l in resp.json().get("value", []):
            if l.get("name") == list_name:
                return l.get("id")
    print("Error fetching lists:", resp.json())
    return None
def get_list_items(site_id, list_id):
    """
    Get all items from a SharePoint list with expanded person/group fields, including pagination.
    """
    headers = get_graph_headers()
    expand_fields = "fields($expand=AssignedTo,Author,Editor)"
    url = f"{GRAPH_API_ENDPOINT}/sites/{site_id}/lists/{list_id}/items?expand={expand_fields}"

    all_items = []

    while url:
        resp = requests.get(url, headers=headers)
        if resp.status_code == 200:
            data = resp.json()
            all_items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")  # Get the next page URL if exists
        else:
            print("Error fetching list items:", resp.json())
            break

    return all_items


def flatten_fields(fields):
    """
    Flatten SharePoint list item fields for structured output
    """
    flat = {}
    for k, v in fields.items():
        if isinstance(v, dict):
            # Person or Group fields
            if 'displayName' in v:
                flat[k] = v['displayName']
            # Lookup fields
            elif 'lookupValue' in v:
                flat[k] = v['lookupValue']
            else:
                flat[k] = str(v)
        elif isinstance(v, list):
            # Multi-value person or choice fields
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
    structured_items = []
    for item in items:
        fields = item.get("fields", {})
        structured_items.append(flatten_fields(fields))
    return structured_items


import base64

def get_users_with_photos():
    """
    Fetches all users from Microsoft Graph along with their profile pictures.
    Returns a list of dictionaries:
    [
        {
            "id": "user-id",
            "displayName": "User Name",
            "mail": "user@example.com",
            "photo": "data:image/jpeg;base64,..."
        },
        ...
    ]
    """
    access_token = session.get('access_token')
    if not access_token:
        return []

    headers = {'Authorization': f'Bearer {access_token}'}

    # Step 1: Get all users
    response = requests.get(
        'https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail',
        headers=headers
    )
    if response.status_code != 200:
        return []

    users = response.json().get('value', [])

    # Step 2: Fetch profile photos for each user
    for user in users:
        user_id = user['id']
        photo_response = requests.get(f'https://graph.microsoft.com/v1.0/users/{user_id}/photo/$value', headers=headers)
        if photo_response.status_code == 200:
            photo_b64 = base64.b64encode(photo_response.content).decode('utf-8')
            user['photo'] = f"data:image/jpeg;base64,{photo_b64}"
        else:
            user['photo'] = None

    return users



























# import pandas as pd

# def load_excel_to_df(file_path):
#     """
#     Load only Table1 from the Excel file and convert it to a DataFrame
#     """
#     tables = get_excel_tables(file_path)

#     # Find Table1
#     table1 = next((t for t in tables if t.get("name") == "Table1"), None)
#     if not table1:
#         print("Table1 not found")
#         return pd.DataFrame(columns=["Username", "Submission status", "Order Status"])

#     rows = get_table_data(file_path, "Table1")
#     all_rows = []

#     # Use Table1's columns
#     columns = [c.get("name") for c in table1.get("columns", [])]

#     for r in rows:
#         if "values" in r:
#             for row_values in r["values"]:  # <-- iterate ALL rows
#                 if not columns or len(columns) != len(row_values):
#                     columns = [f"col{i}" for i in range(len(row_values))]
#                 all_rows.append(dict(zip(columns, row_values)))

#     df = pd.DataFrame(all_rows)

#     # Normalize column names (remove spaces, lowercase)
#     df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]

#     # Map to standard names expected by analytics functions
#     col_mapping = {
#         "username": "Username",
#         "submission_status": "Submission status",
#         "submissionstatus": "Submission status",
#         "order_status": "Order Status",
#         "orderstatus": "Order Status"
#     }
#     df.rename(columns=col_mapping, inplace=True)

#     # Ensure required columns exist
#     for col in ["Username", "Submission status", "Order Status"]:
#         if col not in df.columns:
#             df[col] = None

#     return df

# def compute_overall_analytics(df):
#     """
#     Compute the key metrics for the dashboard
#     """
#     if df.empty:
#         return {
#             "total_users": 0,
#             "total_tasks": 0,
#             "tasks_completed": 0,
#             "tasks_pending": 0,
#             "orders_received": 0
#         }

#     total_users = df['Username'].nunique()
#     total_tasks = len(df)
#     tasks_completed = len(df[df['Submission status'] == 'Submitted'])
#     tasks_pending = total_tasks - tasks_completed
#     orders_received = len(df[df['Order Status'] == 'Received'])

#     return {
#         "total_users": total_users,
#         "total_tasks": total_tasks,
#         "tasks_completed": tasks_completed,
#         "tasks_pending": tasks_pending,
#         "orders_received": orders_received
#     }

# ----------------------------------------------------------------------------------------------