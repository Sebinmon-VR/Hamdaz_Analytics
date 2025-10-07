import msal
import requests
import pandas as pd
from io import BytesIO
from dotenv import load_dotenv
import os
load_dotenv(override=True)

# ------------------------------------------------
# 1️⃣ Azure app credentials
# ------------------------------------------------
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")

# ------------------------------------------------
# 2️⃣ The Excel file path in OneDrive
#    Example: /Documents/Reports/sales_data.xlsx
# ------------------------------------------------
EXCEL_PATH = "/Documents/UserAnalytics.xlsx"

# ------------------------------------------------
# 3️⃣ Get an access token using MSAL
# ------------------------------------------------
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

result = app.acquire_token_silent(SCOPE, account=None)

if not result:
    result = app.acquire_token_for_client(scopes=SCOPE)

if "access_token" not in result:
    raise Exception("Could not obtain access token:", result.get("error_description"))

access_token = result["access_token"]

# ------------------------------------------------
# 4️⃣ Download the Excel file from OneDrive
# ------------------------------------------------
graph_endpoint = f"https://graph.microsoft.com/v1.0/me/drive/root:{EXCEL_PATH}:/content"

headers = {"Authorization": f"Bearer {access_token}"}
response = requests.get(graph_endpoint, headers=headers)

if response.status_code == 200:
    # Load Excel into pandas
    df = pd.read_excel(BytesIO(response.content))
    print("✅ Successfully read data from OneDrive Excel file!")
    print(df.head())
else:
    print("❌ Error:", response.status_code, response.text)
