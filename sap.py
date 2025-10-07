import requests

# Replace with your region's token endpoint
token_url = "https://api.ariba.com/v2/oauth/token"
client_id = "<CLIENT_ID>"
client_secret = "<CLIENT_SECRET>"

# Obtain access token
token_response = requests.post(
    token_url,
    data={
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret
    },
    headers={"Content-Type": "application/x-www-form-urlencoded"}
)

if token_response.status_code == 200:
    access_token = token_response.json()["access_token"]
else:
    print("Error obtaining access token:", token_response.status_code, token_response.text)
    exit()

# Define the events endpoint
events_url = "https://openapi.ariba.com/v1/events"

# Set up headers with the obtained access token
headers = {
    "Authorization": f"Bearer {access_token}",
    "Accept": "application/json"
}

# Make the API request to retrieve events
events_response = requests.get(events_url, headers=headers)

if events_response.status_code == 200:
    events = events_response.json()
    for event in events.get("items", []):
        print(f"Event ID: {event['id']}, Title: {event['title']}, Status: {event['status']}")
else:
    print("Error retrieving events:", events_response.status_code, events_response.text)
