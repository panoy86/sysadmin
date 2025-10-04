import requests

# Acquire a Microsoft Graph API token using service principal credentials.
# Returns the access token as a string.
def get_graph_api_token(tenant_id, client_id, client_secret):
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default"
    }
    response = requests.post(token_url, data=token_data)
    response.raise_for_status()
    return response.json()["access_token"]


# Retrieves the first 3 messages from the specified user's mailbox using Microsoft Graph API.
def get_first_three_messages(access_token, user_email):
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages?$top=3"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json().get("value", [])


# Example usage, get the token and fetch messages
if __name__ == "__main__":
    tenant_id = "YOUR_TENANT_ID"
    client_id = "YOUR_APPLICATION_ID"
    client_secret = "YOUR_CLIENT_SECRET"
    token = get_graph_api_token(tenant_id, client_id, client_secret)
    print("Access Token:", token)

    # Test if the service principal can access a user's mailbox
    user_email = "user@example.com"
    messages = get_first_three_messages(token, user_email)
    print("First 3 Messages:")
    for msg in messages:
        date = msg.get('receivedDateTime', 'N/A')
        sender = msg.get('from', {}).get('emailAddress', {}).get('address', 'N/A')
        subject = msg.get('subject', 'N/A')
        print(f"Date: {date}\nSender: {sender}\nSubject: {subject}\n{'-'*40}")