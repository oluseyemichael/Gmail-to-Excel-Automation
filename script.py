import os
import base64
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from dotenv import load_dotenv

# Gmail API configuration
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
load_dotenv()

def authenticate_gmail():
    """Authenticate and create Gmail API service."""
    CLIENT_SECRET_FILE = os.getenv("GMAIL_CLIENT_SECRET")
    TOKEN_FILE = os.getenv("TOKEN_FILE")

    # Check for existing credentials
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    
    # Refresh or create new credentials if needed
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        
        # Save new credentials
        with open(TOKEN_FILE, "w") as token:
            token.write(creds.to_json())
    
    return build("gmail", "v1", credentials=creds)

def parse_email_date(date_string):
    """Parse email date with multiple possible formats."""
    date_formats = [
        "%a, %d %b %Y %H:%M:%S %z",
        "%Y-%m-%d %H:%M:%S"
    ]
    
    for date_format in date_formats:
        try:
            return datetime.strptime(date_string, date_format).strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            continue
    
    return "Unknown Date"

def extract_email_body(payload):
    """Extract email body, preferring plain text over HTML."""
    # Check direct body first
    if "data" in payload.get("body", {}):
        return base64.urlsafe_b64decode(payload["body"]["data"]).decode("utf-8", errors="ignore")
    
    # Check parts
    for part in payload.get("parts", []):
        mime_type = part.get("mimeType")
        if mime_type == "text/plain":
            return base64.urlsafe_b64decode(part["body"]["data"]).decode("utf-8", errors="ignore")
        
        if mime_type == "text/html":
            html_content = base64.urlsafe_b64decode(part["body"]["data"]).decode("utf-8", errors="ignore")
            return BeautifulSoup(html_content, "html.parser").get_text()
    
    return "No Body Found"

def fetch_emails(service, max_emails=50):
    """Fetch unread emails from Gmail."""
    results = service.users().messages().list(
        userId="me", 
        labelIds=["INBOX"], 
        q="is:unread"
    ).execute()
    
    messages = results.get("messages", [])
    
    if not messages:
        print("No unread emails found.")
        return []

    print(f"Found {len(messages)} unread emails!")
    
    email_data = []
    for msg in messages[:max_emails]:
        msg_content = service.users().messages().get(userId="me", id=msg["id"]).execute()
        payload = msg_content.get("payload", {})
        headers = payload.get("headers", [])

        # Extract email metadata
        sender = next((h["value"] for h in headers if h["name"] == "From"), "Unknown")
        subject = next((h["value"] for h in headers if h["name"] == "Subject"), "No Subject")
        email_date = next((h["value"] for h in headers if h["name"] == "Date"), "Unknown Date")
        
        # Parse date and body
        parsed_date = parse_email_date(email_date)
        email_body = extract_email_body(payload)

        email_data.append([parsed_date, sender, subject, email_body[:200]])

    return email_data

def save_to_excel(data, filename="unread_mails.xlsx"):
    """Save email data to Excel spreadsheet."""
    df = pd.DataFrame(data, columns=["Date", "Sender", "Subject", "Email Preview"])
    df.to_excel(filename, index=False)
    print(f"Data saved successfully to {filename}")

def main():
    """Main execution flow with error handling."""
    try:
        gmail_service = authenticate_gmail()
        email_records = fetch_emails(gmail_service)

        if email_records:
            save_to_excel(email_records)
        else:
            print("No new emails found matching the criteria.")
    except PermissionError:
        print("Cannot save file. Please close any open Excel files and try again.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()