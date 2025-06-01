import requests as req
from datetime import datetime
from msal import ConfidentialClientApplication
import base64
import os
import json
import pdfkit
from tabulate import tabulate
import copy
from azure.storage.blob import BlobServiceClient
#-------------------------- CONFIGURATION -----------------------

# Entra stuff
TENANT_ID = "66576a43-4694-4f10-9c82-043ade5de9e8"
CLIENT_ID= "3442c453-4849-427c-b76a-4fc1aa9b0b5d"
CLIENT_SECRET = "TPs8Q~ONu~ehezBlW8-H94gGcjUtBfr0Ttb.xdow"

# Graph stuff
GROUPS_ENDPOINT = "https://graph.microsoft.com/v1.0/groups?$"
FILTER = "filter=startswith(displayName, 'Test')"

# Azure Stuff
STORAGE_CONNECTION_STRING = "DefaultEndpointsProtocol=https;AccountName=vedastorage01;AccountKey=De6/bhM8xUu0ik3xGATPrS72t3vKf2lSdWusIgu4UetVQtevfx6DcRkNQ3ff5juWcNTWBMEl/xAf+ASthGAS8w==;EndpointSuffix=core.windows.net"
CONTAINER_NAME = "mycontainer"

# Other globals
FILE_NAME = "report_test_group_membership_"
FILE_TITLE = "Test Group Membership Report"

BASE_DIR = "C:/Users/ikvesi/Documents/"
WKHTMLTOPDF_PATH = r"C:\Users\ikvesi\Documents\azure_test_scripts\wkhtmltopdf\bin\wkhtmltopdf.exe"


SENDER_EMAIL = "jastivedasri25_gmail.com#EXT#@jastivedasri25gmail.onmicrosoft.com"
RECIPIENTS = ["veda.srijasti@ikpartners.com"]
EMAIL_TEMPLATE = {
        "message": {
            "subject": "Report: Test Group Membership",
            "body": {
                "contentType": "Text",
                "content": "Hi!\n\nPlease find attached the latest 'Test Group' report.\n\nKind regards,\nThe IT Dept."
            },
            "toRecipients": [],
            "attachments": [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": None,
                "contentBytes": None
            }]
        },
        "saveToSentItems": "true"
    }

# ----------------------- AUTHENTICATION -----------------------
def get_token():

    """
    Get an Azure AD token using client credentials.
    """
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = ConfidentialClientApplication(CLIENT_ID, CLIENT_SECRET, authority=authority)
    token_result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    
    if "access_token" in token_result:
        return token_result["access_token"]
    else:
        raise Exception(f"Failed to get access token. Error: {token_result.get('error_description')}")

# ----------------------- GROUP DATA FETCHING -----------------------
def get_groups():
    """
    Fetch all Azure AD groups that match the FILTER query.
    """
    token = get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{GROUPS_ENDPOINT}{FILTER}"
    response = req.get(url, headers=headers)

    if response.status_code == 200:
        return response.json()["value"]
    else:
        print(f"Error fetching groups: {response.json()}")
        return []

def get_all_group_members():
    """
    Fetch members of all filtered groups using Microsoft Graph batch API.
    """
    token = get_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    groups = get_groups()

    batch_size = 20
    group_members = {}
    batch_requests = []
    batch_counter = 1

    for group in groups:
        batch_requests.append({
            "id": str(batch_counter),
            "method": "GET",
            "url": f"/groups/{group['id']}/members"
        })
        batch_counter += 1

        if len(batch_requests) == batch_size or group == groups[-1]:
            batch_payload = {"requests": batch_requests}
            batch_response = req.post("https://graph.microsoft.com/v1.0/$batch", headers=headers, json=batch_payload)

            if batch_response.status_code == 200:
                batch_results = batch_response.json()["responses"]
                for result in batch_results:
                    group_index = int(result["id"]) - 1
                    group_name = groups[group_index]["displayName"]
                    if result["status"] == 200:
                        full_members_data = result["body"].get("value", [])
                        members = [member["displayName"] for member in full_members_data]
                        group_members[group_name] = members
                    else:
                        print(f"[✘] Error fetching members for group: {group_name}")
            else:
                print(f"[✘] Batch request failed: {batch_response.json()}")
            batch_requests = []

    return group_members

# ----------------------- SNAPSHOT COMPARISON -----------------------
def generate_snapshot(current, previous):
    """Generate a comparison snapshot showing added, removed, and unchanged group members."""
    snapshot = {}
    all_groups = set(current.keys()).union(previous.keys())
    for group in all_groups:
        cur = set(current.get(group, []))
        prev = set(previous.get(group, []))
        added = cur - prev
        removed = prev - cur
        unchanged = cur & prev
        snapshot[group] = (
            [f"{m} (new)" for m in sorted(added)] +
            [f"{m} (removed)" for m in sorted(removed)] +
            sorted(unchanged)
        )
    return snapshot

def generate_pdf(snapshot, output_file, wkhtmltopdf_path):
    """Convert the snapshot dictionary to an HTML report and generate a PDF file."""
    html = """
    <html><head><style>
    body { font-family: Arial; }
    .new { color: green; }
    .removed { color: orange; }
    </style></head><body><h1>Azure AD Group Report</h1>
    """
    for group, members in snapshot.items():
        html += f"<h2>{group}</h2><ul>"
        for m in members:
            cls = "new" if m.endswith("(new)") else "removed" if m.endswith("(removed)") else ""
            html += f"<li class='{cls}'>{m}</li>"
        html += "</ul>"
    html += "</body></html>"
    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
    pdfkit.from_string(html, output_file, configuration=config)

def upload_to_blob(local_path, blob_name):
    """Upload a local file to Azure Blob Storage under the specified blob name."""
    blob_service = BlobServiceClient.from_connection_string(STORAGE_CONNECTION_STRING)
    blob_client = blob_service.get_blob_client(CONTAINER_NAME, blob=blob_name)
    with open(local_path, "rb") as file:
        blob_client.upload_blob(file, overwrite=True)
    return blob_name

#def email_result(filename, blob_data):
    """Send an email with a file attachment using Microsoft Graph API."""
    token = get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    attachment_data = base64.b64encode(blob_data).decode("utf-8")
    email_json = copy.deepcopy(EMAIL_TEMPLATE)
    email_json["message"]["toRecipients"] = [{"emailAddress": {"address": r}} for r in RECIPIENTS]
    email_json["message"]["attachments"][0].update({"name": filename, "contentBytes": attachment_data})
    # Send email
    email_endpoint = f"https://graph.microsoft.com/v1.0/users/jastivedasri25_gmail.com%23EXT%23@jastivedasri25gmail.onmicrosoft.com/sendMail"
    #jastivedasri25_gmail.com%23EXT%23@jastivedasri25gmail.onmicrosoft.com

    response = req.post(email_endpoint, headers=headers, json=email_json)
    if response.status_code == 202:
        print(f"Email successfully sent.")
    else:
        print(f"Failed to send email, please investigate or try again. Error: {response.status_code}")

# ----------------------- MAIN -----------------------
WKHTMLTOPDF_PATH = r"C:\\Users\\ikvesi\\Documents\\azure_test_scripts\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
snapshot_json_file = f"snapshot_{timestamp}.json"
snapshot_pdf_file = f"snapshot_{timestamp}.pdf"

print("📥 Fetching group members...")
current = get_all_group_members()

print("📤 Loading last snapshot...")
previous = {}  # Can pull from blob if needed

print("🔄 Generating comparison snapshot...")
snapshot = generate_snapshot(current, previous)

print("📝 Writing snapshot to local JSON for upload...")
with open(snapshot_json_file, "w", encoding="utf-8") as f:
    json.dump(snapshot, f, indent=2)

print("📄 Generating PDF...")
generate_pdf(snapshot, snapshot_pdf_file, WKHTMLTOPDF_PATH)

print("☁️ Uploading to Azure Blob...")
json_blob_name = upload_to_blob(snapshot_json_file, snapshot_json_file)
pdf_blob_name = upload_to_blob(snapshot_pdf_file, snapshot_pdf_file)

#print("📧 Sending email with PDF attached...")
#with open(snapshot_pdf_file, "rb") as file:
   # email_result(snapshot_pdf_file, file.read())

print("✅ All done. Snapshots uploaded and email sent.")
