import requests as req
from datetime import datetime
from msal import ConfidentialClientApplication
import base64
import os
import json
import pdfkit
import copy
from azure.storage.blob import BlobServiceClient
#-------------------------- CONFIGURATION -----------------------

# Entra stuff
#TENANT_ID = "66576a43-4694-4f10-9c82-043ade5de9e8"
#CLIENT_ID= "3442c453-4849-427c-b76a-4fc1aa9b0b5d"
#CLIENT_SECRET = "TPs8Q~ONu~ehezBlW8-H94gGcjUtBfr0Ttb.xdow"

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
    Get an Azure AD token using client credentials from pipeline variables.
    """
    # Get credentials from pipeline environment variables
    tenant_id = os.environ.get('TENANT_ID')
    client_id = os.environ.get('CLIENT_ID') 
    client_secret = os.environ.get('CLIENT_SECRET')
    
    if not all([tenant_id, client_id, client_secret]):
        raise Exception("Missing required environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET")
    
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = ConfidentialClientApplication(client_id, client_secret, authority=authority)
    token_result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    
    if "access_token" in token_result:
        return token_result["access_token"]
    else:
        raise Exception(f"Failed to get access token. Error: {token_result.get('error_description')}")

# ----------------------- GROUP DATA FETCHING -----------------------
def get_groups():
    """
    Fetch all Azure AD groups that match the filter query.
    """
    token = get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    # Get filter from environment variable or use default
    filter_query = os.environ.get('GROUPS_FILTER', '')
    groups_endpoint = "https://graph.microsoft.com/v1.0/groups"
    url = f"{groups_endpoint}?{filter_query}" if filter_query else groups_endpoint
    
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

# ----------------------- PIPELINE ARTIFACT FUNCTIONS -----------------------
def load_previous_snapshot():
    """
    Load the previous snapshot from pipeline artifacts directory.
    In Azure Pipelines, we'll check if a previous snapshot exists.
    """
    artifacts_dir = os.environ.get('PIPELINE_WORKSPACE', './artifacts')
    previous_snapshot_path = os.path.join(artifacts_dir, 'previous_snapshot.json')
    
    if os.path.exists(previous_snapshot_path):
        print(f"📂 Loading previous snapshot from: {previous_snapshot_path}")
        try:
            with open(previous_snapshot_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"⚠️ Error loading previous snapshot: {e}")
            return {}
    else:
        print("📂 No previous snapshot found, treating as initial run")
        return {}

def save_current_as_previous(current_snapshot, artifacts_dir):
    """
    Save current snapshot as previous for next pipeline run.
    """
    previous_snapshot_path = os.path.join(artifacts_dir, 'previous_snapshot.json')
    with open(previous_snapshot_path, 'w', encoding='utf-8') as f:
        json.dump(current_snapshot, f, indent=2)
    print(f"💾 Saved current snapshot as previous: {previous_snapshot_path}")

# ----------------------- SNAPSHOT COMPARISON -----------------------
def generate_snapshot(current, previous):
    """Generate a comparison snapshot showing added, removed, and unchanged group members."""
    snapshot = {}
    all_groups = set(current.keys()).union(previous.keys())
    
    changes_detected = False
    
    for group in all_groups:
        cur = set(current.get(group, []))
        prev = set(previous.get(group, []))
        added = cur - prev
        removed = prev - cur
        unchanged = cur & prev
        
        if added or removed:
            changes_detected = True
        
        snapshot[group] = {
            'added': sorted(list(added)),
            'removed': sorted(list(removed)),
            'unchanged': sorted(list(unchanged)),
            'total_members': len(cur)
        }
    
    return snapshot, changes_detected

def generate_comparison_report(snapshot, changes_detected):
    """Generate a readable comparison report."""
    report_lines = []
    report_lines.append("# Azure AD Group Membership Report")
    report_lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report_lines.append("")
    
    if not changes_detected:
        report_lines.append("✅ No changes detected in group memberships.")
        report_lines.append("")
    
    total_groups = len(snapshot)
    groups_with_changes = sum(1 for group_data in snapshot.values() 
                             if group_data['added'] or group_data['removed'])
    
    report_lines.append(f"## Summary")
    report_lines.append(f"- Total Groups: {total_groups}")
    report_lines.append(f"- Groups with Changes: {groups_with_changes}")
    report_lines.append("")
    
    for group_name, group_data in snapshot.items():
        if group_data['added'] or group_data['removed'] or not changes_detected:
            report_lines.append(f"## {group_name}")
            report_lines.append(f"Total Members: {group_data['total_members']}")
            
            if group_data['added']:
                report_lines.append(f"### ➕ Added Members ({len(group_data['added'])})")
                for member in group_data['added']:
                    report_lines.append(f"- {member}")
                report_lines.append("")
            
            if group_data['removed']:
                report_lines.append(f"### ➖ Removed Members ({len(group_data['removed'])})")
                for member in group_data['removed']:
                    report_lines.append(f"- {member}")
                report_lines.append("")
            
            if group_data['unchanged'] and not changes_detected:
                report_lines.append(f"### 👥 Current Members ({len(group_data['unchanged'])})")
                for member in group_data['unchanged']:
                    report_lines.append(f"- {member}")
                report_lines.append("")
            
            report_lines.append("---")
            report_lines.append("")
    
    return '\n'.join(report_lines)

# ----------------------- MAIN PIPELINE FUNCTION -----------------------
def main():
    """
    Main function to run the group report generation in Azure Pipeline.
    """
    # Set up artifacts directory
    artifacts_dir = os.environ.get('BUILD_ARTIFACTSTAGINGDIRECTORY', './pipeline-artifacts')
    os.makedirs(artifacts_dir, exist_ok=True)
    
    # Generate timestamp for this run
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    try:
        print("🚀 Starting Azure AD Group Report Generation")
        print(f"📁 Artifacts directory: {artifacts_dir}")
        
        # Step 1: Fetch current group members
        print("📥 Fetching current group members...")
        current = get_all_group_members()
        print(f"✅ Fetched members for {len(current)} groups")
        
        # Step 2: Load previous snapshot
        print("📤 Loading previous snapshot...")
        previous = load_previous_snapshot()
        
        # Step 3: Generate comparison snapshot
        print("🔄 Generating comparison snapshot...")
        snapshot, changes_detected = generate_snapshot(current, previous)
        
        # Step 4: Generate reports
        print("📝 Generating reports...")
        
        # Save detailed snapshot as JSON
        snapshot_json_file = os.path.join(artifacts_dir, f"group_snapshot_{timestamp}.json")
        with open(snapshot_json_file, "w", encoding="utf-8") as f:
            json.dump(snapshot, f, indent=2)
        print(f"💾 Saved detailed snapshot: {snapshot_json_file}")
        
        # Generate readable report
        report_content = generate_comparison_report(snapshot, changes_detected)
        report_file = os.path.join(artifacts_dir, f"group_report_{timestamp}.md")
        with open(report_file, "w", encoding="utf-8") as f:
            f.write(report_content)
        print(f"📄 Generated report: {report_file}")
        
        # Save current data as raw JSON for troubleshooting
        raw_data_file = os.path.join(artifacts_dir, f"raw_group_data_{timestamp}.json")
        with open(raw_data_file, "w", encoding="utf-8") as f:
            json.dump(current, f, indent=2)
        print(f"🔧 Saved raw data: {raw_data_file}")
        
        # Step 5: Save current snapshot as previous for next run
        save_current_as_previous(current, artifacts_dir)
        
        # Step 6: Set pipeline variables for downstream tasks
        if changes_detected:
            print("##vso[task.setvariable variable=GroupChangesDetected;isOutput=true]true")
            print(f"##vso[task.setvariable variable=GroupsChanged;isOutput=true]{groups_with_changes}")
        else:
            print("##vso[task.setvariable variable=GroupChangesDetected;isOutput=true]false")
            print("##vso[task.setvariable variable=GroupsChanged;isOutput=true]0")
        
        print("✅ Group report generation completed successfully!")
        
        if changes_detected:
            print(f"⚠️ Changes detected in {groups_with_changes} groups")
        else:
            print("ℹ️ No changes detected")
            
    except Exception as e:
        print(f"❌ Error during execution: {str(e)}")
        print("##vso[task.logissue type=error]Group report generation failed")
        raise

if __name__ == "__main__":
    main()