import os
import json
import pdfkit
import base64
import urllib.parse
import unicodedata
import re
from datetime import datetime
import requests as req
from msal import ConfidentialClientApplication

# --- Recipient helpers ---
SEPARATORS = re.compile(r"[;, \n\r\t]+")

def parse_recipients(value: str) -> list[str]:
    """Parse a string of email addresses into a list, handling various separators and de-duplicating."""
    if not value:
        return []
    items = [x.strip() for x in SEPARATORS.split(value) if x.strip()]
    # De-dup while keeping order
    seen = set()
    result = []
    for x in items:
        if x.lower() not in seen:
            seen.add(x.lower())
            result.append(x)
    return result

def to_recipient_objects(addresses: list[str]) -> list[dict]:
    """Convert emails to Graph recipient objects."""
    return [{"emailAddress": {"address": a}} for a in addresses]

#------------------- Text Normalization ------------------
def clean_text(s: str) -> str:
    if not s:
        return ""
    # strip BOM, trim, normalize accents to NFC
    return unicodedata.normalize("NFC", s.lstrip("\ufeff").strip())

def normalize_snapshot_keys(snap: dict) -> dict:
    """Normalize group names and member names inside a snapshot dict."""
    fixed = {}
    for k, v in (snap or {}).items():
        nk = clean_text(k)
        fixed[nk] = sorted({ clean_text(m) for m in (v or []) })
    return fixed
# ------------------ Authentication ------------------
def get_token():
    tenant_id = os.environ.get('TENANT_ID')
    client_id = os.environ.get('CLIENT_ID')
    client_secret = os.environ.get('CLIENT_SECRET')
    
    if not all([tenant_id, client_id, client_secret]):
        raise Exception("Missing environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET")

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = ConfidentialClientApplication(client_id, client_secret, authority=authority)
    token_result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if "access_token" in token_result:
        return token_result["access_token"]
    else:
        raise Exception(f"Token error: {token_result.get('error_description')}")


#-----------------Load  Groups from CSV --------------
def load_groups_from_csv(file_path):
    groups = []
    # utf-8-sig strips BOM on first line automatically
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        for line in f:
            name = clean_text(line)
            if name:
                groups.append(name)
    return groups


#---------------Get Group ids from names----------------
def get_group_ids_from_names(group_names):
    token = get_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    group_ids = {}

    for name in group_names:
        # Escape single quotes for Graph filter syntax
        escaped_name = name.replace("'", "''")

        # URL encode only the filter value, not the whole query string
        filter_query = f"displayName eq '{escaped_name}'"
        encoded_query = urllib.parse.quote(filter_query, safe="=$'()")

        url = f"https://graph.microsoft.com/v1.0/groups?$filter={encoded_query}"
        response = req.get(url, headers=headers)
        response.raise_for_status()

        data = response.json().get("value", [])
        if data:
            group_ids[name] = data[0]["id"]
        else:
            print(f"[WARN] No match found for group '{name}'")

    return group_ids


# ------------- Fetch members for ONE group (handles paging) -------------
def fetch_group_members(group_id, headers):
    # Pull members with paging; keep displayName normalized
    url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/members?$select=id,displayName,userPrincipalName&$top=999"
    members = []
    while url:
        r = req.get(url, headers=headers)
        r.raise_for_status()
        j = r.json()
        for m in j.get("value", []):
            dn = clean_text(m.get("displayName") or m.get("id"))
            members.append(dn)
        url = j.get("@odata.nextLink")
    return members

def get_all_group_members():
    token = get_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    # 1) load names (one per line)
    group_names = load_groups_from_csv('inputs/critical_groups.csv')

    # 2) resolve names -> IDs (uses your safe resolver)
    name_to_id = get_group_ids_from_names(group_names)

    print(f"Total groups in input: {len(group_names)}")
    print(f"Resolved IDs: {len(name_to_id)}")

    # 3) fetch and normalize
    group_members = {}
    for name in group_names:
        gid = name_to_id.get(name)
        if not gid:
            print(f"[WARN] Skipping '{name}' – no GroupId resolved.")
            continue
        members = fetch_group_members(gid, headers)
        group_members[clean_text(name)] = sorted(set(clean_text(x) for x in members))

    return group_members



# ------------------ Snapshot Handling ------------------
def load_previous_snapshot():
    path = os.path.join(
        os.environ.get("PIPELINE_WORKSPACE", "./"),
        "group-report-artifacts",
        "previous_snapshot.json",
    )
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return normalize_snapshot_keys(data)
    print("No previous snapshot found, treating this as first run.")
    return {}


def save_current_snapshot(data):
    artifacts_dir = os.environ.get('BUILD_ARTIFACTSTAGINGDIRECTORY', './pipeline-artifacts')
    os.makedirs(artifacts_dir, exist_ok=True)
    # normalize before save; ensure_ascii=False keeps accented chars as-is
    norm = normalize_snapshot_keys(data)
    with open(os.path.join(artifacts_dir, 'previous_snapshot.json'), 'w', encoding='utf-8') as f:
        json.dump(norm, f, indent=2, ensure_ascii=False)

# ------------------ Comparison Logic ------------------
def compare_snapshots(current, previous):
    result = {}
    added_groups = []
    deleted_groups = []
    all_keys = set(current.keys()).union(previous.keys())
    changes_detected = False

    for group in all_keys:
        cur_members = set(current.get(group, []))
        prev_members = set(previous.get(group, []))

        # Check if group is added
        if group not in previous:
            added_groups.append(group)
            result[group] = {
                'added': sorted(current[group]),
                'removed': [],
                'unchanged': []
            }
            changes_detected = True

        # Check if group is deleted
        elif group not in current:
            deleted_groups.append(group)
            result[group] = {   
                'added': [],
                'removed': sorted(previous[group]),
                'unchanged': []
            }
            changes_detected = True

        # Check if group has changes
        else:
            added = list(cur_members - prev_members)
            removed = list(prev_members - cur_members)
            unchanged = list(cur_members & prev_members)

            if added or removed:
                changes_detected = True
            result[group] = {
                'added': sorted(added),
                'removed': sorted(removed),
                'unchanged': sorted(cur_members & prev_members)
            }
    return result, changes_detected, added_groups, deleted_groups

#-------------------Generate Report----------
def generate_html_report(snapshot, output_path, added_groups, deleted_groups):
    html = [
        "<html><head><meta charset='UTF-8'><style>",
        "body { font-family: Segoe UI, Arial, Helvetica Neue, sans-serif; }",
        "h2 { color: #333; }",
        ".added { color: green; }",
        ".removed { color: darkorange; }",
        ".unchanged { color: black; }",
        "table { border-collapse: collapse; width: 100%; }",
        "th, td { padding: 8px 12px; border: 1px solid #ccc; text-align: left; }",
        "</style></head><body>",
        "<h1>Critical Group Membership Report</h1>",
        f"<p>Report generated on: <strong>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</strong></p><br>"
        "<p>Total Groups Analyzed: " + str(len(snapshot)) + "</p><br>"
    ]

    # Add added and deleted groups
    if added_groups:
        html.append("<h2>Added Groups</h2>")
        for group in added_groups:
            html.append(f"<p>{group}</p>")
        html.append("<br>")

    if deleted_groups:
        html.append("<h2>Deleted Groups</h2>")
        for group in deleted_groups:
            html.append(f"<p>{group}</p>")
        html.append("<br>")

    # Separate changed and unchanged groups
    changed_groups = {}
    unchanged_groups = {}
    for group, data in snapshot.items():
        if data["added"] or data["removed"]:
            changed_groups[group] = data
        else:
            unchanged_groups[group] = data

    def append_group_section(groups):
        for group, data in groups:
            html.append(f"<h2>{group}</h2>")
            html.append("<table><tr><th>Change Type</th><th>Members</th></tr>")
            for change_type in ["added", "removed", "unchanged"]:
                class_name = change_type
                for member in data.get(change_type, []):
                    html.append(f"<tr><td class='{class_name}'>{change_type.capitalize()}</td><td class='{class_name}'>{member}</td></tr>")
            html.append("</table><br>")

    # Sort group names alphabetically
    changed_sorted = sorted(changed_groups.items())
    unchanged_sorted = sorted(unchanged_groups.items())

    # Display groups with changes first
    html.append("<h1>Groups With Changes</h1>")
    if changed_sorted:
        append_group_section(changed_sorted)
    else:
        html.append("<p>No changes detected in any group.</p>")

    # Then all groups sorted alphabetically
    html.append("<h1>All Groups</h1>")
    all_sorted_groups = changed_sorted + unchanged_sorted

    append_group_section(sorted(all_sorted_groups))

    html.append("</body></html>")

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(html))

#-------------------Generate PDF Report----------
def generate_pdf_report(html_path, pdf_path):
    pdfkit.from_file(html_path, pdf_path, options={'encoding': "UTF-8"})
    print(f"PDF report saved to: {pdf_path}")
   
#------------------Email report----------
def send_email(html_path, pdf_path):
    SENDER_EMAIL = os.environ.get('SENDER_EMAIL')
    RECIPIENT_EMAIL = os.environ.get('RECIPIENT_EMAIL')

    to_list = parse_recipients(RECIPIENT_EMAIL)
    if not to_list:
        raise ValueError(
            "No recipients provided. Set RECIPIENT_EMAIL or RECIPIENT_EMAILS (comma/semicolon/space/newline separated)."
        )

    token = get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    # Use /users/{sender}/sendMail if SENDER_EMAIL is set, otherwise /me/sendMail
    if SENDER_EMAIL:
        url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"
    else:
        url = "https://graph.microsoft.com/v1.0/me/sendMail"

    with open(html_path, 'r', encoding='utf-8') as f:
        html_content = f.read()
    with open(pdf_path, 'rb') as f:
        pdf_content = f.read()

    # Add your custom message at the top of the email body
    custom_message = """
    <p>Dear recipient,</p>
    <p>Please find the report for Critical Group Membership Changes . The PDF is attached for your convinience.</p>
    <p>Best regards,<br>IT Team</p>
    <hr>
    """
    full_html_body = custom_message 

    email_payload = {
        "message": {
            "subject": "Report: Critical Group Membership changes",
            "body": {
                "contentType": "HTML",
                "content": full_html_body
            },
            "toRecipients": to_recipient_objects(to_list),
            "attachments": [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": "Critical_groups_membership_report.pdf",
                "contentBytes": base64.b64encode(pdf_content).decode('utf-8')
            }]
        },
        "saveToSentItems": "true"
    }

    response = req.post(url, headers=headers, json=email_payload)
    try:
        response.raise_for_status()
        print("Email sent successfully")
    except Exception as e:
        print(f"Failed to send email: {response.text}")
        raise

# ------------------ Entry ------------------
def main():
    print(" Starting group snapshot comparison...")

    artifacts_dir = os.environ.get('BUILD_ARTIFACTSTAGINGDIRECTORY', './pipeline-artifacts')
    os.makedirs(artifacts_dir, exist_ok=True)

    current = get_all_group_members()
    previous = load_previous_snapshot()

    if not previous:
        print("No previous snapshot found. This is likely the first run.")
        print("Saving current snapshot for future comparison.")
        save_current_snapshot(current)
        return

    snapshot, changes_detected, added_groups, deleted_groups = compare_snapshots(current, previous)
    save_current_snapshot(current)

    print("Snapshot comparison complete.")

    # Save comparison result
    with open(os.path.join(artifacts_dir, 'comparison_result.json'), 'w', encoding='utf-8') as f:
        json.dump(snapshot, f, indent=2, ensure_ascii=False)

    # Generate HTML report
    html_report_path = os.path.join(artifacts_dir, 'Critical_groups_membership_report{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.html')
    generate_html_report(snapshot, html_report_path, added_groups, deleted_groups)
    print(f"HTML report saved to: {html_report_path}")

    # Generate PDF report
    pdf_report_path = os.path.join(artifacts_dir, 'Critical_groups_membership_report{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}.pdf')
    generate_pdf_report(html_report_path, pdf_report_path)
    
    # Send email
    send_email(html_report_path, pdf_report_path)

if __name__ == "__main__":
    main()
