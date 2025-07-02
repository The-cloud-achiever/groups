import os
import json
from datetime import datetime
import requests as req
from msal import ConfidentialClientApplication

# ------------------ Authentication ------------------
filter_query = os.environ.get('GROUPS_FILTER')
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

# ------------------ Group Fetching ------------------
def get_groups():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    if filter_query:
        url = f"https://graph.microsoft.com/v1.0/groups?{filter_query}"
    else:
        url = "https://graph.microsoft.com/v1.0/groups"

    response = req.get(url, headers=headers)
    response.raise_for_status()
    return response.json().get("value", [])

def get_all_group_members():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    groups = get_groups()
    group_members = {}
    batch_requests = []
    batch_size = 20
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
            response = req.post("https://graph.microsoft.com/v1.0/$batch", headers=headers, json=batch_payload)
            response.raise_for_status()

            results = response.json()["responses"]
            for result in results:
                idx = int(result["id"]) - 1
                group_name = groups[idx]["displayName"]
                if result["status"] == 200:
                    members = [m["displayName"] for m in result["body"].get("value", [])]
                    group_members[group_name] = members
            batch_requests = []

    return group_members

# ------------------ Snapshot Handling ------------------
def load_previous_snapshot():
    path = os.path.join(os.environ.get("PIPELINE_WORKSPACE", "./"),"group-report-artifacts","previous_snapshot.json")
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    print("No previous snapshot found, treating this as first run.")
    return {}

def save_current_snapshot(data):
    artifacts_dir = os.environ.get('BUILD_ARTIFACTSTAGINGDIRECTORY', './pipeline-artifacts')
    os.makedirs(artifacts_dir, exist_ok=True)
    with open(os.path.join(artifacts_dir, 'previous_snapshot.json'), 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)

# ------------------ Comparison Logic ------------------
def compare_snapshots(current, previous):
    result = {}
    all_keys = set(current.keys()).union(previous.keys())
    changes_detected = False

    for group in all_keys:
        cur = set(current.get(group, []))
        prev = set(previous.get(group, []))
        added = list(cur - prev)
        removed = list(prev - cur)

        if added or removed:
            changes_detected = True
        result[group] = {
            'added': sorted(added),
            'removed': sorted(removed),
            'unchanged': sorted(cur & prev)
        }

    return result, changes_detected
#-------------------Generate Report----------
def generate_html_report(snapshot, output_path):
    html = [
        "<html><head><style>",
        "body { font-family: Arial, sans-serif; }",
        "h2 { color: #333; }",
        ".added { color: green; }",
        ".removed { color: darkorange; }",
        ".unchanged { color: black; }",
        "table { border-collapse: collapse; width: 100%; }",
        "th, td { padding: 8px 12px; border: 1px solid #ccc; text-align: left; }",
        "</style></head><body>",
        "<h1>Azure AD Group Membership Report</h1>"
    ]

    # Separate changed and unchanged groups
    changed_groups = {}
    unchanged_groups = {}
    for group, data in snapshot.items():
        if data["added"] or data["removed"]:
            changed_groups[group] = data
        else:
            unchanged_groups[group] = data

    # Sort group names alphabetically
    
    
    def append_group_section(groups):
        for group, data in groups:
            html.append(f"<h2>{group}</h2>")
            html.append("<table><tr><th>Change Type</th><th>Members</th></tr>")
            for change_type in ["added", "removed", "unchanged"]:
                class_name = change_type
                for member in data.get(change_type, []):
                    html.append(f"<tr><td class='{class_name}'>{change_type.capitalize()}</td><td class='{class_name}'>{member}</td></tr>")
            html.append("</table><br>")

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
    append_group_section(changed_sorted + unchanged_sorted)

    html.append("</body></html>")

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(html))

    

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

    snapshot, changes = compare_snapshots(current, previous)
    save_current_snapshot(current)

    print("Snapshot comparison complete.")

    # Save comparison result
    with open(os.path.join(artifacts_dir, 'comparison_result.json'), 'w', encoding='utf-8') as f:
        json.dump(snapshot, f, indent=2)

    # Generate HTML report
    html_report_path = os.path.join(artifacts_dir, 'group_membership_report.html')
    generate_html_report(snapshot, html_report_path)
    print(f"HTML report saved to: {html_report_path}")

if __name__ == "__main__":
    main()
