import os
import json
import pdfkit
import base64
from datetime import datetime
import requests as req
from msal import ConfidentialClientApplication

# ------------------ Auth ------------------
filter_query = os.environ.get('GROUPS_FILTER')

def get_token():
    tenant_id     = os.environ.get('TENANT_ID')
    client_id     = os.environ.get('CLIENT_ID')
    client_secret = os.environ.get('CLIENT_SECRET')
    if not all([tenant_id, client_id, client_secret]):
        raise Exception("Missing environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET")
    app = ConfidentialClientApplication(
        client_id, client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}"
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        return result["access_token"]
    raise Exception(f"Token error: {result.get('error_description')}")

# ------------------ Delta helpers ------------------
def fetch_delta_pages(start_url, headers):
    """Follow @odata.nextLink pages and return (all_items, delta_link)."""
    items      = []
    delta_link = None
    url        = start_url
    while url:
        resp = req.get(url, headers=headers)
        resp.raise_for_status()
        data = resp.json()
        items.extend(data.get("value", []))
        delta_link = data.get("@odata.deltaLink")
        url        = data.get("@odata.nextLink")
    return items, delta_link

def member_label(m):
    """Best display string for a member object."""
    return m.get("displayName") or m.get("mail") or m.get("userPrincipalName", "")

# ------------------ State ------------------
def get_state_path():
    return os.path.join(
        os.environ.get("PIPELINE_WORKSPACE", "."),
        "group-report-artifacts",
        "delta_state.json"
    )

def load_state():
    path = get_state_path()
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            s = json.load(f)
        s.setdefault("groups_delta_link",   None)
        s.setdefault("current_groups",      [])
        s.setdefault("members_delta_links", {})
        print(f"Loaded state: {len(s['current_groups'])} known groups, "
              f"{len(s['members_delta_links'])} member delta links.")
        return s
    print("No previous state — first run.")
    return {"groups_delta_link": None, "current_groups": [], "members_delta_links": {}}

def save_state(state):
    artifacts_dir = os.environ.get('BUILD_ARTIFACTSTAGINGDIRECTORY', './pipeline-artifacts')
    os.makedirs(artifacts_dir, exist_ok=True)
    path = os.path.join(artifacts_dir, 'delta_state.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(state, f, indent=2)
    print(f"State saved: {path}")

# ------------------ Groups delta ------------------
def sync_groups_delta(headers, state):
    """
    Run the groups delta query.
      First run  : fetches all groups as a silent baseline → no new/deleted reported.
      Later runs : returns only groups added or deleted since the last run.
    Updates state in-place.
    Returns: (current_groups, new_group_names, deleted_group_names)
    """
    is_first = not state["groups_delta_link"]
    if is_first:
        url = "https://graph.microsoft.com/v1.0/groups/delta?$select=id,displayName,mail"
        if filter_query:
            url += f"&{filter_query}"
    else:
        url = state["groups_delta_link"]

    delta_items, delta_link   = fetch_delta_pages(url, headers)
    state["groups_delta_link"] = delta_link

    current_map   = {g["id"]: g for g in state["current_groups"]}
    new_names     = []
    deleted_names = []

    for g in delta_items:
        gid = g["id"]
        if "@removed" in g:
            deleted_names.append(current_map.get(gid, {}).get("displayName", gid))
            current_map.pop(gid, None)
            state["members_delta_links"].pop(gid, None)
        else:
            if not is_first and gid not in current_map:
                new_names.append(g.get("displayName", gid))
            current_map[gid] = g

    state["current_groups"] = list(current_map.values())
    return state["current_groups"], sorted(new_names), sorted(deleted_names)

# ------------------ Members ------------------
def get_current_members(group_id, headers):
    """Fetch the full current member list for a group (handles pagination)."""
    url     = (f"https://graph.microsoft.com/v1.0/groups/{group_id}/members"
               f"?$select=displayName,mail,userPrincipalName")
    members = []
    while url:
        resp = req.get(url, headers=headers)
        resp.raise_for_status()
        data = resp.json()
        members.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return sorted(member_label(m) for m in members if member_label(m))

def sync_members_delta(group_id, headers, state):
    """
    Run the member delta for one group.
      First run for this group : fetches all members as baseline → returns ([], []).
      Later runs               : returns (added_names, removed_names).
    Updates state in-place.
    """
    is_first = group_id not in state["members_delta_links"]
    url = (
        f"https://graph.microsoft.com/v1.0/groups/{group_id}/members/delta"
        f"?$select=displayName,mail,userPrincipalName"
        if is_first else state["members_delta_links"][group_id]
    )

    delta_items, delta_link = fetch_delta_pages(url, headers)
    if delta_link:
        state["members_delta_links"][group_id] = delta_link

    if is_first:
        print(f"  Baseline established ({len(delta_items)} members).")
        return [], []

    added, removed = [], []
    for item in delta_items:
        label = member_label(item)
        if not label:
            continue
        if "@removed" in item:
            removed.append(label)
        else:
            added.append(label)
    return sorted(added), sorted(removed)

# ------------------ Report ------------------
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
        "<h1>Azure AD Group Membership Report</h1>",
        f"<p>Report generated on: <strong>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</strong></p><br>"
    ]

    if added_groups:
        html.append("<h2>Added Groups</h2>")
        for g in added_groups:
            html.append(f"<p>{g}</p>")
        html.append("<br>")

    if deleted_groups:
        html.append("<h2>Deleted Groups</h2>")
        for g in deleted_groups:
            html.append(f"<p>{g}</p>")
        html.append("<br>")

    changed   = sorted((g, d) for g, d in snapshot.items() if d["added"] or d["removed"])
    unchanged = sorted((g, d) for g, d in snapshot.items() if not d["added"] and not d["removed"])

    def append_section(groups):
        for group, data in groups:
            html.append(f"<h2>{group}</h2>")
            html.append("<table><tr><th>Change Type</th><th>Members</th></tr>")
            for change_type in ["added", "removed", "unchanged"]:
                for member in data.get(change_type, []):
                    cls = change_type
                    html.append(
                        f"<tr><td class='{cls}'>{change_type.capitalize()}</td>"
                        f"<td class='{cls}'>{member}</td></tr>"
                    )
            html.append("</table><br>")

    html.append("<h1>Groups With Changes</h1>")
    if changed:
        append_section(changed)
    else:
        html.append("<p>No changes detected in any group.</p>")

    # Changed groups first (alphabetical), then unchanged (alphabetical)
    html.append("<h1>All Groups</h1>")
    append_section(changed + unchanged)

    html.append("</body></html>")
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(html))

def generate_pdf_report(html_path, pdf_path):
    pdfkit.from_file(html_path, pdf_path)
    print(f"PDF report saved to: {pdf_path}")

def send_email(pdf_path):
    sender    = os.environ.get('SENDER_EMAIL')
    recipient = os.environ.get('RECIPIENT_EMAIL')
    if not recipient:
        raise ValueError("RECIPIENT_EMAIL environment variable is not set.")

    token   = get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url     = (
        f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail"
        if sender else
        "https://graph.microsoft.com/v1.0/me/sendMail"
    )

    with open(pdf_path, 'rb') as f:
        pdf_b64 = base64.b64encode(f.read()).decode('utf-8')

    payload = {
        "message": {
            "subject": "Report: All AD Group members",
            "body": {
                "contentType": "HTML",
                "content": (
                    "<p>Dear recipient,</p>"
                    "<p>Please find the report for Azure AD Group Membership Changes. "
                    "The PDF version is attached for your convenience.</p>"
                    "<p>Best regards,<br>IT Team</p>"
                )
            },
            "toRecipients": [{"emailAddress": {"address": recipient}}],
            "attachments": [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": "group_membership_report.pdf",
                "contentBytes": pdf_b64
            }]
        },
        "saveToSentItems": "true"
    }

    resp = req.post(url, headers=headers, json=payload)
    try:
        resp.raise_for_status()
        print("Email sent successfully.")
    except Exception:
        print(f"Failed to send email: {resp.text}")
        raise

# ------------------ Entry ------------------
def main():
    print("Starting group delta report...")

    state   = load_state()
    token   = get_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    # Detect new / deleted groups via Graph groups delta
    current_groups, new_group_names, deleted_group_names = sync_groups_delta(headers, state)
    print(f"Groups — current: {len(current_groups)}, "
          f"new: {len(new_group_names)}, deleted: {len(deleted_group_names)}")

    new_group_set = set(new_group_names)
    snapshot      = {}

    for group in sorted(current_groups, key=lambda g: g.get("displayName", "").lower()):
        gid  = group["id"]
        name = group.get("displayName", gid)
        print(f"Processing: {name}")

        current_members = get_current_members(gid, headers)

        if name in new_group_set:
            # Newly created group — show every member as Added; establish delta baseline
            sync_members_delta(gid, headers, state)   # sets up the delta link
            added, removed = sorted(current_members), []
        else:
            added, removed = sync_members_delta(gid, headers, state)

        added_set = set(added)
        snapshot[name] = {
            "added":     added,
            "removed":   removed,
            "unchanged": sorted(m for m in current_members if m not in added_set)
        }

    save_state(state)

    artifacts_dir = os.environ.get('BUILD_ARTIFACTSTAGINGDIRECTORY', './pipeline-artifacts')
    os.makedirs(artifacts_dir, exist_ok=True)

    html_path = os.path.join(artifacts_dir, 'group_membership_report.html')
    generate_html_report(snapshot, html_path, new_group_names, deleted_group_names)
    print(f"HTML report: {html_path}")

    pdf_path = os.path.join(artifacts_dir, 'group_membership_report.pdf')
    generate_pdf_report(html_path, pdf_path)

    send_email(pdf_path)
    print("Done.")

if __name__ == "__main__":
    main()
