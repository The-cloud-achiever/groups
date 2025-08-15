param (
    [string]$appId,
    [string]$orgName,
    [string]$thumbprint,
    [string]$previous,
    [string]$report,

    # Email
    [string]$mailFrom,
    [string]$mailTo,
    [string]$mailSubject = "Distribution List Membership Report",

    # Graph (tenant)
    [string]$tenantId
)

# ---------- Helpers ----------
function AsStringArray {
    param($InputValue)
    $out = @()
    if ($null -eq $InputValue) { return @() }
    foreach ($i in @($InputValue)) {
        if ($null -eq $i) { continue }
        if ($i -is [psobject]) {
            if     ($i.PSObject.Properties['User'])                 { $s = [string]$i.User }
            elseif ($i.PSObject.Properties['PrimarySmtpAddress'])   { $s = [string]$i.PrimarySmtpAddress }
            elseif ($i.PSObject.Properties['Value'])                { $s = [string]$i.Value }
            elseif ($i.PSObject.Properties['InputObject'])          { $s = [string]$i.InputObject }
            else                                                    { $s = [string]$i }
        } else {
            $s = [string]$i
        }
        if ($s) { $out += $s.Trim() }
    }
    return $out
}

# ---------- Connect EXO ----------
Write-Host "Connecting to Exchange Online..."
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -AppId $appId -Organization $orgName -CertificateThumbprint $thumbprint

# ---------- Gather current state ----------
Write-Host "Fetching Distribution Lists..."
$distributionLists = Get-DistributionGroup | Sort-Object DisplayName
Write-Host "Total distribution lists: $($distributionLists.Count)"

$currentMembers = @{}
foreach ($distributionList in $distributionLists) {
    $display = $distributionList.DisplayName
    if (-not $display) { continue }
    try {
        $members = Get-DistributionGroupMember -Identity $distributionList.PrimarySmtpAddress |
                   Select-Object -ExpandProperty PrimarySmtpAddress
        $currentMembers[$display] = AsStringArray $members
    } catch {
        Write-Warning "Unable to fetch members for $display : $_"
        $currentMembers[$display] = @()
    }
}

# ---------- Load previous snapshot ----------
$oldmembers = @{}
if (Test-Path $previous) {
    Write-Host "Loading previous state from $previous"
    $json = Get-Content $previous -Raw
    $converted = if ($json) { $json | ConvertFrom-Json } else { $null }
    if ($converted) {
        foreach ($entry in $converted.PSObject.Properties) {
            $oldmembers[$entry.Name] = AsStringArray $entry.Value
        }
        Write-Host "Loaded $($oldmembers.Keys.Count) groups from previous state."
    } else {
        Write-Host "Previous file empty or invalid JSON; using empty baseline."
    }
} else {
    Write-Host "No previous state found. Using empty baseline."
}

# ---------- Compare by GROUP NAMES ----------
$currentGroupNames = @($currentMembers.Keys)
$oldGroupNames     = @($oldmembers.Keys)

$newGroups     = $currentGroupNames | Where-Object { $_ -notin $oldGroupNames } | Sort-Object
$deletedGroups = $oldGroupNames     | Where-Object { $_ -notin $currentGroupNames } | Sort-Object
$commonGroups  = $currentGroupNames | Where-Object { $_ -in $oldGroupNames } | Sort-Object

$groupsWithChanges = @{}
$allGroupsTable    = @{}

# New groups -> Added
foreach ($g in $newGroups) {
    $groupsWithChanges[$g] = @()
    foreach ($user in $currentMembers[$g]) {
        $groupsWithChanges[$g] += @{ Type = 'Added'; User = $user }
    }
}

# Deleted groups -> Removed
foreach ($g in $deletedGroups) {
    $groupsWithChanges[$g] = @()
    foreach ($user in $oldmembers[$g]) {
        $groupsWithChanges[$g] += @{ Type = 'Removed'; User = $user }
    }
}

# Common groups -> Compare-Object (safe)
foreach ($g in $commonGroups) {
    $curr = AsStringArray $currentMembers[$g]
    $old  = AsStringArray $oldmembers[$g]
    if ($null -eq $curr) { $curr = @() }
    if ($null -eq $old)  { $old  = @() }

    $diff    = Compare-Object -ReferenceObject $old -DifferenceObject $curr
    $added   = @($diff | Where-Object SideIndicator -eq '=>' | Select-Object -ExpandProperty InputObject | ForEach-Object { ([string]$_).Trim() })
    $removed = @($diff | Where-Object SideIndicator -eq '<=' | Select-Object -ExpandProperty InputObject | ForEach-Object { ([string]$_).Trim() })

    if ($added.Count -gt 0 -or $removed.Count -gt 0) {
        $groupsWithChanges[$g] = @()
        foreach ($u in $added)   { if ($u) { $groupsWithChanges[$g] += @{ Type = 'Added';   User = $u } } }
        foreach ($u in $removed) { if ($u) { $groupsWithChanges[$g] += @{ Type = 'Removed'; User = $u } } }
    }
}

# All Groups for the report (current groups only)
foreach ($g in ($currentGroupNames | Sort-Object)) {
    $allGroupsTable[$g] = @()
    foreach ($user in $currentMembers[$g]) {
        $status = 'Unchanged'
        if ($groupsWithChanges.ContainsKey($g)) {
            $change = $groupsWithChanges[$g] | Where-Object { $_.User -eq $user }
            if ($change) { $status = $change.Type }
        }
        $allGroupsTable[$g] += @{ Type = $status; User = $user }
    }
}

# ---------- HTML ----------
$html = @"
<html>
<head>
<style>
    body { font-family: Arial; }
    .added { color: green; }
    .removed { color: darkorange; }
    .unchanged { color: black; }
    table { border-collapse: collapse; width: 100%; }
    th, td { padding: 8px 12px; border: 1px solid #ccc; text-align: left; }
    th { background-color: #eee; }
    h3 { margin-top: 20px; }
</style>
</head>
<body>
<h1>Distribution List Membership Report - $(Get-Date -Format 'yyyy-MM-dd')</h1>
"@

$html += "<h2>New Distribution Lists</h2><ul>"
foreach ($g in $newGroups) { $html += "<li>$g</li>" }
if (-not $newGroups) { $html += "<li><em>None</em></li>" }
$html += "</ul>"

$html += "<h2>Deleted Distribution Lists</h2><ul>"
foreach ($g in $deletedGroups) { $html += "<li>$g</li>" }
if (-not $deletedGroups) { $html += "<li><em>None</em></li>" }
$html += "</ul>"

$html += "<h2>Groups With Changes</h2>"
if ($groupsWithChanges.Keys.Count -eq 0) {
    $html += "<p><em>No changes detected.</em></p>"
} else {
    foreach ($group in ($groupsWithChanges.Keys | Sort-Object)) {
        $html += "<h3>$group</h3><table><tr><th>Change Type</th><th>Member</th></tr>"
        foreach ($entry in $groupsWithChanges[$group]) {
            $cls = $entry.Type.ToLower()
            $usr = [System.Web.HttpUtility]::HtmlEncode($entry.User)
            $html += "<tr><td class='$cls'>$($entry.Type)</td><td class='$cls'>$usr</td></tr>"
        }
        $html += "</table>"
    }
}

$html += "<h2>All Groups</h2>"
foreach ($group in ($allGroupsTable.Keys | Sort-Object)) {
    $html += "<h3>$group</h3><table><tr><th>Change Type</th><th>Member</th></tr>"
    foreach ($entry in $allGroupsTable[$group]) {
        $cls = $entry.Type.ToLower()
        $usr = [System.Web.HttpUtility]::HtmlEncode($entry.User)
        $html += "<tr><td class='$cls'>$($entry.Type)</td><td>$usr</td></tr>"
    }
    $html += "</table>"
}
$html += "</body></html>"

# ---------- Save report + snapshot ----------
Write-Host "Saving report to $report"
$html | Out-File -Encoding utf8 $report

Write-Host "Saving current DL state to $previous"
$currentMembers | ConvertTo-Json -Depth 5 | Out-File $previous

# ---------- Email via Microsoft Graph (cert-based, no password) ----------
# Fallback to env if not provided via param
if (-not $tenantId) { $tenantId = $env:TENANT_ID }

function NormalizeAddresses {
    param([string]$raw)
    if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
    $parts = $raw -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $emails = @()
    foreach ($p in $parts) {
        # handle "Display Name <user@domain>" or just "user@domain"
        if ($p -match '<([^>]+)>') { $addr = $matches[1].Trim() } else { $addr = $p }
        # very light validation
        if ($addr -match '^[^\s@]+@[^\s@]+\.[^\s@]+$') { $emails += $addr }
    }
    return $emails
}

try {
    Import-Module Microsoft.Graph.Users.Actions -ErrorAction Stop
    Connect-MgGraph -TenantId $tenantId -ClientId $appId -CertificateThumbprint $thumbprint -NoWelcome

    $fromAddr = (NormalizeAddresses $mailFrom) | Select-Object -First 1
    $toAddrs  = NormalizeAddresses $mailTo

    if (-not $fromAddr) { throw "MAIL_FROM is empty or invalid." }
    if (-not $toAddrs -or $toAddrs.Count -eq 0) { throw "MAIL_TO is empty or invalid after normalization." }

    # (Optional) brief log to help debug (no full emails printed)
    Write-Host "Email from: $($fromAddr)"
    Write-Host "Email TO: $($toAddrs)"
    Write-Host "Email to count: $($toAddrs.Count)"

    $bodyHtml = Get-Content $report -Raw
    $bytes    = [IO.File]::ReadAllBytes($report)
    $b64      = [Convert]::ToBase64String($bytes)

    $message = @{
        subject = $mailSubject
        body    = @{ contentType = "HTML"; content = $bodyHtml }
        toRecipients = $toAddrs | ForEach-Object { @{ emailAddress = @{ address = $_ } } }
        attachments = @(
            @{
                '@odata.type' = '#microsoft.graph.fileAttachment'
                name          = (Split-Path $report -Leaf)
                contentType   = 'text/html'
                contentBytes  = $b64
            }
        )
        from = @{ emailAddress = @{ address = $fromAddr } }
    }

    Send-MgUserMail -UserId $fromAddr -Message $message -SaveToSentItems
    Write-Host "Graph email sent to $($toAddrs -join ', ')"
}
catch {
    Write-Warning "Failed to send via Graph: $($_.Exception.Message)"
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Done."
}