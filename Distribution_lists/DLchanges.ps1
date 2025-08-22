param (
    [Parameter(Mandatory)] [string]$AppId,
    [Parameter(Mandatory)] [string]$TenantId,
    [Parameter(Mandatory)] [string]$OrgName,
    [Parameter(Mandatory)] [string]$Thumbprint,
    [Parameter(Mandatory)] [string]$Previous,
    [Parameter(Mandatory)] [string]$Report,
    [Parameter(Mandatory)] [string]$MailFrom,
    [Parameter(Mandatory)] [string]$MailTo,
    [string]$MailSubject = "Distribution List Report"
)

# Helper: always return a string[] from any input shape (never $null)
function AsStringArray {
    param($InputValue)
    $out = @()
    if ($null -eq $InputValue) { return @() }
    foreach ($i in @($InputValue)) {
        if ($null -eq $i) { continue }
        if ($i -is [psobject]) {
            if ($i.PSObject.Properties['User'])                  { $s = [string]$i.User }
            elseif ($i.PSObject.Properties['PrimarySmtpAddress']){ $s = [string]$i.PrimarySmtpAddress }
            elseif ($i.PSObject.Properties['Value'])             { $s = [string]$i.Value }
            elseif ($i.PSObject.Properties['InputObject'])       { $s = [string]$i.InputObject }
            else                                                 { $s = [string]$i }
        } else {
            $s = [string]$i
        }
        if ($s) { $out += $s.Trim() }
    }
    return $out
}

Write-Host "Connecting to Exchange Online..."
Connect-ExchangeOnline -AppId $AppId -Organization $OrgName -CertificateThumbprint $Thumbprint -ShowBanner:$false

Write-Host "Connecting to Microsoft Graph (certificate auth)..."
# Requires application permissions: Mail.Send
Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome

Write-Host "Fetching Distribution Lists..."
$distributionLists = Get-DistributionGroup | Sort-Object DisplayName
Write-Host "Total distribution lists: $($distributionLists.Count)"

# Current snapshot: DisplayName -> string[] of member SMTPs
$currentMembers = @{}
foreach ($distributionList in $distributionLists) {
    $display = $distributionList.DisplayName
    if (-not $display) { continue }
    try {
        $members = Get-DistributionGroupMember -Identity $distributionList.PrimarySmtpAddress `
                 | Select-Object -ExpandProperty PrimarySmtpAddress
        $currentMembers[$display] = AsStringArray $members
    } catch {
        Write-Warning "Unable to fetch members for $display : $_"
        $currentMembers[$display] = @()
    }
}

# Previous snapshot
$oldmembers = @{}
if (Test-Path $Previous) {
    Write-Host "Loading previous state from $Previous"
    $json = Get-Content $Previous -Raw
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

# New/Deleted/Common by NAME only
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

# Common groups: safe Compare-Object
# Common groups: null-safe Compare-Object
foreach ($g in $commonGroups) {

    # fetch raw values if present
    $currSrc = $null
    if ($currentMembers.ContainsKey($g)) { $currSrc = $currentMembers[$g] }

    $oldSrc  = $null
    if ($oldmembers.ContainsKey($g))     { $oldSrc  = $oldmembers[$g] }

    # normalize to string[]
    $curr = AsStringArray $currSrc
    $old  = AsStringArray $oldSrc

    # harden against $null (AsStringArray *should* return @(), but be explicit)
    if ($null -eq $curr) { $curr = @() }
    if ($null -eq $old)  { $old  = @() }

    # nothing to compare? next group
    if ($curr.Count -eq 0 -and $old.Count -eq 0) { continue }

    $diff    = Compare-Object -ReferenceObject $old -DifferenceObject $curr
    $added   = @($diff | Where-Object SideIndicator -eq '=>' | Select-Object -ExpandProperty InputObject | ForEach-Object { ([string]$_).Trim() })
    $removed = @($diff | Where-Object SideIndicator -eq '<=' | Select-Object -ExpandProperty InputObject | ForEach-Object { ([string]$_).Trim() })

    if ($added.Count -gt 0 -or $removed.Count -gt 0) {
        $groupsWithChanges[$g] = @()
        foreach ($u in $added)   { if ($u) { $groupsWithChanges[$g] += @{ Type = 'Added';   User = $u } } }
        foreach ($u in $removed) { if ($u) { $groupsWithChanges[$g] += @{ Type = 'Removed'; User = $u } } }
    }
}


# All Groups table (current groups only, alphabetical)
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
    body { font-family: Arial, sans-serif; }
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
foreach ($g in $newGroups) { $html += "<li>$([System.Web.HttpUtility]::HtmlEncode($g))</li>" }
if (-not $newGroups) { $html += "<li><em>None</em></li>" }
$html += "</ul>"

$html += "<h2>Deleted Distribution Lists</h2><ul>"
foreach ($g in $deletedGroups) { $html += "<li>$([System.Web.HttpUtility]::HtmlEncode($g))</li>" }
if (-not $deletedGroups) { $html += "<li><em>None</em></li>" }
$html += "</ul>"

$html += "<h2>Groups With Changes</h2>"
if ($groupsWithChanges.Keys.Count -eq 0) {
    $html += "<p><em>No changes detected.</em></p>"
} else {
    foreach ($group in ($groupsWithChanges.Keys | Sort-Object)) {
        $html += "<h3>$([System.Web.HttpUtility]::HtmlEncode($group))</h3><table><tr><th>Change Type</th><th>Member</th></tr>"
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
    $html += "<h3>$([System.Web.HttpUtility]::HtmlEncode($group))</h3><table><tr><th>Change Type</th><th>Member</th></tr>"
    foreach ($entry in $allGroupsTable[$group]) {
        $cls = $entry.Type.ToLower()
        $usr = [System.Web.HttpUtility]::HtmlEncode($entry.User)
        $html += "<tr><td class='$cls'>$($entry.Type)</td><td>$usr</td></tr>"
    }
    $html += "</table>"
}

$html += "</body></html>"

Write-Host "Saving report to $Report"
$html | Out-File -Encoding utf8 $Report

Write-Host "Saving current DL state to $Previous"
$currentMembers | ConvertTo-Json -Depth 5 | Out-File $Previous -Encoding utf8

function Send-ReportEmail {
    param(
        [Parameter(Mandatory)] [string]$From,          # UPN or userId of the mailbox to send as
        [Parameter(Mandatory)] [string]$To,            # one or many, comma/semicolon separated
        [Parameter(Mandatory)] [string]$Subject,
        [Parameter(Mandatory)] [string]$AttachmentPath
    )

    if ([string]::IsNullOrWhiteSpace($From)) { Write-Warning "MAIL_FROM is empty"; return }
    if ([string]::IsNullOrWhiteSpace($To))   { Write-Warning "MAIL_TO is empty";   return }
    if (-not (Test-Path $AttachmentPath))    { Write-Warning "Attachment not found: $AttachmentPath"; return }

    # Build recipients (supports comma/semicolon lists)
    $recips = @()
    foreach ($addr in ($To -split '[;,]')) {
        $a = $addr.Trim()
        if ($a) { $recips += @{ emailAddress = @{ address = $a } } }
    }
    if ($recips.Count -eq 0) { Write-Warning "No valid recipients parsed from MAIL_TO."; return }

    # Build the file attachment
    $attachment = @{
        '@odata.type' = '#microsoft.graph.fileAttachment'
        name          = [IO.Path]::GetFileName($AttachmentPath)
        contentBytes  = [Convert]::ToBase64String([IO.File]::ReadAllBytes($AttachmentPath))
        contentType   = 'text/html'
    }

    # Message object for the correct parameter set
    $message = @{
        subject      = $Subject
        body         = @{ contentType = 'HTML'; content = 'Please find the attached Distribution List report.' }
        toRecipients = $recips
        attachments  = @($attachment)
    }

    try {
        # IMPORTANT: Use -Message (or -BodyParameter). Without this, parameter binding fails.
        Send-MgUserMail -UserId $From -Message $message -SaveToSentItems -ErrorAction Stop
        Write-Host "Email sent to $($recips.Count) recipient(s) from $From"
    }
    catch {
        Write-Warning "Failed to send email: $($_.Exception.Message)"
    }
}


Write-Host "Sending report email to $MailTo"
Send-ReportEmail -From $MailFrom -To $MailTo -Subject $MailSubject -AttachmentPath $Report


Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph
Write-Host "Done."
