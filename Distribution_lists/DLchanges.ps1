param (
    [string]$appId,
    [string]$orgName,
    [string]$thumbprint,
    [string]$previous,
    [string]$report
)

Write-Host "Connecting to Exchange Online..."
Connect-ExchangeOnline -AppId $appId -Organization $orgName -CertificateThumbprint $thumbprint

Write-Host "Fetching Distribution Lists..."
$distributionLists = Get-DistributionGroup | Sort-Object DisplayName
Write-Host "Total distribution lists: $($distributionLists.Count)"

# Build current state: Hashtable keyed by DisplayName with array of PrimarySmtpAddress members
$currentMembers = @{}
foreach ($distributionList in $distributionLists) {
    try {
        $members = Get-DistributionGroupMember -Identity $distributionList.PrimarySmtpAddress |
                   Select-Object -ExpandProperty PrimarySmtpAddress
        if (-not $members) { $members = @() }  # ensure array, not $null
    } catch {
        Write-Warning "Unable to fetch members for $($distributionList.DisplayName): $_"
        $members = @()
    }
    $currentMembers[$distributionList.DisplayName] = $members
}

# Load previous state
$oldmembers = @{}
if (Test-Path $previous) {
    Write-Host "Loading previous state from $previous"
    $json = Get-Content $previous -Raw
    $converted = $json | ConvertFrom-Json

    foreach ($entry in $converted.PSObject.Properties) {
        # Normalize to array (JSON may deserialize singletons differently)
        $val = $entry.Value
        if ($null -eq $val) { $val = @() }
        elseif ($val -isnot [System.Collections.IEnumerable] -or $val -is [string]) { $val = @($val) }
        $oldmembers[$entry.Name] = $val
    }
    Write-Host "Loaded $($oldmembers.Keys.Count) groups from previous state."
} else {
    Write-Host "No previous state found. Using empty baseline."
}

# ---------- New logic: decide new/deleted from NAMES only ----------
$newGroups = @()
$deletedGroups = @()
$groupsWithChanges = @{}
$allGroupsTable = @{}

$currentGroupNames = @($currentMembers.Keys)
$oldGroupNames     = @($oldmembers.Keys)

# New groups = in current, not in old (by name)
$newGroups = $currentGroupNames | Where-Object { $_ -notin $oldGroupNames } | Sort-Object

# Deleted groups = in old, not in current (by name)
$deletedGroups = $oldGroupNames | Where-Object { $_ -notin $currentGroupNames } | Sort-Object

# Common groups = present in both snapshots (by name)
$commonGroups = $currentGroupNames | Where-Object { $_ -in $oldGroupNames } | Sort-Object

# Record changes for NEW groups: all current members are "Added"
foreach ($g in $newGroups) {
    $groupsWithChanges[$g] = @()
    foreach ($user in ($currentMembers[$g] | ForEach-Object { $_ } )) {
        $groupsWithChanges[$g] += @{ Type = 'Added'; User = $user }
    }
}

# Record changes for DELETED groups: all old members are "Removed"
foreach ($g in $deletedGroups) {
    $groupsWithChanges[$g] = @()
    foreach ($user in ($oldmembers[$g] | ForEach-Object { $_ } )) {
        $groupsWithChanges[$g] += @{ Type = 'Removed'; User = $user }
    }
}

# For COMMON groups, compare members only (added/removed)
foreach ($g in $commonGroups) {
    $curr = $currentMembers[$g]; if (-not $curr) { $curr = @() }
    $old  = $oldmembers[$g];     if (-not $old)  { $old  = @() }

    $added   = Compare-Object -ReferenceObject $old -DifferenceObject $curr -PassThru | Where-Object { $_ -in $curr }
    $removed = Compare-Object -ReferenceObject $old -DifferenceObject $curr -PassThru | Where-Object { $_ -in $old  }

    if (($added | Measure-Object).Count -gt 0 -or ($removed | Measure-Object).Count -gt 0) {
        $groupsWithChanges[$g] = @()
        foreach ($u in $added)   { $groupsWithChanges[$g] += @{ Type = 'Added';   User = $u } }
        foreach ($u in $removed) { $groupsWithChanges[$g] += @{ Type = 'Removed'; User = $u } }
    }
}

# --------- Build "All Groups" table (alphabetical) ----------
$allGroupsByName = @($currentGroupNames + $deletedGroups | Sort-Object -Unique)  # include deleted so they show once with Removed entries
foreach ($g in $allGroupsByName) {
    $allGroupsTable[$g] = @()
    $listToShow = $currentMembers.ContainsKey($g) ? $currentMembers[$g] : @()  # if deleted, show nothing under All Groups (optional)
    foreach ($user in $listToShow) {
        $status = 'Unchanged'
        if ($groupsWithChanges.ContainsKey($g)) {
            $change = $groupsWithChanges[$g] | Where-Object { $_.User -eq $user }
            if ($change) { $status = $change.Type }
        }
        $allGroupsTable[$g] += @{ Type = $status; User = $user }
    }
}

# ---------- Build HTML report ----------
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
</style>
</head>
<body>
<h1>Distribution List Membership Report - $(Get-Date -Format 'yyyy-MM-dd')</h1>
"@

$html += "<h2>New Distribution Lists</h2><ul>"
foreach ($g in $newGroups) { $html += "<li>$g</li>" }
$html += "</ul>"

$html += "<h2>Deleted Distribution Lists</h2><ul>"
foreach ($g in $deletedGroups) { $html += "<li>$g</li>" }
$html += "</ul>"

$html += "<h2>Groups With Changes</h2>"
foreach ($group in ($groupsWithChanges.Keys | Sort-Object)) {
    $html += "<h3>$group</h3><table><tr><th>Change Type</th><th>Member</th></tr>"
    foreach ($entry in $groupsWithChanges[$group]) {
        $html += "<tr><td class='$($entry.Type.ToLower())'>$($entry.Type)</td><td class='$($entry.Type.ToLower())'>$($entry.User)</td></tr>"
    }
    $html += "</table>"
}

$html += "<h2>All Groups</h2>"
foreach ($group in ($allGroupsTable.Keys | Sort-Object)) {
    $html += "<h3>$group</h3><table><tr><th>Change Type</th><th>Member</th></tr>"
    foreach ($entry in $allGroupsTable[$group]) {
        $html += "<tr><td class='$($entry.Type.ToLower())'>$($entry.Type)</td><td>$($entry.User)</td></tr>"
    }
    $html += "</table>"
}

$html += "</body></html>"

# Save report + current snapshot
Write-Host "Saving report to $report"
$html | Out-File -Encoding utf8 $report

Write-Host "Saving current DL state to $previous"
$currentMembers | ConvertTo-Json -Depth 5 | Out-File $previous

Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Done."
