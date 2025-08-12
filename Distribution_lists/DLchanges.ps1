param (
    [string]$appId,
    [string]$orgName,
    [string]$thumbprint,
    [string]$previous = "previousMembers.json",
    [string]$report = "DLchanges_report.html"
)

Write-Host "Connecting to Exchange Online..."
Connect-ExchangeOnline -AppId $appId -Organization $orgName -CertificateThumbprint $thumbprint

Write-Host "Fetching Distribution Lists..."

#Fetch distribution lists and sort by display name
$distributionLists = Get-DistributionGroup | Sort-Object DisplayName

# Fetch current members for each distribution list using hash table
$currentMembers = @{}
foreach ($distributionList in $distributionLists) {
    try {
        $members = Get-DistributionGroupMember -Identity $distributionList.PrimarySmtpAddress | Select-Object -ExpandProperty PrimarySmtpAddress
    } catch {
        Write-Warning "Unable to fetch members for $($distributionList.DisplayName): $_"
        $members = @()
    }
    $currentMembers[$distributionList.DisplayName] = $members
}

# Load previous state
# Load previous state
Write-Host "Looking for previous state at path: $previous"

$oldmembers = @{}
if (Test-Path $previous) {
    Write-Host "Loading previous report from $previous"
    $json = Get-Content $previous -Raw
    $converted = $json | ConvertFrom-Json

    # Convert PSCustomObject to hashtable
    $oldmembers = @{}
    foreach ($entry in $converted.PSObject.Properties) {
        $oldmembers[$entry.Name] = $entry.Value
    }

    Write-Host "Loaded $($oldmembers.Keys.Count) groups from previous state."
} else {
    Write-Host "No previous report found — creating baseline."
}
Write-Host "Previous state keys:"
$oldmembers.Keys | ForEach-Object { Write-Host "- $_" }

# Categorize groups
$newGroups = @()
$deletedGroups = @()
$groupsWithChanges = @{}
$allGroupsTable = @{}

$allGroups = $currentMembers.Keys + $oldmembers.Keys | Sort-Object -Unique

foreach ($group in $allGroups) {
    $current = $currentMembers[$group]
    $old = $oldmembers[$group]

    if ($null -eq $old) { #checks if DL is new and adds it to newGroups
        $newGroups += $group
        $groupsWithChanges[$group] = @()
        if ($null -ne $current) {
            foreach ($user in $current) {
                $groupsWithChanges[$group] += @{ Type = 'Added'; User = $user }
            }
        }
    }
    elseif ($null -eq $current) { #checks if DL is deleted and adds it to deletedGroups
        $deletedGroups += $group
        $groupsWithChanges[$group] = @()
        if ($null -ne $old) {
            foreach ($user in $old) {
                $groupsWithChanges[$group] += @{ Type = 'Removed'; User = $user }
            }
        }
    }
    else { #checks if DL has changes in members and adds it to groupsWithChanges
        $added = Compare-Object -ReferenceObject $old -DifferenceObject $current -PassThru | Where-Object { $_ -in $current }
        $removed = Compare-Object -ReferenceObject $old -DifferenceObject $current -PassThru | Where-Object { $_ -in $old }

        if ($added.Count -gt 0 -or $removed.Count -gt 0) {
            $groupsWithChanges[$group] = @()
            foreach ($user in $added) {
                $groupsWithChanges[$group] += @{ Type = 'Added'; User = $user }
            }
            foreach ($user in $removed) {
                $groupsWithChanges[$group] += @{ Type = 'Removed'; User = $user }
            }
        }
    }

    # All groups section
    $allGroupsTable[$group] = @()
    if ($null -ne $current) {
        foreach ($user in $current) {
            $status = 'Unchanged'
            if ($groupsWithChanges.ContainsKey($group)) {
                $change = $groupsWithChanges[$group] | Where-Object { $_.User -eq $user }
                if ($change) {
                    $status = $change.Type
                }
            }
            $allGroupsTable[$group] += @{ Type = $status; User = $user }
        }
    }
}

# Generate HTML
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
<h1>Distribution List Membership Report - $(Get-Date -Format "yyyy-MM-dd")</h1>
"@

# Section: New DLs
$html += '<h2> New Distribution Lists</h2><ul>'
foreach ($g in $newGroups) { $html += "<li>$g</li>" }
$html += '</ul>'

# Section: Deleted DLs
$html += '<h2> Deleted Distribution Lists</h2><ul>'
foreach ($g in $deletedGroups) { $html += "<li>$g</li>" }
$html += '</ul>'

# Section: Groups with Changes
$html += '<h2>Groups With Changes</h2>'
foreach ($group in $groupsWithChanges.Keys) {
    $html += "<h3>$group</h3><table><tr><th>Change Type</th><th>Member</th></tr>"
    foreach ($entry in $groupsWithChanges[$group]) {
        $html += "<tr><td class='$($entry.Type.ToLower())'>$($entry.Type)</td><td class='$($entry.Type.ToLower())'>$($entry.User)</td></tr>"
    }
    $html += '</table>'
}

# Section: All Groups
$html += '<h2>All Groups</h2>'
foreach ($group in $allGroupsTable.Keys | Sort-Object) {
    $html += "<h3>$group</h3><table><tr><th>Change Type</th><th>Member</th></tr>"
    foreach ($entry in $allGroupsTable[$group]) {
        $html += "<tr><td class='$($entry.Type.ToLower())'>$($entry.Type)</td><td>$($entry.User)</td></tr>"
    }
    $html += '</table>'
}

$html += '</body></html>'

# Save report
Write-Host "Saving report to $report"
$html | Out-File -Encoding utf8 $report

# Save current state
Write-Host "Saving current DL state to $previous"
$currentMembers | ConvertTo-Json -Depth 5 | Out-File $previous

Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Done."
