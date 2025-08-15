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
$oldmembers = @{}
if (Test-Path $previous) {
    Write-Host "Loading previous state from $previous"
    $json = Get-Content $previous -Raw
    $converted = $json | ConvertFrom-Json

    foreach ($entry in $converted.PSObject.Properties) {
        $oldmembers[$entry.Name] = $entry.Value
    }

    Write-Host "Loaded $($oldmembers.Keys.Count) groups from previous state."
} else {
    Write-Host "No previous state found. Using empty baseline."
}

$newGroups = @()
$deletedGroups = @()
$groupsWithChanges = @{}
$allGroupsTable = @{}

$allGroups = $currentMembers.Keys + $oldmembers.Keys | Sort-Object -Unique

foreach ($group in $allGroups) {
    $current = $currentMembers[$group]
    $old = $oldmembers[$group]

    if ($null -eq $old) {
        $newGroups += $group
        $groupsWithChanges[$group] = @()
        if ($null -ne $current) {
            foreach ($user in $current) {
                $groupsWithChanges[$group] += @{ Type = 'Added'; User = $user }
            }
        }
    }
    elseif ($null -eq $current) {
        $deletedGroups += $group
        $groupsWithChanges[$group] = @()
        if ($null -ne $old) {
            foreach ($user in $old) {
                $groupsWithChanges[$group] += @{ Type = 'Removed'; User = $user }
            }
        }
    }
    else {
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

# Build HTML report
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
foreach ($group in $groupsWithChanges.Keys) {
    $html += "<h3>$group</h3><table><tr><th>Change Type</th><th>Member</th></tr>"
    foreach ($entry in $groupsWithChanges[$group]) {
        $html += "<tr><td class='$($entry.Type.ToLower())'>$($entry.Type)</td><td class='$($entry.Type.ToLower())'>$($entry.User)</td></tr>"
    }
    $html += "</table>"
}

$html += "<h2>All Groups</h2>"
foreach ($group in $allGroupsTable.Keys | Sort-Object) {
    $html += "<h3>$group</h3><table><tr><th>Change Type</th><th>Member</th></tr>"
    foreach ($entry in $allGroupsTable[$group]) {
        $html += "<tr><td class='$($entry.Type.ToLower())'>$($entry.Type)</td><td>$($entry.User)</td></tr>"
    }
    $html += "</table>"
}

$html += "</body></html>"

# Save report
Write-Host "Saving report to $report"
$html | Out-File -Encoding utf8 $report

# Save current state
Write-Host "Saving current DL state to $previous"
$currentMembers | ConvertTo-Json -Depth 5 | Out-File $previous

Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Done."
