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

# Debug: Show what we fetched
Write-Host "Current members fetched. Total groups: $($currentMembers.Keys.Count)"
Write-Host "First few groups: $($currentMembers.Keys | Select-Object -First 3)"
Write-Host "Sample group members:"
$sampleGroup = $currentMembers.Keys | Select-Object -First 1
if ($sampleGroup) {
    Write-Host "  $sampleGroup : $($currentMembers[$sampleGroup] -join ', ')"
}

# Load previous state
$oldmembers = @{}
if (Test-Path $previous) {
    Write-Host "Loading previous report from $previous"
    $jsonContent = Get-Content $previous -Raw
    Write-Host "JSON content length: $($jsonContent.Length)"
    $oldmembers = $jsonContent | ConvertFrom-Json
    
    # Debug: Check what we loaded
    Write-Host "Previous state loaded. Type: $($oldmembers.GetType().Name)"
    Write-Host "Previous state keys count: $($oldmembers.Keys.Count)"
    Write-Host "First few keys: $($oldmembers.Keys | Select-Object -First 3)"
    
    # Convert PSCustomObject to hashtable if needed
    if ($oldmembers -is [System.Management.Automation.PSCustomObject]) {
        Write-Host "Converting PSCustomObject to hashtable..."
        $tempHash = @{}
        foreach ($prop in $oldmembers.PSObject.Properties) {
            $tempHash[$prop.Name] = $prop.Value
        }
        $oldmembers = $tempHash
        Write-Host "Converted to hashtable. Keys count: $($oldmembers.Keys.Count)"
    }
} else {
    Write-Host "No previous report found, creating new baseline."
}

# Categorize groups
$newGroups = @()
$deletedGroups = @()
$groupsWithChanges = @{}
$allGroupsTable = @{}

$allGroups = $currentMembers.Keys + $oldmembers.Keys | Sort-Object -Unique

foreach ($group in $allGroups) {
    $current = $currentMembers[$group]
    $old = $oldmembers[$group]
    
    # Debug: Show what we're comparing
    Write-Host "Comparing group: $group"
    Write-Host "  Current: $($current -join ', ')"
    Write-Host "  Old: $($old -join ', ')"
    Write-Host "  Current type: $($current.GetType().Name), Old type: $($old.GetType().Name)"
    Write-Host "  Current null: $($current -eq $null), Old null: $($old -eq $null)"

    if ($old -eq $null -and $current -ne $null) {
        # This is a truly new group (exists now but didn't exist before)
        Write-Host "  -> Marking as NEW GROUP"
        $newGroups += $group
        $groupsWithChanges[$group] = @()
        foreach ($user in $current) {
            $groupsWithChanges[$group] += @{ Type = 'Added'; User = $user }
        }
    }
    elseif ($current -eq $null -and $old -ne $null) {
        # This is a truly deleted group (existed before but doesn't exist now)
        $deletedGroups += $group
        $groupsWithChanges[$group] = @()
        foreach ($user in $old) {
            $groupsWithChanges[$group] += @{ Type = 'Removed'; User = $user }
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

# Generate HTML
$htmlParts = @()
$htmlParts += @"
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
$htmlParts += "<h2> New Distribution Lists</h2><ul>"
foreach ($g in $newGroups) { $htmlParts += "<li>$g</li>" }
$htmlParts += "</ul>"

# Section: Deleted DLs
$htmlParts += "<h2> Deleted Distribution Lists</h2><ul>"
foreach ($g in $deletedGroups) { $htmlParts += "<li>$g</li>" }
$htmlParts += "</ul>"

# Section: Groups with Changes
$htmlParts += "<h2>Groups With Changes</h2>"
foreach ($group in $groupsWithChanges.Keys) {
    $htmlParts += "<h3>$group</h3><table><tr><th>Change Type</th><th>Member</th></tr>"
    foreach ($entry in $groupsWithChanges[$group]) {
        $htmlParts += "<tr><td class='$($entry.Type.ToLower())'>$($entry.Type)</td><td class='$($entry.Type.ToLower())'>$($entry.User)</td></tr>"
    }
    $htmlParts += "</table>"
}

# Section: All Groups
$htmlParts += "<h2>All Groups</h2>"
foreach ($group in $allGroupsTable.Keys | Sort-Object) {
    $htmlParts += "<h3>$group</h3><table><tr><th>Change Type</th><th>Member</th></tr>"
    foreach ($entry in $allGroupsTable[$group]) {
        $htmlParts += "<tr><td class='$($entry.Type.ToLower())'>$($entry.Type)</td><td>$($entry.User)</td></tr>"
    }
    $htmlParts += "</table>"
}

$htmlParts += "</body></html>"

$html = $htmlParts -join "`n"

# Save report
Write-Host "Saving report to $report"
$html | Out-File -Encoding utf8 $report

# Save current state
Write-Host "Saving current DL state to $previous"
$currentMembers | ConvertTo-Json -Depth 5 | Out-File $previous

Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Done."
