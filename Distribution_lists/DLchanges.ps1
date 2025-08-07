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
$distributionLists = Get-DistributionGroup | Sort-Object DisplayName

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
    Write-Host "Loading previous report from $previous"
    $oldmembers = Get-Content $previous | ConvertFrom-Json 
} else {
    Write-Host "No previous report found, creating new baseline."
}

# Generate HTML report
$html = @"
<html>
<head>
    <style>
        body { font-family: Arial; }
        .added { color: green; }
        .removed { color: darkorange; }
        .unchanged { color: black; }
        .newgroup, .deletedgroup { font-weight: bold; background-color: #f9f9f9; }
        table { border-collapse: collapse; width: 100%; }
        th, td { padding: 8px 12px; border: 1px solid #ccc; text-align: left; }
        th { background-color: #eee; }
    </style>
</head>
<body>
<h2>Distribution List Membership Changes</h2>
<table>
<tr><th>Group</th><th>Change</th><th>Member</th></tr>
"@

$allGroups = $currentMembers.Keys + $oldmembers.Keys | Sort-Object -Unique
$groupReportRows = @()

foreach ($group in $allGroups) {
    $current = $currentMembers[$group]
    $old = $oldmembers[$group]

    if ($null -eq $old) {
        $groupReportRows += "<tr class='newgroup'><td>$group</td><td colspan='2'>🆕 New Distribution List</td></tr>"
        foreach ($user in $current) {
            $groupReportRows += "<tr><td>$group</td><td class='added'>Added</td><td class='added'>$user</td></tr>"
        }
        continue
    }

    if ($null -eq $current) {
        $groupReportRows += "<tr class='deletedgroup'><td>$group</td><td colspan='2'>❌ Deleted Distribution List</td></tr>"
        foreach ($user in $old) {
            $groupReportRows += "<tr><td>$group</td><td class='removed'>Removed</td><td class='removed'>$user</td></tr>"
        }
        continue
    }

    $added = Compare-Object -ReferenceObject $old -DifferenceObject $current -PassThru | Where-Object { $_ -in $current }
    $removed = Compare-Object -ReferenceObject $old -DifferenceObject $current -PassThru | Where-Object { $_ -in $old }

    foreach ($user in $added) {
        $groupReportRows += "<tr><td>$group</td><td class='added'>Added</td><td class='added'>$user</td></tr>"
    }
    foreach ($user in $removed) {
        $groupReportRows += "<tr><td>$group</td><td class='removed'>Removed</td><td class='removed'>$user</td></tr>"
    }
}

if ($groupReportRows.Count -eq 0) {
    $groupReportRows += "<tr><td colspan='3'>✅ No changes detected</td></tr>"
}

$html += ($groupReportRows -join "`n")
$html += "</table></body></html>"

Write-Host "Saving report to $report"
$html | Out-File -Encoding utf8 $report

Write-Host "Saving current DL state to $previous"
$currentMembers | ConvertTo-Json -Depth 5 | Out-File $previous

Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Done."
