param (
    [string]$appId,
    [string]$orgName,
    [string]$thumbprint,
    [string]$previous = "previousMembers.json",
    [string]$report = "DLchanges_report.html"
)

write-host "Connecting to Exchange Online"

Connect-ExchangeOnline -AppId $appId -Organization $orgName -CertificateThumbprint $thumbprint

write-host "Fetching Distribution Lists"

$distributionLists = Get-DistributionGroup | sort-object DisplayName

$currentMembers = @{}
foreach ($distributionList in $distributionLists) {
    try {
        $members = Get-DistributionGroupMember -Identity $distributionList.Identity | select-object -ExpandProperty PrimarySmtpAddress
    } Catch {
        write-warning "Unable to fetch members for $($distributionList.DisplayName): $_"
        $members = @()
    }
    $currentMembers[$distributionList.DisplayName] = $members
}

# Load Previous report
$oldmembers = @{}
if (Test-Path $previous) {
    write-host "Loading previous report from $previous"
    $oldmembers = Get-Content $previous | ConvertFrom-Json -Depth 5
} else {
    write-host "No previous report found, creating new one"
}

#Generate HTML report 

$html = @"
<html>
<head>
    <style>
        body { fonr-family : Arial;}
        ".added { color: green; }",
        ".removed { color: darkorange; }",
        ".unchanged { color: black; }",
        "table { border-collapse: collapse; width: 100%; }",
        "th, td { padding: 8px 12px; border: 1px solid #ccc; text-align: left; }",
    </style>
</head>
<body>
<h2>Distribution List Membership Changes</h2>
<table>
<tr><th>Group</th><th>Change</th><th>Member</th></tr>
"@

#collect All Group names and sort

$allgroups = $currentMembers.keys + $oldmembers.keys | sort-Object -Unique
$groupReportRows = @()

foreach ($group in $allgroups){
    $current = $currentMembers[$group]
    $old = $oldmembers[$group]

    #Check if new lists added
    if($null -eq $old){
        #New Dl
        $groupReportRows += "<tr class='newgroup'><td>$group</td><td colspan='2'>🆕 New Distribution List</td></tr>"
        foreach ($user in $current) {
            $groupReportRows += "<tr><td>$group</td><td class='added'>Added</td><td class='added'>$user</td></tr>"
        }
        continue
    }
    
    #check if old groups added
    if ($null -eq $current) {
        # Deleted DL
        $groupReportRows += "<tr class='deletedgroup'><td>$group</td><td colspan='2'>❌ Deleted Distribution List</td></tr>"
        foreach ($user in $old) {
            $groupReportRows += "<tr><td>$group</td><td class='removed'>Removed</td><td class='removed'>$user</td></tr>"
        }
        continue
    }
    
    #find added and removed members
    $added = Compare-Object -ReferenceObject $old -DifferenceObject $current -PassThru | Where-Object{ $_ -in $current }
    $removed = Compare-Object -ReferenceObject $old -DifferenceObject $current -PassThru | Where-Object{ $_ -in $old }

    #Add retrived added and removed in report
    foreach ($user in $added){
        $groupReportRows += "<tr><td>$group</td><td class='added'>Added</td><td class='added'>$user</td></tr>"
    }
    foreach ($user in $removed) {
        $groupReportRows += "<tr><td>$group</td><td class='removed'>Removed</td><td class='removed'>$user</td></tr>"
    }
}

#If no changes 
if ($groupReportRows.Count -eq 0) {
    $groupReportRows += "<tr><td colspan='3'>✅ No changes detected</td></tr>"
}


$html += ($groupReportRows -join "`n")
$html += "</table></body></html>"

# Save report
Write-Host "Saving HTML report to $reportFile"
$html | Out-File -Encoding utf8 $reportFile

# Save current state
Write-Host "Saving current DL state to $historyFile"
$currentState | ConvertTo-Json -Depth 5 | Out-File $historyFile

Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Done."