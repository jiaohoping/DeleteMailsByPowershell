param(
    [Parameter()]
    [String]$subject,
    [String]$searchName
    
)
Connect-IPPSSession -UserPrincipalName xxxx@xxxx.com

# $Search = New-ComplianceSearch -Name $searchName -ExchangeLocation All  -ContentMatchQuery "Subject:$subject"
$Search = New-ComplianceSearch -Name $searchName -ExchangeLocation All  -ContentMatchQuery "Subject:$subject"
Start-ComplianceSearch $Search.Name

while ((Get-ComplianceSearch $Search.Name).Status -ne "Completed") {
    Write-Host " ." -NoNewline
    Start-Sleep -s 3
    
}
Write-Host "Search Done！"


$test = Get-ComplianceSearch $Search.Name
$SearchResultsBySplit = $test.SuccessResults.Trim("{", "}")
$ResultsHashTable = @{}
foreach ($item in $SearchResultsBySplit -split "Location:") {
    $tmp = $item.split(',').Trim()
    if ($tmp) {
        $ResultsHashTable[$tmp[0]] = $tmp[1].split(':')[1]
    }
}


foreach ($exist in $ResultsHashTable.Keys) {
    if ([Int]$ResultsHashTable[$exist] -gt 0) {
        Write-Host $exist
    }
}

$isDelete = Read-Host "Are u delete  (Y/N): "

if ($isDelete -eq 'Y') {
    New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType SoftDelete -Confirm:$false
}



# Disconnect-ExchangeOnline