param(
    [Parameter()]
    [String]$searchName,
    [String]$searchExpression,
    [switch]$deleteMail
)

# connect to scc
$mypwd = ConvertTo-SecureString -String 'the cert password' -Force -AsPlainText
Connect-IPPSSession -CertificateFilePath "cert file path" -CertificatePassword $mypwd  -AppID "app_id" -Organization "domain_name, such as hesaitech.com" -AzureADAuthorizationEndpointUri https://login.chinacloudapi.cn/common -ConnectionUri https://ps.compliance.protection.partner.outlook.cn/powershell-liveid -ShowBanner:$false

function Delete-MailBySearchName{
    param ($searchName)
    Write-Host "start run delete mail function -> " $searchName

    try {
        New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType SoftDelete  -ErrorAction Stop
    }
    catch {
        Write-Host "Run delete [" $searchName "] get Exception: " $PSItem.ToString()
        exit 9001
    }
}


if ($deleteMail){
    Delete-MailBySearchName -searchName $searchName
    exit
}

Write-Host "start run New-ComplianceSearch -> " $searchName

$Search=New-ComplianceSearch -Name $searchName -ExchangeLocation All -ContentMatchQuery $searchExpression
write-Host $Search

Start-ComplianceSearch -Identity $Search.Identity

# wait for the search to finish
while ((Get-ComplianceSearch $Search.Name).Status -ne "Completed") {
    Write-Host " ." -NoNewline
    Start-Sleep -s 3
}
Write-Host "search done"
$targetComplianceSearch = Get-ComplianceSearch $Search.Name
$searchResult = $targetComplianceSearch.SuccessResults

# filter the search result
$filteredItems = $searchResult | Select-String -Pattern 'Location:(.*?), Item count: (\d+)' -AllMatches | ForEach-Object {
    $_.Matches | ForEach-Object {
        [PSCustomObject]@{
            Location = $_.Groups[1].Value.Trim()
            'Item count' = $_.Groups[2].Value.Trim()
        }
    }
} | Where-Object { $_.'Item count' -ge 1 }

# write result to json file
$jsonString = $filteredItems | ConvertTo-Json
$jsonString | Out-File -FilePath 'output_file.json'
