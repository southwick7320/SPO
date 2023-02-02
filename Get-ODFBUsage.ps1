# credits to:
# Joel Phillips 
# Mac Edwards
$SPOAdminUrl = '<TenantURL>-admin.sharepoint.com'
Import-Module -Name ActiveDirectory,ImportExcel,PnP.PowerShell

Try {
    Connect-PnPOnline -Url $SPOAdminUrl -Credential $Credential
} Catch {
    
    Write-host "Failed to connect to PNPONline"
    $ERROR[0]
    Exit 1

}


function array2hash ($array, [string]$keyName) {  
    $hash = @{}  
    foreach ($element in $array) {    
        $key = $element."$keyName"   
        if ($hash[$key] -eq $null) {       
           $hash[$key] = $element   
        } 
        elseif ($hash[$key] -is [Collections.ArrayList]) {       
           ($hash[$key]).Add($element)    
        } else {       
           $hash[$key] = [Collections.ArrayList]@(($hash[$key]), $element)    
        }  
     }  
     return $hash
}

$OutputProperties = @(
    'UserPrincipalName',
    'DisplayName',
    'BlockCredential',
    'Department'
    'OneDriveURL',
    'OneDriveUsageMB',
    'Surname',
    'Givenname',
    'Title',
    'WhenCreated'
    @{label = 'Licenses';Expression={$_.Licenses.accountskuid -join ','}},
    'Office',
    'ResourceQuota',
    'ResourceQuotaWarningLevel',
    'ResourceUsageCurrent',
    'StorageUsage',
    'StorageMaximumLevel',
    'StorageWarningLevel',
    'StorageQuotaWarningLevel',
    'StorageQuota'


)

Write-host "Getting MSOL User data"
$msolusers = get-msoluser -All | where-Object {($_.licenses).AccountSkuId -match "swmail:SPE_F1"}

Write-Host "Getting all OneDrive objects"
$AllOneDrives = array2hash (Get-PnPTenantSite -IncludeOneDriveSites -Detailed -Filter "Url -like '-my.sharepoint.com/personal/'" ) 'Owner'

Write-Host "Associating OneDrive data"

$msoluser_prop = $msolusers | Select-Object -Property $OutputProperties
$msoluser_prop| ForEach-Object {
    if ($AllOneDrives.($_.UserPrincipalName)) {
        $_.OneDriveURL = $AllOneDrives.($_.UserPrincipalName).Url
        $_.OneDriveUsageMB = $AllOneDrives.($_.UserPrincipalName).StorageUsageCurrent
        $_.ResourceQuota = $AllOneDrives.($_.UserPrincipalName).ResourceQuota
        $_.ResourceUsageCurrent = $AllOneDrives.($_.UserPrincipalName).ResourceUsageCurrent
        $_.StorageUsage = $AllOneDrives.($_.UserPrincipalName).StorageUsage
        $_.StorageMaximumLevel = $AllOneDrives.($_.UserPrincipalName).StorageMaximumLevel
        $_.StorageWarningLevel = $AllOneDrives.($_.UserPrincipalName).StorageWarningLevel
        $_.StorageQuota = $AllOneDrives.($_.UserPrincipalName).StorageQuota
        $_.StorageQuotaWarningLevel = $AllOneDrives.($_.UserPrincipalName).StorageQuotaWarningLevel


        #write-host "x $(($_.UserPrincipalName)) $($AllOneDrives.($_.UserPrincipalName).StorageMaximumLevel)" -ForegroundColor Green
    }
    else {
        $_.OneDriveURL = $null
        $_.OneDriveUsageMB = $null
    }
}
$msoluser_sort= $msoluser_prop | Sort-Object onedriveusagemb -Descending

$filename = read-host "Enter file name"

$msoluser_sort | export-csv -Path .\$filename -NoTypeInformation

Invoke-Item .\$filename
Disconnect-PnPOnline
Write-host "Goodbye"
