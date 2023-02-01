$SPOAdminUrl = 'https://myunt-admin.sharepoint.com'
Import-Module -Name ActiveDirectory,ImportExcel,PnP.PowerShell
$Credential = New-Object -TypeName PSCredential -ArgumentList $UserName,(ConvertTo-SecureString -String $Password -AsPlainText -Force)
Try {
    Connect-PnPOnline -Url $SPOAdminUrl -Credential $Credential
} Catch {
    
    Write-host "Failed to connect to PNPONline"
    $ERROR[0]
    Exit 1

}


$ADProperties = @(
    'UserPrincipalName',
    'EmployeeID',
    'DisplayName',
    'eduPersonAffiliation',
    'Enabled',
    'MemberOf',
    'Surname',
    'Givenname'
)

$OutputProperties = @(
    'UserPrincipalName',
    'EmployeeID',
    'DisplayName',
    @{Label='eduPersonAffiliation';Expression={$_.'eduPersonAffiliation' -join ','}},
    'Enabled',
    'OneDriveURL',
    'OneDriveUsageMB',
    'MemberOf',
    'Surname',
    'Givenname'
)

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


Write-Host "Getting all OneDrive objects"
$AllOneDrives = array2hash (Get-PnPTenantSite -IncludeOneDriveSites -Detailed -Filter "Url -like '-my.sharepoint.com/personal/'" | Select-Object -Property Owner,Url,StorageUsageCurrent) 'Owner'
Write-Host "Associating OneDrive data"
$AlumniAccounts | ForEach-Object {
    if ($AllOneDrives.($_.UserPrincipalName)) {
        $_.OneDriveURL = $AllOneDrives.($_.UserPrincipalName).Url
        $_.OneDriveUsageMB = $AllOneDrives.($_.UserPrincipalName).StorageUsageCurrent
    }
    else {
        $_.OneDriveURL = $null
        $_.OneDriveUsageMB = $null
    }
}
$AlumniAccounts = $AlumniAccounts | Sort-Object onedriveusagemb
