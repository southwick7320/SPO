# credits to:
# Joel Phillips 
# Mac Edwards


param (
    $UserName = $null,
    $Password = $null,
    $UploadUrl = 'https://myunt.sharepoint.com/sites/ECS',
    $UploadFolder = 'SecureReports',
    $OutputFile = 'AlumniAudit.xlsx'
)



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

#Get users to downgrade
$ADFilter = @(
    
    'eduPersonAffiliation -eq "alum" -and',
    'eduPersonAffiliation -ne "applicant" -and',
    'eduPersonAffiliation -ne "student" -and',
    'msExchExtensionCustomAttribute5 -eq "MSOL" -and '
    'memberof -like "CN=GBL_Alumni_Downgrade,OU=Student Licensing,OU=O365_Licensing,OU=Services,DC=ad,DC=unt,DC=edu"'
    
) -join ' '

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

Write-Host "Getting on-prem AD alumni user objects"

$AlumniAccounts = Get-ADUser -Properties $ADProperties -Server students.ad.unt.edu -Filter $ADFilter | Select-Object -Property $OutputProperties
Write-host "Alumni accounts to process $(($alumniaccounts | measure-object).count)"

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

foreach($account in $alumniaccounts){

    $lastname = $account.Surname -replace "'","\"
    $firstname = $account.Givenname -replace "'","\"
    $displayname = $account.DIsplayname -replace "'","\"
    
    Write-host "Adding $($account.UserPrincipalName) to DB"
    $fields = [ordered]@{

        userprincipalname      = "'" + $account.userprincipalname + "'"
        revokeaccessdate       = "'" + $(get-date ((get-date).adddays(45)) -format MM/dd/yyyy) + "'"
        firstnotificationsent  = "'" + $null + "'"
        secondnotificationsent = "'" + $null + "'"
        finalnotificationsent  = "'" + $null + "'"
        exclude                = "'" + $null + "'"
        licensedowngraded      = "'" + $Null + "'"
        licensedowngradedate   = "'" + $null + "'"
        SPOsizeMB              = "'" + $account.OneDriveUsageMB + "'"
        notes                  = "'" + $null +"'"
        Firstname              = "'" + $FirstName           + "'"
        Lastname               = "'" + $LastName            + "'"
        OneDriveURL            = "'" + $account.OneDriveURL          + "'"
        EmployeeID             = "'" + $account.EmployeeID                 + "'"
        edupersonaffilation    = "'" + $account.eduPersonAffiliation  + "'"
        DisplayName            = "'" + $displayname          + "'"

}

    $DBFields = $fields.Keys -join ","
    $DBValues = $fields.Values -join ","
    
    Invoke-Sqlcmd -Query "IF NOT EXISTS (Select * FROM alumndowngrade WHERE EMPLOYEEID = $($fields.EMPLOYEEID)) BEGIN INSERT INTO alumndowngrade ($DBFields) VALUES ($DBValues) END"  -ServerInstance "MOS-Legion" -Database "AlumLicenseDowngrade"
}

Disconnect-PnPOnline
Write-host "Goodbye"
