using namespace System.Net
param($Timer)


Import-Module AzureADPreview -UseWindowsPowershell
Import-Module PartnerCenter -UseWindowsPowershell


$ApplicationId = $ENV:ApplicationId
$ApplicationSecret = $ENV:ApplicationSecret
$secPas = $ApplicationSecret| ConvertTo-SecureString -AsPlainText -Force
$tenantID = $ENV:tenantID
$refreshToken = $ENV:refeshToken
$ExchangeRefreshToken = $ENV:ExchangeRefreshToken
$upn = $ENV:upn
$credential = New-Object System.Management.Automation.PSCredential($ApplicationId, $secPas)
###Connect to your Own Partner Center to get a list of customers/tenantIDs #########
$aadGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $tenantID
$graphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $tenantID

Connect-AzureAD -AadAccessToken $aadGraphToken.AccessToken -AccountId $upn -MsAccessToken $graphToken.AccessToken

$customers = Get-AzureADContract
 
Write-Host "Found $($customers.Count) customers." -ForegroundColor DarkGreen


foreach ($customer in $customers) {

    Write-Host "Checking Shared Mailboxes for $($Customer.DisplayName)" -ForegroundColor Green
    $token = New-PartnerAccessToken -ApplicationId 'a0c73c16-a7e3-4564-9a95-2bdf47383716'-RefreshToken $ExchangeRefreshToken -Scopes 'https://outlook.office365.com/.default' -Tenant $customer.CustomerContextId
    $tokenValue = ConvertTo-SecureString "Bearer $($token.AccessToken)" -AsPlainText -Force
    $credential1 = New-Object System.Management.Automation.PSCredential($upn, $tokenValue)
    $customerId = $customer.DefaultDomainName
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell-liveid?DelegatedOrg=$($customerId)&BasicAuthToOAuthConversion=true" -Credential $credential1 -Authentication Basic -AllowRedirection
    Import-PSSession $session -AllowClobber
    $sharedMailboxes = Get-Mailbox -ResultSize Unlimited -Filter {recipienttypedetails -eq "SharedMailbox"}
    Remove-PSSession $session

try{
$CustAadGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes ‘https://graph.windows.net/.default’ -ServicePrincipal -Tenant $customer.CustomerContextId
    $CustGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes ‘https://graph.microsoft.com/.default’ -ServicePrincipal -Tenant $customer.CustomerContextId
   Connect-AzureAD -AadAccessToken $CustAadGraphToken.AccessToken -AccountId $upn -MsAccessToken $CustGraphToken.AccessToken -TenantId $customer.CustomerContextId
   $licensedUsers = Get-AzureADUser -All $true | where  {$_.AssignedLicenses}
}catch{"There was an error"}

foreach ($mailbox in $sharedMailboxes) {
    Add-Type -Path "C:\home\data\ManagedDependencies\210414162202318.r\AzureADPreview\2.0.2.134\Microsoft.Open.AzureADBeta.Graph.PowerShell.dll"
    Add-Type -Path "C:\home\data\ManagedDependencies\210414162202318.r\AzureADPreview\2.0.2.134\Microsoft.Open.AzureAD16.Graph.Client.dll"
        if ($licensedUsers.ObjectId -contains $mailbox.ExternalDirectoryObjectID) {
            Write-Host "$($mailbox.displayname) is a licensed shared mailbox" -ForegroundColor Yellow
             $userUPN="$($mailbox.userPrincipalName)"
             $userList = Get-AzureADUserLicenseDetail -ObjectID $userUPN
             $Skus = $userList.SkuId
             Write-Host $Skus
if($Skus -is [array])
    {
        $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
 foreach ($lic in $($Skus -split ' ')) {
            $Licenses.RemoveLicenses = $lic
            Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $licenses
        }

    } else {
        $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        $Licenses.RemoveLicenses =  $Skus
        Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $licenses
    }
Disconnect-AzureAD
}
}
}
