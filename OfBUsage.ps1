<#
-----------------------------------------------------------------
OfBUsage.ps1
Get a statistic of the current usage of 
OneDrive for Business in your configured tenant
and write the result to an CSV file.
This script reads all SPO user profiles. 
OfBUsage made by 
Christoph Wilfing, Martina Grom, Toni Pohl (atwork.at)
version 1.0 (2016-01-14), based on the idea of
https://technet.microsoft.com/en-us/library/dn911464.aspx
but optimized and combined with querying the OfB services
for each user who has a provisioned OfB.
Prerequesits: Install SharePoint Online Management Shell
from https://www.microsoft.com/en-us/download/details.aspx?id=35588
-----------------------------------------------------------------
#>
#----------------------------------------------------------------
# Configure your values here:
# Specifies the URL for your organization's SPO admin service
$AdminURI = "https://<yourtenant>-admin.sharepoint.com"
# Specifies the URL for the personal SPO site
$MySiteUrl = "https://<yourtenant>-my.sharepoint.com"
# Specifies the User account for an Office 365 global admin in your organization
$AdminAccount = '<youradministrator>@<yourtenant>.onmicrosoft.com'
$AdminPass = '<yourpassword>'
# Specifies the location where the list of MySites should be saved
$ResultFile = '.\OfBUsage.csv'
#----------------------------------------------------------------
# Here we start
Write-Host "Starting and Authenticating..."

# with user interaction:
# $cred = Get-Credential
# without user interaction:
$encryptedPassword = ConvertTo-SecureString $AdminPass -asplaintext -force
$encryptedPasswordString = ConvertFrom-SecureString -secureString $encryptedPassword
$cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $AdminAccount, $encryptedPassword
$AdminPass = ""

# Begin the process
# optimization: we do not need to load the libraries as in the orginial script
# (... [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client") etc.)
Connect-SPOService -url $AdminURI -Credential $cred

# Take the AdminAccount and the password, and create a credential
$ProfileServiceCred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($AdminAccount, $encryptedPassword)

# Add the path of the User Profile Service to the SPO admin URL, then create a new webservice proxy to access it
$proxyaddr = "$AdminURI/_vti_bin/UserProfileService.asmx?wsdl"
$UserProfileService= New-WebServiceProxy -Uri $proxyaddr -UseDefaultCredential False
$UserProfileService.Credentials = $ProfileServiceCred

# Set variables for authentication cookies
$strAuthCookie = $ProfileServiceCred.GetAuthenticationCookie($AdminURI)
$uri = New-Object System.Uri($AdminURI)
$container = New-Object System.Net.CookieContainer
$container.SetCookies($uri, $strAuthCookie)
$UserProfileService.CookieContainer = $container

# Sets the first User profile, at index -1
$UserProfileResult = $UserProfileService.GetUserProfileByIndex(-1)
$NumProfiles = $UserProfileService.GetUserProfileCount()
$UserProfileURLList = New-Object System.Collections.ArrayList
$i = 0
Write-Host "Getting User Profiles ($NumProfiles)"

# As long as the next User profile is NOT the one we started with (at -1)...
While ($UserProfileResult.NextValue -ne -1) 
{
    $i++
    if ($i % 50 -eq 0) {Write-Host '.'} else {Write-Host '.' -NoNewline}
    # Look for the Personal Space object in the User Profile and retrieve it
    # (PersonalSpace is the name of the path to a user's OneDrive for Business site. Users who have not yet created a 
    # OneDrive for Business site might not have this property set.)
    $Prop = $UserProfileResult.UserProfile | Where-Object { $_.Name -eq 'PersonalSpace' } 
    $Url= $Prop.Values[0].Value

    # If "PersonalSpace" (which we've copied to $Url) exists, add it to our UserProfileURLList
    if ($Url) {
        $UserProfileURLList.Add($Url) | Out-Null
    }
    
    # And now we check the next profile the same way...
    $UserProfileResult = $UserProfileService.GetUserProfileByIndex($UserProfileResult.NextValue)
}
Write-Host '.'

Write-Host 'Getting Storage Data of users'
$ofbsitelist = New-Object System.Collections.ArrayList
$j = 0
$UserProfileURLList |  % { $_.tostring().Substring(0,$_.ToString().Length-1) } | % {
    $j ++
    if ($j % 50 -eq 0) {Write-Host '.'} else {Write-Host '.' -NoNewline}
    $FullPath = $MySiteUrl + $_
    $ofbsitelist += Get-SPOSite -Identity $FullPath -Detailed
}
Write-Host '.'

# Delete the $ResultFile if existing
if (Test-Path $ResultFile) {
    Remove-Item -Path $ResultFile -Force
}
# Output of our generated list
$ofbsitelist | Select-Object owner,@{n='StorageUsageinMB';e={$_.StorageUsageCurrent}},URL | Export-Csv -Path $ResultFile -NoClobber -NoTypeInformation -Encoding UTF8 -Force -Delimiter ';'

Write-Host "Done. $i User Profiles. $j SPO sites. Check $ResultFile"
# End of script.
