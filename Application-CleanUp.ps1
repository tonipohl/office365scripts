<#
-----------------------------------------------------------------
Application-CleanUp.ps1
Script by atwork.at, Toni Pohl, 2019-08-20
This script shows all applications of an Azure Active Directory
filtered by app name and URL to remove "old" apps
automatically if required. 
The newest app is the one with the latest EndDate, all others
could be removed if not used. Adapt this script as required.

Install the Azure PowerShell module as described at
https://docs.microsoft.com/en-us/powershell/azure/install-az-ps?view=azps-2.5.0

Furthermore a Global Admin of your Microsoft 365 tenant is needed.
-----------------------------------------------------------------
#>

# Connect (if not already connected)
Connect-AzAccount

# What is the app and URL you are looking for?
$appname = 'Delegate 365'
$url = 'https://localhost:44300/'
Write-Output "List all Apps '$appname' with Home URL $url..."

# Create objects for the result
$result = @()

Class app
{
    [Int]$Order
    [Bool]$Deleted
    [String]$Id
    [String]$DisplayName 
    [String]$ApplicationId
    [String]$KeyId
    [String]$ReplyUrls
    [datetime]$StartDate
    [datetime]$EndDate
    [Bool]$MultiTenant
}

# Get all apps
# https://github.com/Azure/azure-powershell/blob/master/src/Resources/Resources/help/Get-AzADApplication.md
$apps = Get-AzADApplication -DisplayNameStartWith $appname -First 100  | ? HomePage -like "*$($url)*"
$apps | Format-Table

foreach ($a in $apps) {
    # Get one app
    # https://github.com/Azure/azure-powershell/blob/master/src/Resources/Resources/help/Get-AzADApplication.md
    $oneapp = Get-AzADApplication -ObjectId $a.ObjectId
    # Get the credentials & validity
    # https://github.com/Azure/azure-powershell/blob/master/src/Resources/Resources/help/Get-AzADAppCredential.md
    $onecred = Get-AzADAppCredential -ObjectId $a.ObjectId

    # Create a new appitem with the summarized data for simple output
    $appitem = New-Object app
    # $appitem = New-Object app -Property @{Id=$oneapp.ObjectId;DisplayName=$oneapp.DisplayName;ApplicationId=$oneapp.ApplicationId;...}

    $appitem.Id = $oneapp.ObjectId
    $appitem.DisplayName = $oneapp.DisplayName
    $appitem.ApplicationId = $oneapp.ApplicationId
    $appitem.ReplyUrls = $oneapp.ReplyUrls
    $appitem.MultiTenant = $oneapp.AvailableToOtherTenants
    $appitem.StartDate = $onecred.StartDate
    $appitem.EndDate = $onecred.EndDate
    $appitem.KeyId = $onecred.KeyId
    $appitem.Deleted = $false

    $result += $appitem
}

# Sort by relevance: The newest app on top
# $result = $result | Sort-Object -Property EndDate -Descending
$result = $result | Sort-Object { $_.EndDate -as [datetime] } -Descending 

Write-Output "Delete all old Apps '$appname' with Home URL $url..."
$i = 0
foreach ($a in $result) {
    # Get one app and just let the first one alive, remove the rest.
    if ($i -gt 0) {
        $a.Deleted = $true
        # If you want to remove the old app...
        # Remove-AzADApplication -ObjectId $a.Id -Force
    }
    $a.Order = $i
    $i++
}

$result | Format-Table
# End.
