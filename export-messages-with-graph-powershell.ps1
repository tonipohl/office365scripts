# export-messages-with-graph-powershell.ps1
# atwork.at, Toni Pohl, Christoph Wilfing

# One-time process: Install the Graph module
Install-Module Microsoft.Graph -Scope CurrentUser
# Or update the existing module to the latest version
# Update-Module Microsoft.Graph

# Check the cmdlets
# Get-InstalledModule Microsoft.Graph

Import-Module Microsoft.Graph.Mail

# Connect with Mail.Read permissions
Connect-MgGraph -Scopes "Mail.Read"

# Show the user context just as info
Get-MgContext

# get your user id - insert your own primary email address here
$user = Get-MgUser -Filter "UserPrincipalName eq '<your-email-address>'"
# Get a list of all mail folders
$folders = Get-MgUserMailFolder -UserId $user.Id -All
# Select the Inbox
$inbox = $folders | Where-Object { $_.DisplayName -eq "Inbox" }
# Get a list of all sub folders of the Inbox
$childs = Get-MgUserMailFolderChildFolder -UserId $user.Id -MailFolderId $inbox.Id -All
# Select the desired folder
$myfolder = $childs | Where-Object { $_.DisplayName -eq "<your-subfolder>" }

# Get all mails and export them (add an optional where filter if needed).
# We remove all HTML tags, repair line breaks and HTML spaces to get a readable text in the result file.
Get-MgUserMailFolderMessage -All `
    -UserId $user.Id `
    -MailFolderId $myfolder.Id | `
    Select-Object `
    @{N = 'Received'; E = { $_.ReceivedDateTime } }, `
    @{N = 'Sender'; E = { $_.Sender.foreach{ ($_.Emailaddress) }.address } }, `
    @{N = 'ToRecipient'; E = { $_.ToRecipients.foreach{ ($_.Emailaddress) }.address } }, `
    @{N = 'ccRecipient'; E = { $_.ccRecipients.foreach{ ($_.Emailaddress) }.address } }, `
    @{N = 'Subject'; E = { $_.Subject } }, `
    @{N = 'Importance'; E = { $_.Importance } }, `
    @{N = 'Body'; E = { ($_.Body.Content -replace '</p>',"`r`n" -replace "<[^>]+>",'' -replace "&nbsp;",' ').trim() } } | `
    Where-Object {( ($_.Subject -notlike "*newsletter*") -and ($_.Subject -notlike "*FYI*") ) } | `
    Export-Csv ".\mails.csv" -Delimiter "`t" -Encoding utf8

# End. Check the mails.csv file. 
# Best, open it with Microsoft Excel: Menu Data, From Text/CSV and follow the wizard.

# Disconnect when done
Disconnect-MgGraph
