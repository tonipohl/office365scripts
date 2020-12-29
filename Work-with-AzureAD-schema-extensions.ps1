#--------------------------------------------------------
# Work-with-AzureAD-schema-extensions
# Demo how to use Microsoft.Graph PowerShell
# to create and manage Azure AD schema extensions
# atwork.at, Dec. 29, 2020
# Christoph Wilfing, Toni Pohl, Martina Grom
#--------------------------------------------------------

#-------------------------------------
# Install Microsoft.Graph PS
#-------------------------------------
# https://docs.microsoft.com/en-us/graph/powershell/installation
# Install-Module Microsoft.Graph -Scope CurrentUser
# For updating (latest version is '1.2.0')
# Update-Module Microsoft.Graph -Force
# Check the module:
# Get-InstalledModule Microsoft.Graph

Import-Module Microsoft.Graph

# Connect to the Graph as user and request the following permissions:
Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All","Application.ReadWrite.All", "Directory.AccessAsUser.All", "Directory.ReadWrite.All"

# See the context data
Get-MgContext

#-------------------------------------
# Read all existing schema extensions
#-------------------------------------
# https://docs.microsoft.com/en-us/graph/api/schemaextension-get?view=graph-rest-1.0
# GET /schemaExtensions/{id}
Get-MgSchemaExtension -All

#-------------------------------------
# Create a new schema extension
#-------------------------------------
# We create a new, empty ArrayList
$SchemaProperties = New-Object -TypeName System.Collections.ArrayList

# define our keys and the types
$prop1 = @{
    'name' = 'costcenter';
    'type' = 'String';
}
$prop2 = @{
    'name' = 'pin';
    'type' = 'Integer';
}

# and add them to the SchemaProperties
[void]$SchemaProperties.Add( $prop1)
[void]$SchemaProperties.Add( $prop2)

# Now we can create the new schema extension for the resource User
$SchemaExtension = New-MgSchemaExtension -TargetTypes  @('User') `
    -Properties $SchemaProperties `
    -Id 'myapp1' `
    -Description 'my organization additional user properties' `
    -Owner "992cc0fe-1c66-478e-8c67-2846dee6d149" 


#-------------------------------------
# Check the new schema extension:
#-------------------------------------
Get-MgSchemaExtension -SchemaExtensionId $SchemaExtension.Id | fl
# or search for the app name:
Get-MgSchemaExtension  -All | ? id -like '*myapp1' | fl

#-------------------------------------
# Add another property
#-------------------------------------
# Create a new schema property
$prop3 = @{
    'name' = 'isdirector';
    'type' = 'Boolean';
}
# and add our new property to our existing ArrayList
$SchemaProperties.Add($prop3)

# Update the schem extension with a) the full schema properties list and b) the owner!
Update-MgSchemaExtension -SchemaExtensionId $SchemaExtension.Id `
    -Properties $SchemaProperties `
    -Owner "992cc0fe-1c66-478e-8c67-2846dee6d149" 

#-------------------------------------
# Set the status to Available
# (if needed for production)
#-------------------------------------
# Update-MgSchemaExtension -SchemaExtensionId $SchemaExtension.Id `
#     -Status 'Available' `
#     -Owner "992cc0fe-1c66-478e-8c67-2846dee6d149" 

# Check again
Get-MgSchemaExtension -SchemaExtensionId $SchemaExtension.Id | fl

#-------------------------------------
# Remove a schema extension
#-------------------------------------
Remove-MgSchemaExtension -SchemaExtensionId $SchemaExtension.Id

#-------------------------------------
# Use aka.ms/ge for testing
#-------------------------------------

#-------------------------------------
# Get the data with the appid
#-------------------------------------
# https://graph.microsoft.com/v1.0/users/?$select=<extensionid>,userprincipalname
# https://graph.microsoft.com/v1.0/users/?$select=ext2irw6qzw_myapp1,userprincipalname

#-------------------------------------
# Set new data with the appid
#-------------------------------------
# https://graph.microsoft.com/v1.0/users/<userid>
# PATCH
# {
#     "ext2irw6qzw_myapp1": {
#       "costcenter": "K100",
#       "pin": 1220,
#       "isdirector": true }
# }

# End.
