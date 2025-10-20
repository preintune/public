<#
.SYNOPSIS
    Configure Managed Identity Access to SharePoint for Azure Logic Apps

.DESCRIPTION
    This PowerShell script enables Managed Identity access to SharePoint Online sites for Azure Logic Apps.
    The script performs two main operations:
    
    1. Grants Microsoft Graph API permissions (Sites.Selected) to the Managed Identity, allowing it to access 
       SharePoint sites through the Graph API.
    
    2. Assigns specific permissions (read or write) to the Managed Identity for a designated SharePoint site,
       enabling the Azure Logic App to interact with SharePoint content using its system-assigned or 
       user-assigned managed identity.
    
    This approach eliminates the need for storing credentials or certificates, providing a more secure
    authentication method for Logic Apps to access SharePoint resources. The Managed Identity authentication
    is particularly useful for automated workflows that need to read from or write to SharePoint sites
    without requiring interactive authentication.

.NOTES
    Prerequisites:
    - Azure PowerShell modules: Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Microsoft.Graph.Sites
    - Global Administrator or Application Administrator permissions in Azure AD
    - The Logic App must have a system-assigned or user-assigned managed identity enabled
    
    Author: Pascal Reimann
    Version: 1.0
    Date: 20. October 2025

.EXAMPLE
    # Update the parameters section with your specific values and run the script
    .\Configure-ManagedIdentity-SharePoint.ps1
#>

#region Parameters - CHANGE THESE VALUES AS NEEDED

# Add the correct ObjectID for the Managed Identity
$ObjectId = "e8800382-610d-4761-9b15-873065e53227"
 
# Add the correct Graph scope to grant
$graphScope = "Sites.Selected"

# Add the correct ApplicationID and Displayname for the Managed Identity
$application = @{
 
id = "827fc69f-2814-44d7-96bc-492f2bf21c83"
 
displayName = "test-team-site-automation"
 
}

# Add the correct role to grant the Managed Identity (read or write)
$appRole = "write"
 
# Add the correct SharePoint Online tenant URL and site name
$spoTenant = "tenant.sharepoint.com"
$spoSite  = "TestTeamSite"

#endregion

Connect-MgGraph -Scope AppRoleAssignment.ReadWrite.All

$graph = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"

$graphAppRole = $graph.AppRoles | Where-Object { $_.Value -eq $graphScope }

$appRoleAssignment = @{
 
    "principalId" = $ObjectId
 
    "resourceId"  = $graph.Id
 
    "appRoleId"   = $graphAppRole.Id
 
}
 
New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ObjectID -BodyParameter $appRoleAssignment | Format-List
 
# No need to change anything below

$spoSiteId = $spoTenant + ":/sites/" + $spoSite + ":"
 
Import-Module Microsoft.Graph.Sites
 
Connect-MgGraph -Scope Sites.FullControl.All
 
New-MgSitePermission -SiteId $spoSiteId -Roles $appRole -GrantedToIdentities @{ Application = $application }