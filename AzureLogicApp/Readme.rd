# Azure Logic Apps - SharePoint Integration

This repository contains tools and scripts for integrating Azure Logic Apps with SharePoint Online using Managed Identity authentication.

## Overview

Azure Logic Apps provide a powerful platform for creating automated workflows in the cloud. This project focuses on enabling secure access to SharePoint Online resources through Managed Identity authentication, eliminating the need for storing credentials or certificates in your Logic Apps.

## Features

- **Secure Authentication**: Uses Azure Managed Identity for authentication to SharePoint
- **No Credential Storage**: Eliminates the need to store usernames, passwords, or certificates
- **Automated Setup**: PowerShell scripts to configure permissions automatically
- **Flexible Permissions**: Support for both read and write access to SharePoint sites
- **Graph API Integration**: Leverages Microsoft Graph API for SharePoint operations

**Purpose**: Configures Managed Identity permissions for SharePoint access in Azure Logic Apps.

**What it does**:
1. Connects to Microsoft Graph with appropriate permissions
2. Grants `Sites.Selected` permission to the Managed Identity
3. Assigns specific read/write permissions to the designated SharePoint site
4. Enables the Logic App to access SharePoint content without stored credentials

**Parameters to Configure**:
- `$ObjectId`: The Object ID of your Logic App's Managed Identity
- `$graphScope`: Microsoft Graph permission scope (default: "Sites.Selected")
- `$application`: Logic App application ID and display name
- `$appRole`: Permission level - "read" or "write"
- `$spoTenant`: Your SharePoint Online tenant URL
- `$spoSite`: Name of the target SharePoint site

## Usage Examples

### Basic SharePoint File Operations in Logic Apps

Once configured, your Logic App can perform SharePoint operations using HTTP actions with Managed Identity authentication:

1. **Read Files**: Use HTTP GET requests to retrieve file contents
2. **Create Files**: Use HTTP POST requests to create new documents
3. **Update Files**: Use HTTP PATCH requests to modify existing files
4. **List Items**: Query SharePoint lists and libraries

### Authentication in Logic App HTTP Actions

Configure your HTTP actions in Logic Apps with:
- **Authentication Type**: Managed Identity
- **Managed Identity**: System-assigned (or User-assigned if applicable)
- **Audience**: `https://graph.microsoft.com/`

## Security Best Practices

- **Principle of Least Privilege**: Only grant the minimum required permissions
- **Regular Review**: Periodically review and audit Managed Identity permissions
- **Site-Specific Access**: Use `Sites.Selected` scope and grant access only to required sites
- **Monitor Usage**: Enable logging and monitoring for SharePoint access patterns

## Troubleshooting

### Common Issues

1. **Permission Denied Errors**:
   - Verify Managed Identity is enabled on the Logic App
   - Ensure correct Object ID is used in the configuration
   - Check that Graph permissions have been granted

2. **Site Access Issues**:
   - Confirm SharePoint site URL is correct
   - Verify site permissions have been properly assigned
   - Check if the site requires special access policies

3. **Authentication Failures**:
   - Ensure the Logic App is using the correct Managed Identity
   - Verify the audience URL in HTTP actions is set to `https://graph.microsoft.com/`

## License

This project is licensed under the MIT License - see the [LICENSE](../LICENSE) file for details.

## Author

**Pascal Reimann**
- Version: 1.0
- Date: October 2025