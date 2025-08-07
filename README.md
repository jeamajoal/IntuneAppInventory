# IntuneInventory PowerShell Module

A comprehensive PowerShell module for inventorying Microsoft Intune applications, scripts, and remediations with SQLite backend storage and advanced reporting capabilities.

## Overview

The IntuneInventory module provides a complete solution for:
- Inventorying Intune applications, PowerShell scripts, and proactive remediations
- Storing comprehensive metadata and source code in a local SQLite database
- Generating digestible reports for quick reference and analysis
- Managing source code information with post-inventory addition capabilities
- Maintaining historical data for auditing and tracking changes

## Features

### Core Inventory Capabilities
- **Application Inventory**: Complete metadata, installation details, and source code collection
- **Script Inventory**: PowerShell script content, metadata, and execution context
- **Remediation Inventory**: Detection and remediation script collection with metadata
- **Assignment Tracking**: Group assignments and deployment settings

### Data Management
- **SQLite Backend**: Local database storage for fast querying and offline access
- **Source Code Management**: Add source code after initial inventory with version tracking
- **Historical Tracking**: Maintain inventory run history and change logs
- **Data Integrity**: Transaction-based operations with rollback capabilities

### Reporting and Analysis
- **Multiple Report Types**: Summary, detailed, and executive reports
- **Export Formats**: CSV, JSON, HTML, and object output
- **Missing Source Code Analysis**: Identify items requiring source code addition
- **Executive Summaries**: High-level overviews for management reporting

## Prerequisites

### Required PowerShell Modules
```powershell
# Install required Microsoft Graph modules
Install-Module Microsoft.Graph.Authentication -Force
Install-Module Microsoft.Graph.DeviceManagement -Force
Install-Module Microsoft.Graph.DeviceManagement.Administration -Force
```

### Required Permissions
The module requires the following Microsoft Graph API permissions:
- `DeviceManagementApps.Read.All`
- `DeviceManagementConfiguration.Read.All`
- `DeviceManagementManagedDevices.Read.All`
- `Directory.Read.All`

### SQLite Support
The module includes SQLite support through System.Data.SQLite. If you encounter issues, you may need to install the SQLite provider manually.

## Installation

1. Clone or download the module to your PowerShell modules directory:
```powershell
# Copy the module to your PowerShell modules path
$ModulePath = "$env:USERPROFILE\Documents\PowerShell\Modules\IntuneInventory"
# Copy files to $ModulePath
```

2. Import the module:
```powershell
Import-Module IntuneInventory
```

## Quick Start

### 1. Initialize the Database
```powershell
# Initialize the SQLite database (creates default location)
Initialize-IntuneInventoryDatabase

# Or specify a custom location
Initialize-IntuneInventoryDatabase -DatabasePath "C:\IntuneInventory\inventory.db"
```

### 2. Connect to Intune
```powershell
# Interactive authentication
Connect-IntuneInventory

# With specific tenant
Connect-IntuneInventory -TenantId "contoso.onmicrosoft.com"

# Service principal authentication
$ClientSecret = ConvertTo-SecureString "your-secret" -AsPlainText -Force
Connect-IntuneInventory -TenantId "tenant-id" -ClientId "client-id" -ClientSecret $ClientSecret
```

### 3. Run Inventory
```powershell
# Inventory applications
Invoke-IntuneApplicationInventory

# Inventory scripts
Invoke-IntuneScriptInventory

# Inventory remediations
Invoke-IntuneRemediationInventory

# Include assignments (optional)
Invoke-IntuneApplicationInventory -IncludeAssignments
```

### 4. Generate Reports
```powershell
# Quick summary
Get-IntuneInventoryReport -ReportType Summary

# Detailed application report
Get-IntuneInventoryReport -ReportType Applications -Format Table

# Export comprehensive reports
Export-IntuneInventoryReport -OutputDirectory "C:\Reports"

# Find items missing source code
Get-IntuneInventoryReport -ReportType All -FilterMissingSourceCode
```

## Core Functions

### Connection Management
- `Initialize-IntuneInventoryDatabase` - Set up the SQLite database
- `Connect-IntuneInventory` - Authenticate to Microsoft Graph
- `Disconnect-IntuneInventory` - Clean up connections

### Inventory Operations
- `Invoke-IntuneApplicationInventory` - Inventory applications
- `Invoke-IntuneScriptInventory` - Inventory PowerShell scripts
- `Invoke-IntuneRemediationInventory` - Inventory proactive remediations

### Source Code Management
- `Add-IntuneInventorySourceCode` - Add source code to existing items
- `Get-IntuneInventorySourceCode` - Retrieve source code with history
- `Get-IntuneInventoryItem` - Get detailed item information

### Reporting and Analysis
- `Get-IntuneInventoryReport` - Generate various report types
- `Export-IntuneInventoryReport` - Export comprehensive reports
- `Get-IntuneInventoryAssignments` - Analyze assignment information

### Data Management
- `Update-IntuneInventoryItem` - Update existing inventory items
- `Remove-IntuneInventoryItem` - Remove items from inventory
- `Backup-IntuneInventoryDatabase` - Backup the database
- `Restore-IntuneInventoryDatabase` - Restore from backup

## Advanced Usage

### Adding Source Code After Inventory
```powershell
# Add source code to an application
$SourceCode = Get-Content "C:\Scripts\AppInstaller.ps1" -Raw
Add-IntuneInventorySourceCode -ItemId "app-guid" -ItemType "Application" -SourceCode $SourceCode -Comments "Installation script" -Version "1.0"

# Add source code to a script
$ScriptContent = Get-Content "C:\Scripts\ConfigScript.ps1" -Raw
Add-IntuneInventorySourceCode -ItemId "script-guid" -ItemType "Script" -SourceCode $ScriptContent -Comments "Updated configuration script"
```

### Custom Queries
```powershell
# The module stores data in SQLite, allowing direct queries
$DatabasePath = (Get-Module IntuneInventory).PrivateData.DatabasePath
# Use SQLite tools or PowerShell to query the database directly
```

### Automated Reporting
```powershell
# Schedule regular inventory and reporting
$ReportPath = "\\server\share\IntuneReports"
Connect-IntuneInventory
Invoke-IntuneApplicationInventory -Force
Invoke-IntuneScriptInventory -Force
Invoke-IntuneRemediationInventory -Force
Export-IntuneInventoryReport -OutputDirectory $ReportPath
Disconnect-IntuneInventory
```

## Database Schema

The module creates the following tables:
- **Applications** - Application metadata and source code
- **Scripts** - PowerShell script information and content
- **Remediations** - Proactive remediation details and scripts
- **Assignments** - Assignment information for all item types
- **InventoryRuns** - History of inventory operations
- **SourceCodeHistory** - Version history of source code changes

## Configuration

### Database Location
By default, the database is stored in:
```
$env:APPDATA\IntuneInventory\inventory.db
```

### Logging Configuration
The module includes placeholder logging functions that can be customized based on your organization's logging standards. Update the `Write-IntuneInventoryLog` function in `Private\Utilities.ps1` to integrate with your logging system.

### Authentication Configuration
The module supports multiple authentication methods:
- Interactive authentication (default)
- Service principal with client secret
- Certificate-based authentication

## Troubleshooting

### Common Issues

1. **SQLite Assembly Loading**
   ```powershell
   # If SQLite assembly fails to load, install the provider
   Install-Package System.Data.SQLite.Core
   ```

2. **Graph Permission Errors**
   ```powershell
   # Verify required permissions are granted in Azure AD
   Get-MgContext | Select-Object Scopes
   ```

3. **Database Connection Issues**
   ```powershell
   # Check database file permissions and path
   Test-Path $Script:DatabasePath
   ```

### Debug Mode
Enable verbose output for troubleshooting:
```powershell
$VerbosePreference = 'Continue'
Connect-IntuneInventory -Verbose
```

## Contributing

This module is designed to be extended and customized for specific organizational needs. Key areas for customization:
- Logging implementation in `Private\Utilities.ps1`
- Authentication methods in `Public\Connection.ps1`
- Custom reporting formats in `Private\ReportingHelpers.ps1`
- Additional data collection in inventory functions

## Version History

- **1.0.0** - Initial release with core inventory and reporting functionality

## License

This module is provided as-is for Sound Physicians internal use. Modify and extend as needed for your organization's requirements.

## Support

For issues and enhancements:
1. Check the verbose logs for detailed error information
2. Verify Microsoft Graph permissions and connectivity
3. Ensure SQLite database is accessible and not corrupted
4. Review the module's `.github\copilot-instructions.md` for development guidelines
