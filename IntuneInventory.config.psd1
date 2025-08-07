# IntuneInventory Configuration Example
# Copy this file and customize for your environment

@{
    # Database Configuration
    DatabasePath = "$env:APPDATA\IntuneInventory\inventory.db"
    
    # Connection Settings
    TenantId = "your-tenant-id.onmicrosoft.com"
    ClientId = "your-app-registration-client-id"  # Optional for interactive auth
    
    # Inventory Settings
    InventorySettings = @{
        IncludeAssignments = $true
        ForceFullInventory = $false
        InventorySchedule = "Daily"  # Daily, Weekly, Monthly
    }
    
    # Reporting Settings
    ReportSettings = @{
        OutputDirectory = "C:\IntuneReports"
        DefaultFormats = @("CSV", "JSON")
        IncludeSourceCode = $false
        AutoExport = $true
    }
    
    # Logging Configuration (customize based on your logging standards)
    LoggingSettings = @{
        LogLevel = "Info"  # Verbose, Info, Warning, Error
        LogPath = "$env:APPDATA\IntuneInventory\Logs"
        MaxLogSizeMB = 10
        MaxLogFiles = 5
    }
    
    # Required Graph Scopes
    RequiredScopes = @(
        'DeviceManagementApps.Read.All',
        'DeviceManagementConfiguration.Read.All',
        'DeviceManagementManagedDevices.Read.All',
        'Directory.Read.All'
    )
    
    # Email Notification Settings (optional)
    EmailSettings = @{
        Enabled = $false
        SmtpServer = "smtp.company.com"
        Port = 587
        UseSsl = $true
        From = "intune-inventory@company.com"
        To = @("admin@company.com")
        Subject = "Intune Inventory Report - {0}"  # {0} will be replaced with date
    }
    
    # Source Code Management
    SourceCodeSettings = @{
        RequireComments = $true
        VersionTracking = $true
        DefaultVersion = "1.0"
    }
}

<#
Usage Example:

# Load configuration
$Config = Import-PowerShellDataFile -Path "IntuneInventory.config.psd1"

# Initialize with config
Initialize-IntuneInventoryDatabase -DatabasePath $Config.DatabasePath

# Connect with config
Connect-IntuneInventory -TenantId $Config.TenantId

# Run inventory with config settings
Invoke-IntuneApplicationInventory -IncludeAssignments:$Config.InventorySettings.IncludeAssignments

# Export reports with config settings
Export-IntuneInventoryReport -OutputDirectory $Config.ReportSettings.OutputDirectory
#>
