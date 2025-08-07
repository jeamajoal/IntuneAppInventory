# JSON-Based Storage System for IntuneInventory
# Replaces SQLite with user-friendly JSON file storage

# Module-scoped storage variables
$Script:StorageRoot = $null
$Script:StoragePaths = @{}
$Script:InventoryData = @{}

function Initialize-JsonStorage {
    <#
    .SYNOPSIS
    Initializes the JSON storage system.
    
    .DESCRIPTION
    Sets up the directory structure and storage paths for JSON-based inventory data.
    
    .PARAMETER StoragePath
    Root path for storage. Defaults to user's AppData folder.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$StoragePath = "$env:APPDATA\IntuneInventory"
    )
    
    try {
        $Script:StorageRoot = $StoragePath
        
        # Ensure storage directory exists
        if (-not (Test-Path $Script:StorageRoot)) {
            New-Item -Path $Script:StorageRoot -ItemType Directory -Force | Out-Null
            Write-IntuneInventoryLog -Message "Created storage directory: $Script:StorageRoot" -Level Info
        }
        
        # Define storage structure
        $Script:StoragePaths = @{
            Metadata = Join-Path $Script:StorageRoot "metadata.json"
            Applications = Join-Path $Script:StorageRoot "applications.json"
            Scripts = Join-Path $Script:StorageRoot "scripts.json"
            Remediations = Join-Path $Script:StorageRoot "remediations.json"
            Assignments = Join-Path $Script:StorageRoot "assignments.json"
            InventoryRuns = Join-Path $Script:StorageRoot "inventory-runs.json"
            SourceCodeHistory = Join-Path $Script:StorageRoot "source-code-history.json"
            Reports = Join-Path $Script:StorageRoot "reports"
            SourceCode = Join-Path $Script:StorageRoot "source-code"
        }
        
        # Create subdirectories
        $SubDirs = @('reports', 'source-code')
        foreach ($Dir in $SubDirs) {
            $DirPath = Join-Path $Script:StorageRoot $Dir
            if (-not (Test-Path $DirPath)) {
                New-Item -Path $DirPath -ItemType Directory -Force | Out-Null
            }
        }
        
        # Initialize metadata if it doesn't exist
        if (-not (Test-Path $Script:StoragePaths.Metadata)) {
            $Metadata = @{
                Version = "2.0"
                StorageType = "JSON"
                Created = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                LastAccessed = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                TenantId = $Script:ProductionTenantId
                Description = "IntuneInventory JSON Storage - User-friendly alternative to SQLite"
            }
            
            Save-JsonData -FilePath $Script:StoragePaths.Metadata -Data $Metadata
            Write-IntuneInventoryLog -Message "Initialized JSON storage metadata" -Level Info
        }
        
        # Load existing data into memory
        Import-InventoryData
        
        Write-IntuneInventoryLog -Message "JSON storage system initialized successfully" -Level Info
        return $true
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to initialize JSON storage: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Import-InventoryData {
    <#
    .SYNOPSIS
    Imports inventory data from JSON files into memory.
    
    .PARAMETER ItemType
    Optional. The specific type of data to import and return (Applications, Scripts, Remediations, Assignments, InventoryRuns).
    If not specified, imports all data types and returns the full inventory data structure.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet('Applications', 'Scripts', 'Remediations', 'Assignments', 'InventoryRuns', 'SourceCodeHistory')]
        [string]$ItemType
    )
    
    try {
        # Initialize if needed
        if (-not $Script:InventoryData) {
            $Script:InventoryData = @{
                Applications = @()
                Scripts = @()
                Remediations = @()
                Assignments = @()
                InventoryRuns = @()
                SourceCodeHistory = @()
            }
        }
        
        # If specific ItemType requested, load and return just that
        if ($ItemType) {
            $FilePath = $Script:StoragePaths[$ItemType]
            if (Test-Path $FilePath) {
                $Data = Import-JsonData -FilePath $FilePath
                if ($Data) {
                    $Script:InventoryData[$ItemType] = $Data
                    Write-Verbose "Imported $($Data.Count) $ItemType items"
                }
            }
            return $Script:InventoryData[$ItemType]
        }
        
        # Load all data types
        foreach ($DataType in $Script:InventoryData.Keys) {
            $FilePath = $Script:StoragePaths[$DataType]
            if (Test-Path $FilePath) {
                $Data = Import-JsonData -FilePath $FilePath
                if ($Data) {
                    $Script:InventoryData[$DataType] = $Data
                    Write-Verbose "Imported $($Data.Count) $DataType items"
                }
            }
        }
        
        # Update last accessed time
        $Metadata = Import-JsonData -FilePath $Script:StoragePaths.Metadata
        if ($Metadata) {
            $Metadata.LastAccessed = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Save-JsonData -FilePath $Script:StoragePaths.Metadata -Data $Metadata
        }
        
        Write-IntuneInventoryLog -Message "Inventory data imported from JSON storage" -Level Verbose
        return $Script:InventoryData
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to import inventory data: $($_.Exception.Message)" -Level Warning
        if ($ItemType) {
            return @()
        } else {
            return $Script:InventoryData
        }
    }
}

function Save-JsonData {
    <#
    .SYNOPSIS
    Safely saves data to a JSON file with backup.
    
    .PARAMETER FilePath
    Path to the JSON file.
    
    .PARAMETER Data
    Data to save.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        
        [Parameter(Mandatory = $true)]
        [object]$Data
    )
    
    try {
        # Create backup if file exists
        if (Test-Path $FilePath) {
            $BackupPath = "$FilePath.backup"
            Copy-Item -Path $FilePath -Destination $BackupPath -Force
        }
        
        # Save data with pretty formatting for readability
        $JsonData = $Data | ConvertTo-Json -Depth 20
        $JsonData | Out-File -FilePath $FilePath -Encoding UTF8 -Force
        
        Write-Verbose "Saved data to: $FilePath"
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to save JSON data to $FilePath : $($_.Exception.Message)" -Level Error
        
        # Restore backup if save failed
        $BackupPath = "$FilePath.backup"
        if (Test-Path $BackupPath) {
            Copy-Item -Path $BackupPath -Destination $FilePath -Force
            Write-IntuneInventoryLog -Message "Restored backup for $FilePath" -Level Warning
        }
        throw
    }
}

function Import-JsonData {
    <#
    .SYNOPSIS
    Safely imports data from a JSON file.
    
    .PARAMETER FilePath
    Path to the JSON file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    try {
        if (-not (Test-Path $FilePath)) {
            return $null
        }
        
        $JsonContent = Get-Content -Path $FilePath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($JsonContent)) {
            return $null
        }
        
        return ($JsonContent | ConvertFrom-Json)
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to import JSON data from $FilePath : $($_.Exception.Message)" -Level Error
        return $null
    }
}

function Add-InventoryItem {
    <#
    .SYNOPSIS
    Adds an item to the inventory storage.
    
    .PARAMETER ItemType
    Type of item (Applications, Scripts, Remediations, etc.).
    
    .PARAMETER Item
    The item data to add.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Applications', 'Scripts', 'Remediations', 'Assignments', 'InventoryRuns', 'SourceCodeHistory')]
        [string]$ItemType,
        
        [Parameter(Mandatory = $true)]
        [object]$Item
    )
    
    try {
        # Add timestamp if not present
        if ($Item -is [hashtable]) {
            # Handle hashtable objects
            if (-not $Item.ContainsKey('LastUpdated')) {
                $Item['LastUpdated'] = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            }
        } else {
            # Handle PSObject/custom objects
            if (-not $Item.PSObject.Properties['LastUpdated']) {
                $Item | Add-Member -MemberType NoteProperty -Name 'LastUpdated' -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') -Force
            }
        }
        
        # Remove existing item with same ID (update scenario)
        $ItemId = if ($Item -is [hashtable]) { $Item['Id'] } else { $Item.Id }
        if ($ItemId) {
            $Script:InventoryData[$ItemType] = @($Script:InventoryData[$ItemType] | Where-Object { 
                $ExistingId = if ($_ -is [hashtable]) { $_['Id'] } else { $_.Id }
                $ExistingId -ne $ItemId 
            })
        }
        
        # Add new item
        $Script:InventoryData[$ItemType] += $Item
        
        # Save to file
        Save-JsonData -FilePath $Script:StoragePaths[$ItemType] -Data $Script:InventoryData[$ItemType]
        
        $DisplayName = if ($Item -is [hashtable]) { $Item['DisplayName'] } else { $Item.DisplayName }
        Write-Verbose "Added $ItemType item: $($DisplayName -or $ItemId)"
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to add $ItemType item: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Get-InventoryItems {
    <#
    .SYNOPSIS
    Retrieves items from inventory storage.
    
    .PARAMETER ItemType
    Type of items to retrieve.
    
    .PARAMETER Filter
    Optional filter scriptblock.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Applications', 'Scripts', 'Remediations', 'Assignments', 'InventoryRuns', 'SourceCodeHistory')]
        [string]$ItemType,
        
        [Parameter()]
        [scriptblock]$Filter
    )
    
    try {
        $Items = $Script:InventoryData[$ItemType]
        
        if ($Filter) {
            $Items = $Items | Where-Object $Filter
        }
        
        return $Items
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to retrieve $ItemType items: $($_.Exception.Message)" -Level Error
        return @()
    }
}

function Remove-InventoryItem {
    <#
    .SYNOPSIS
    Removes an item from inventory storage.
    
    .PARAMETER ItemType
    Type of item to remove.
    
    .PARAMETER ItemId
    ID of the item to remove.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Applications', 'Scripts', 'Remediations', 'Assignments', 'InventoryRuns', 'SourceCodeHistory')]
        [string]$ItemType,
        
        [Parameter(Mandatory = $true)]
        [string]$ItemId
    )
    
    try {
        $Script:InventoryData[$ItemType] = @($Script:InventoryData[$ItemType] | Where-Object { $_.Id -ne $ItemId })
        Save-JsonData -FilePath $Script:StoragePaths[$ItemType] -Data $Script:InventoryData[$ItemType]
        
        Write-IntuneInventoryLog -Message "Removed $ItemType item: $ItemId" -Level Info
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to remove $ItemType item: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Test-StorageConnection {
    <#
    .SYNOPSIS
    Tests if the JSON storage system is properly initialized.
    #>
    [CmdletBinding()]
    param()
    
    return ($null -ne $Script:StorageRoot -and (Test-Path $Script:StorageRoot))
}

function Get-StorageStatistics {
    <#
    .SYNOPSIS
    Gets statistics about the current storage.
    #>
    [CmdletBinding()]
    param()
    
    try {
        $Stats = @{
            StorageRoot = $Script:StorageRoot
            Applications = $Script:InventoryData.Applications.Count
            Scripts = $Script:InventoryData.Scripts.Count
            Remediations = $Script:InventoryData.Remediations.Count
            Assignments = $Script:InventoryData.Assignments.Count
            InventoryRuns = $Script:InventoryData.InventoryRuns.Count
            SourceCodeHistory = $Script:InventoryData.SourceCodeHistory.Count
        }
        
        # Add file sizes
        foreach ($Type in @('Applications', 'Scripts', 'Remediations', 'Assignments')) {
            $FilePath = $Script:StoragePaths[$Type]
            if (Test-Path $FilePath) {
                $FileSize = (Get-Item $FilePath).Length
                $Stats["${Type}FileSize"] = "{0:N2} KB" -f ($FileSize / 1KB)
            }
        }
        
        return [PSCustomObject]$Stats
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to get storage statistics: $($_.Exception.Message)" -Level Error
        return $null
    }
}

function Backup-JsonStorage {
    <#
    .SYNOPSIS
    Creates a backup of the entire JSON storage.
    
    .PARAMETER BackupPath
    Path for the backup file.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$BackupPath = "$Script:StorageRoot\backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').zip"
    )
    
    try {
        if (-not (Test-Path $Script:StorageRoot)) {
            throw "Storage root not found: $Script:StorageRoot"
        }
        
        # Create backup using PowerShell compression
        Compress-Archive -Path "$Script:StorageRoot\*" -DestinationPath $BackupPath -Force
        
        Write-IntuneInventoryLog -Message "Storage backup created: $BackupPath" -Level Info
        return $BackupPath
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to create storage backup: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Clear-StorageCache {
    <#
    .SYNOPSIS
    Clears the in-memory storage cache and reloads from files.
    #>
    [CmdletBinding()]
    param()
    
    try {
        Import-InventoryData
        Write-IntuneInventoryLog -Message "Storage cache cleared and reloaded" -Level Info
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to clear storage cache: $($_.Exception.Message)" -Level Error
        throw
    }
}
