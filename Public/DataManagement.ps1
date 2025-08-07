function Update-IntuneInventoryItem {
    <#
    .SYNOPSIS
    Updates an existing inventory item with new information.
    
    .DESCRIPTION
    Updates metadata or other information for an existing application, script, or remediation
    in the inventory database.
    
    .PARAMETER ItemId
    The ID of the inventory item to update.
    
    .PARAMETER ItemType
    The type of item (Application, Script, Remediation).
    
    .PARAMETER Properties
    Hashtable of properties to update.
    
    .EXAMPLE
    Update-IntuneInventoryItem -ItemId "12345678-1234-1234-1234-123456789012" -ItemType "Application" -Properties @{ Description = "Updated description" }
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ItemId,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Application', 'Script', 'Remediation')]
        [string]$ItemType,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Properties
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Updating $ItemType ID: $ItemId" -Level Info -Source "Update-IntuneInventoryItem"
        
        if (-not $Script:DatabaseConnection -or -not (Test-DatabaseConnection -DatabaseConnection $Script:DatabaseConnection)) {
            throw "Database connection is not available. Please run Connect-IntuneInventory first."
        }
    }
    
    process {
        try {
            $TableName = "${ItemType}s"
            
            # Verify the item exists
            $CheckCommand = $Script:DatabaseConnection.CreateCommand()
            $CheckCommand.CommandText = "SELECT COUNT(*) FROM $TableName WHERE Id = @ItemId;"
            $null = $CheckCommand.Parameters.AddWithValue("@ItemId", $ItemId)
            
            $ItemExists = $CheckCommand.ExecuteScalar()
            $CheckCommand.Dispose()
            
            if ($ItemExists -eq 0) {
                throw "$ItemType with ID '$ItemId' not found in inventory database."
            }
            
            if ($PSCmdlet.ShouldProcess("$ItemType $ItemId", "Update properties")) {
                # Build update query
                $SetClauses = @()
                $UpdateCommand = $Script:DatabaseConnection.CreateCommand()
                
                foreach ($Property in $Properties.Keys) {
                    $SetClauses += "$Property = @$Property"
                    $null = $UpdateCommand.Parameters.AddWithValue("@$Property", $Properties[$Property])
                }
                
                # Always update LastUpdated
                $SetClauses += "LastUpdated = @LastUpdated"
                $null = $UpdateCommand.Parameters.AddWithValue("@LastUpdated", (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
                $null = $UpdateCommand.Parameters.AddWithValue("@ItemId", $ItemId)
                
                $UpdateCommand.CommandText = "UPDATE $TableName SET $($SetClauses -join ', ') WHERE Id = @ItemId;"
                
                $RowsAffected = $UpdateCommand.ExecuteNonQuery()
                $UpdateCommand.Dispose()
                
                if ($RowsAffected -gt 0) {
                    Write-IntuneInventoryLog -Message "Successfully updated $ItemType ID: $ItemId" -Level Info
                    Write-Host "Item updated successfully!" -ForegroundColor Green
                }
                else {
                    Write-IntuneInventoryLog -Message "No changes made to $ItemType ID: $ItemId" -Level Warning
                }
            }
        }
        catch {
            Write-IntuneInventoryLog -Message "Failed to update inventory item: $($_.Exception.Message)" -Level Error
            throw
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Update process completed" -Level Info -Source "Update-IntuneInventoryItem"
    }
}

function Remove-IntuneInventoryItem {
    <#
    .SYNOPSIS
    Removes an inventory item from the database.
    
    .DESCRIPTION
    Removes an application, script, or remediation from the inventory database
    including all related assignments and source code history.
    
    .PARAMETER ItemId
    The ID of the inventory item to remove.
    
    .PARAMETER ItemType
    The type of item (Application, Script, Remediation).
    
    .PARAMETER Force
    Skip confirmation prompt.
    
    .EXAMPLE
    Remove-IntuneInventoryItem -ItemId "12345678-1234-1234-1234-123456789012" -ItemType "Application"
    
    .EXAMPLE
    Remove-IntuneInventoryItem -ItemId "12345678-1234-1234-1234-123456789012" -ItemType "Script" -Force
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ItemId,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Application', 'Script', 'Remediation')]
        [string]$ItemType,
        
        [Parameter()]
        [switch]$Force
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Removing $ItemType ID: $ItemId" -Level Info -Source "Remove-IntuneInventoryItem"
        
        if (-not $Script:DatabaseConnection -or -not (Test-DatabaseConnection -DatabaseConnection $Script:DatabaseConnection)) {
            throw "Database connection is not available. Please run Connect-IntuneInventory first."
        }
    }
    
    process {
        try {
            $TableName = "${ItemType}s"
            
            # Get item information for confirmation
            $InfoCommand = $Script:DatabaseConnection.CreateCommand()
            $InfoCommand.CommandText = "SELECT DisplayName FROM $TableName WHERE Id = @ItemId;"
            $null = $InfoCommand.Parameters.AddWithValue("@ItemId", $ItemId)
            
            $DisplayName = $InfoCommand.ExecuteScalar()
            $InfoCommand.Dispose()
            
            if (-not $DisplayName) {
                throw "$ItemType with ID '$ItemId' not found in inventory database."
            }
            
            if ($Force -or $PSCmdlet.ShouldProcess("$ItemType '$DisplayName' ($ItemId)", "Remove from inventory")) {
                # Start transaction
                $Transaction = $Script:DatabaseConnection.BeginTransaction()
                
                try {
                    # Remove source code history
                    $HistoryCommand = $Script:DatabaseConnection.CreateCommand()
                    $HistoryCommand.Transaction = $Transaction
                    $HistoryCommand.CommandText = "DELETE FROM SourceCodeHistory WHERE ItemId = @ItemId AND ItemType = @ItemType;"
                    $null = $HistoryCommand.Parameters.AddWithValue("@ItemId", $ItemId)
                    $null = $HistoryCommand.Parameters.AddWithValue("@ItemType", $ItemType)
                    $null = $HistoryCommand.ExecuteNonQuery()
                    $HistoryCommand.Dispose()
                    
                    # Remove assignments
                    $AssignmentCommand = $Script:DatabaseConnection.CreateCommand()
                    $AssignmentCommand.Transaction = $Transaction
                    $AssignmentCommand.CommandText = "DELETE FROM Assignments WHERE ItemId = @ItemId AND ItemType = @ItemType;"
                    $null = $AssignmentCommand.Parameters.AddWithValue("@ItemId", $ItemId)
                    $null = $AssignmentCommand.Parameters.AddWithValue("@ItemType", $ItemType)
                    $null = $AssignmentCommand.ExecuteNonQuery()
                    $AssignmentCommand.Dispose()
                    
                    # Remove main item
                    $DeleteCommand = $Script:DatabaseConnection.CreateCommand()
                    $DeleteCommand.Transaction = $Transaction
                    $DeleteCommand.CommandText = "DELETE FROM $TableName WHERE Id = @ItemId;"
                    $null = $DeleteCommand.Parameters.AddWithValue("@ItemId", $ItemId)
                    $RowsAffected = $DeleteCommand.ExecuteNonQuery()
                    $DeleteCommand.Dispose()
                    
                    if ($RowsAffected -gt 0) {
                        $Transaction.Commit()
                        Write-IntuneInventoryLog -Message "Successfully removed $ItemType '$DisplayName' ($ItemId)" -Level Info
                        Write-Host "Item removed successfully!" -ForegroundColor Green
                        Write-Host "${ItemType}`: $DisplayName" -ForegroundColor Cyan
                    }
                    else {
                        $Transaction.Rollback()
                        Write-IntuneInventoryLog -Message "No item was removed" -Level Warning
                    }
                }
                catch {
                    $Transaction.Rollback()
                    throw
                }
                finally {
                    $Transaction.Dispose()
                }
            }
        }
        catch {
            Write-IntuneInventoryLog -Message "Failed to remove inventory item: $($_.Exception.Message)" -Level Error
            throw
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Remove process completed" -Level Info -Source "Remove-IntuneInventoryItem"
    }
}

function Get-IntuneInventoryAssignments {
    <#
    .SYNOPSIS
    Retrieves assignment information for inventory items.
    
    .DESCRIPTION
    Gets detailed assignment information for applications, scripts, or remediations
    including target groups and deployment settings.
    
    .PARAMETER ItemId
    The ID of the specific item to get assignments for.
    
    .PARAMETER ItemType
    The type of item (Application, Script, Remediation).
    
    .PARAMETER TargetGroupName
    Filter assignments by target group name.
    
    .PARAMETER IncludeUnassigned
    Include items that have no assignments.
    
    .EXAMPLE
    Get-IntuneInventoryAssignments -ItemType "Application"
    
    .EXAMPLE
    Get-IntuneInventoryAssignments -ItemId "12345678-1234-1234-1234-123456789012" -ItemType "Script"
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ItemId,
        
        [Parameter()]
        [ValidateSet('Application', 'Script', 'Remediation')]
        [string]$ItemType,
        
        [Parameter()]
        [string]$TargetGroupName,
        
        [Parameter()]
        [switch]$IncludeUnassigned
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Retrieving assignment information" -Level Info -Source "Get-IntuneInventoryAssignments"
        
        if (-not $Script:DatabaseConnection -or -not (Test-DatabaseConnection -DatabaseConnection $Script:DatabaseConnection)) {
            throw "Database connection is not available. Please run Connect-IntuneInventory first."
        }
    }
    
    process {
        try {
            $Query = @"
SELECT 
    a.Id as AssignmentId,
    a.ItemId,
    a.ItemType,
    a.TargetType,
    a.TargetGroupId,
    a.TargetGroupName,
    a.Intent,
    a.CreatedDateTime,
    a.LastModifiedDateTime,
    CASE 
        WHEN a.ItemType = 'Application' THEN app.DisplayName
        WHEN a.ItemType = 'Script' THEN s.DisplayName
        WHEN a.ItemType = 'Remediation' THEN r.DisplayName
    END as ItemDisplayName
FROM Assignments a
LEFT JOIN Applications app ON a.ItemId = app.Id AND a.ItemType = 'Application'
LEFT JOIN Scripts s ON a.ItemId = s.Id AND a.ItemType = 'Script'
LEFT JOIN Remediations r ON a.ItemId = r.Id AND a.ItemType = 'Remediation'
WHERE 1=1
"@
            
            $Parameters = @{}
            
            if ($ItemId) {
                $Query += " AND a.ItemId = @ItemId"
                $Parameters["@ItemId"] = $ItemId
            }
            
            if ($ItemType) {
                $Query += " AND a.ItemType = @ItemType"
                $Parameters["@ItemType"] = $ItemType
            }
            
            if ($TargetGroupName) {
                $Query += " AND a.TargetGroupName LIKE @TargetGroupName"
                $Parameters["@TargetGroupName"] = "%$TargetGroupName%"
            }
            
            $Query += " ORDER BY a.ItemType, ItemDisplayName, a.TargetGroupName;"
            
            $Command = $Script:DatabaseConnection.CreateCommand()
            $Command.CommandText = $Query
            
            foreach ($Param in $Parameters.Keys) {
                $null = $Command.Parameters.AddWithValue($Param, $Parameters[$Param])
            }
            
            $Reader = $Command.ExecuteReader()
            $Assignments = @()
            while ($Reader.Read()) {
                $Assignment = @{
                    AssignmentId         = $Reader["AssignmentId"]
                    ItemId               = $Reader["ItemId"]
                    ItemType             = $Reader["ItemType"]
                    ItemDisplayName      = $Reader["ItemDisplayName"]
                    TargetType           = $Reader["TargetType"]
                    TargetGroupId        = $Reader["TargetGroupId"]
                    TargetGroupName      = $Reader["TargetGroupName"]
                    Intent               = $Reader["Intent"]
                    CreatedDateTime      = $Reader["CreatedDateTime"]
                    LastModifiedDateTime = $Reader["LastModifiedDateTime"]
                }
                
                $Assignments += [PSCustomObject]$Assignment
            }
            $Reader.Close()
            $Command.Dispose()
            
            # If IncludeUnassigned is specified, add items without assignments
            if ($IncludeUnassigned) {
                $UnassignedQuery = @"
SELECT Id, DisplayName, 'Unassigned' as AssignmentStatus
FROM Applications
WHERE Id NOT IN (SELECT DISTINCT ItemId FROM Assignments WHERE ItemType = 'Application')
UNION ALL
SELECT Id, DisplayName, 'Unassigned' as AssignmentStatus
FROM Scripts
WHERE Id NOT IN (SELECT DISTINCT ItemId FROM Assignments WHERE ItemType = 'Script')
UNION ALL
SELECT Id, DisplayName, 'Unassigned' as AssignmentStatus
FROM Remediations
WHERE Id NOT IN (SELECT DISTINCT ItemId FROM Assignments WHERE ItemType = 'Remediation')
"@
                
                if ($ItemType) {
                    $TableName = "${ItemType}s"
                    $UnassignedQuery = @"
SELECT Id, DisplayName, 'Unassigned' as AssignmentStatus
FROM $TableName
WHERE Id NOT IN (SELECT DISTINCT ItemId FROM Assignments WHERE ItemType = @ItemType)
"@
                }
                
                $UnassignedCommand = $Script:DatabaseConnection.CreateCommand()
                $UnassignedCommand.CommandText = $UnassignedQuery
                
                if ($ItemType) {
                    $null = $UnassignedCommand.Parameters.AddWithValue("@ItemType", $ItemType)
                }
                
                $UnassignedReader = $UnassignedCommand.ExecuteReader()
                while ($UnassignedReader.Read()) {
                    $UnassignedItem = @{
                        AssignmentId         = $null
                        ItemId               = $UnassignedReader["Id"]
                        ItemType             = if ($ItemType) { $ItemType } else { "Unknown" }
                        ItemDisplayName      = $UnassignedReader["DisplayName"]
                        TargetType           = "None"
                        TargetGroupId        = $null
                        TargetGroupName      = "Unassigned"
                        Intent               = $null
                        CreatedDateTime      = $null
                        LastModifiedDateTime = $null
                    }
                    
                    $Assignments += [PSCustomObject]$UnassignedItem
                }
                $UnassignedReader.Close()
                $UnassignedCommand.Dispose()
            }
            
            return $Assignments
        }
        catch {
            Write-IntuneInventoryLog -Message "Failed to retrieve assignments: $($_.Exception.Message)" -Level Error
            throw
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Assignment retrieval completed" -Level Info -Source "Get-IntuneInventoryAssignments"
    }
}

function Backup-IntuneInventoryDatabase {
    <#
    .SYNOPSIS
    Creates a backup of the Intune inventory database.
    
    .DESCRIPTION
    Creates a backup copy of the SQLite database to preserve inventory data.
    
    .PARAMETER BackupPath
    The path where the backup file should be created.
    
    .PARAMETER IncludeTimestamp
    Include timestamp in the backup filename.
    
    .EXAMPLE
    Backup-IntuneInventoryDatabase -BackupPath "C:\Backups\IntuneInventory_Backup.db"
    
    .EXAMPLE
    Backup-IntuneInventoryDatabase -BackupPath "C:\Backups" -IncludeTimestamp
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BackupPath,
        
        [Parameter()]
        [switch]$IncludeTimestamp
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting database backup" -Level Info -Source "Backup-IntuneInventoryDatabase"
        
        if (-not $Script:DatabasePath -or -not (Test-Path -Path $Script:DatabasePath)) {
            throw "Source database not found. Please run Connect-IntuneInventory first."
        }
    }
    
    process {
        try {
            # Determine final backup path
            if (Test-Path -Path $BackupPath -PathType Container) {
                $Filename = "IntuneInventory_Backup"
                if ($IncludeTimestamp) {
                    $Filename += "_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                }
                $Filename += ".db"
                $FinalBackupPath = Join-Path -Path $BackupPath -ChildPath $Filename
            }
            else {
                $FinalBackupPath = $BackupPath
                if ($IncludeTimestamp) {
                    $Directory = Split-Path -Path $BackupPath -Parent
                    $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($BackupPath)
                    $Extension = [System.IO.Path]::GetExtension($BackupPath)
                    $FinalBackupPath = Join-Path -Path $Directory -ChildPath "$BaseName`_$(Get-Date -Format 'yyyyMMdd_HHmmss')$Extension"
                }
            }
            
            # Ensure backup directory exists
            $BackupDirectory = Split-Path -Path $FinalBackupPath -Parent
            if (-not (Test-Path -Path $BackupDirectory)) {
                New-Item -Path $BackupDirectory -ItemType Directory -Force | Out-Null
            }
            
            # Close database connection temporarily if open
            $WasConnected = $false
            if ($Script:DatabaseConnection -and $Script:DatabaseConnection.State -eq 'Open') {
                $WasConnected = $true
                $Script:DatabaseConnection.Close()
            }
            
            try {
                # Copy the database file
                Copy-Item -Path $Script:DatabasePath -Destination $FinalBackupPath -Force
                
                Write-IntuneInventoryLog -Message "Database backup created: $FinalBackupPath" -Level Info
                Write-Host "Database backup created successfully!" -ForegroundColor Green
                Write-Host "Backup location: $FinalBackupPath" -ForegroundColor Cyan
                
                return $FinalBackupPath
            }
            finally {
                # Reopen database connection if it was open
                if ($WasConnected) {
                    $Script:DatabaseConnection = Get-DatabaseConnection -DatabasePath $Script:DatabasePath
                }
            }
        }
        catch {
            Write-IntuneInventoryLog -Message "Database backup failed: $($_.Exception.Message)" -Level Error
            throw
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Database backup completed" -Level Info -Source "Backup-IntuneInventoryDatabase"
    }
}

function Restore-IntuneInventoryDatabase {
    <#
    .SYNOPSIS
    Restores the Intune inventory database from a backup.
    
    .DESCRIPTION
    Restores the SQLite database from a backup file, replacing the current database.
    
    .PARAMETER BackupPath
    The path to the backup file to restore from.
    
    .PARAMETER Force
    Skip confirmation prompt.
    
    .EXAMPLE
    Restore-IntuneInventoryDatabase -BackupPath "C:\Backups\IntuneInventory_Backup.db"
    
    .EXAMPLE
    Restore-IntuneInventoryDatabase -BackupPath "C:\Backups\IntuneInventory_Backup.db" -Force
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BackupPath,
        
        [Parameter()]
        [switch]$Force
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting database restore" -Level Info -Source "Restore-IntuneInventoryDatabase"
        
        if (-not (Test-Path -Path $BackupPath)) {
            throw "Backup file not found: $BackupPath"
        }
        
        if (-not $Script:DatabasePath) {
            throw "Target database path not set. Please run Connect-IntuneInventory first."
        }
    }
    
    process {
        try {
            if ($Force -or $PSCmdlet.ShouldProcess($Script:DatabasePath, "Restore database from backup")) {
                # Close database connection
                if ($Script:DatabaseConnection) {
                    Close-DatabaseConnection -DatabaseConnection $Script:DatabaseConnection
                    $Script:DatabaseConnection = $null
                }
                
                # Create backup of current database
                $CurrentBackupPath = "$($Script:DatabasePath).pre-restore.bak"
                if (Test-Path -Path $Script:DatabasePath) {
                    Copy-Item -Path $Script:DatabasePath -Destination $CurrentBackupPath -Force
                    Write-IntuneInventoryLog -Message "Current database backed up to: $CurrentBackupPath" -Level Info
                }
                
                try {
                    # Restore from backup
                    Copy-Item -Path $BackupPath -Destination $Script:DatabasePath -Force
                    
                    # Test the restored database
                    $TestConnection = Get-DatabaseConnection -DatabasePath $Script:DatabasePath
                    if (Test-DatabaseConnection -DatabaseConnection $TestConnection) {
                        Close-DatabaseConnection -DatabaseConnection $TestConnection
                        
                        # Reconnect
                        $Script:DatabaseConnection = Get-DatabaseConnection -DatabasePath $Script:DatabasePath
                        
                        Write-IntuneInventoryLog -Message "Database restored successfully from: $BackupPath" -Level Info
                        Write-Host "Database restored successfully!" -ForegroundColor Green
                        Write-Host "Backup source: $BackupPath" -ForegroundColor Cyan
                        Write-Host "Previous database backed up to: $CurrentBackupPath" -ForegroundColor Yellow
                    }
                    else {
                        throw "Restored database failed validation"
                    }
                }
                catch {
                    # Restore original database if restore failed
                    if (Test-Path -Path $CurrentBackupPath) {
                        Copy-Item -Path $CurrentBackupPath -Destination $Script:DatabasePath -Force
                        Write-IntuneInventoryLog -Message "Restored original database after failed restore" -Level Warning
                    }
                    throw
                }
            }
        }
        catch {
            Write-IntuneInventoryLog -Message "Database restore failed: $($_.Exception.Message)" -Level Error
            throw
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Database restore completed" -Level Info -Source "Restore-IntuneInventoryDatabase"
    }
}
