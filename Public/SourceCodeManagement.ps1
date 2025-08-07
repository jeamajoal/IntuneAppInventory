function Add-IntuneInventorySourceCode {
    <#
    .SYNOPSIS
    Adds source code to an existing inventory item.
    
    .DESCRIPTION
    Allows adding source code information to applications, scripts, or remediations
    that were inventoried without source code initially. Maintains a history of 
    source code additions and updates.
    
    .PARAMETER ItemId
    The ID of the inventory item to update.
    
    .PARAMETER ItemType
    The type of item (Application, Script, Remediation).
    
    .PARAMETER SourceCode
    The source code content to add.
    
    .PARAMETER Comments
    Optional comments about the source code addition.
    
    .PARAMETER Version
    Optional version information for the source code.
    
    .EXAMPLE
    Add-IntuneInventorySourceCode -ItemId "12345678-1234-1234-1234-123456789012" -ItemType "Application" -SourceCode $ScriptContent
    
    .EXAMPLE
    Add-IntuneInventorySourceCode -ItemId "12345678-1234-1234-1234-123456789012" -ItemType "Script" -SourceCode $PowerShellScript -Comments "Added missing deployment script" -Version "1.0"
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ItemId,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Application', 'Script', 'Remediation')]
        [string]$ItemType,
        
        [Parameter(Mandatory = $true)]
        [string]$SourceCode,
        
        [Parameter()]
        [string]$Comments = "",
        
        [Parameter()]
        [string]$Version = ""
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Adding source code for $ItemType ID: $ItemId" -Level Info -Source "Add-IntuneInventorySourceCode"
        
        # Verify database connection
        if (-not $Script:DatabaseConnection -or -not (Test-DatabaseConnection -DatabaseConnection $Script:DatabaseConnection)) {
            throw "Database connection is not available. Please run Connect-IntuneInventory first."
        }
    }
    
    process {
        try {
            # Verify the item exists
            $TableName = "${ItemType}s"
            $CheckCommand = $Script:DatabaseConnection.CreateCommand()
            $CheckCommand.CommandText = "SELECT COUNT(*) FROM $TableName WHERE Id = @ItemId;"
            $null = $CheckCommand.Parameters.AddWithValue("@ItemId", $ItemId)
            
            $ItemExists = $CheckCommand.ExecuteScalar()
            $CheckCommand.Dispose()
            
            if ($ItemExists -eq 0) {
                throw "$ItemType with ID '$ItemId' not found in inventory database."
            }
            
            if ($PSCmdlet.ShouldProcess("$ItemType $ItemId", "Add source code")) {
                $CurrentTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $UserContext = Get-CurrentUserContext
                
                # Start transaction
                $Transaction = $Script:DatabaseConnection.BeginTransaction()
                
                try {
                    # Add to source code history
                    $HistoryCommand = $Script:DatabaseConnection.CreateCommand()
                    $HistoryCommand.Transaction = $Transaction
                    $HistoryCommand.CommandText = @"
INSERT INTO SourceCodeHistory (ItemId, ItemType, SourceCode, AddedBy, AddedDate, Comments, Version)
VALUES (@ItemId, @ItemType, @SourceCode, @AddedBy, @AddedDate, @Comments, @Version);
"@
                    $null = $HistoryCommand.Parameters.AddWithValue("@ItemId", $ItemId)
                    $null = $HistoryCommand.Parameters.AddWithValue("@ItemType", $ItemType)
                    $null = $HistoryCommand.Parameters.AddWithValue("@SourceCode", $SourceCode)
                    $null = $HistoryCommand.Parameters.AddWithValue("@AddedBy", $UserContext.UserPrincipalName)
                    $null = $HistoryCommand.Parameters.AddWithValue("@AddedDate", $CurrentTime)
                    $null = $HistoryCommand.Parameters.AddWithValue("@Comments", $Comments)
                    $null = $HistoryCommand.Parameters.AddWithValue("@Version", $Version)
                    
                    $null = $HistoryCommand.ExecuteNonQuery()
                    $HistoryCommand.Dispose()
                    
                    # Update the main item record
                    $UpdateCommand = $Script:DatabaseConnection.CreateCommand()
                    $UpdateCommand.Transaction = $Transaction
                    $UpdateCommand.CommandText = @"
UPDATE $TableName 
SET SourceCode = @SourceCode, 
    HasSourceCode = 1, 
    SourceCodeAdded = @SourceCodeAdded,
    LastUpdated = @LastUpdated
WHERE Id = @ItemId;
"@
                    $null = $UpdateCommand.Parameters.AddWithValue("@SourceCode", $SourceCode)
                    $null = $UpdateCommand.Parameters.AddWithValue("@SourceCodeAdded", $CurrentTime)
                    $null = $UpdateCommand.Parameters.AddWithValue("@LastUpdated", $CurrentTime)
                    $null = $UpdateCommand.Parameters.AddWithValue("@ItemId", $ItemId)
                    
                    $null = $UpdateCommand.ExecuteNonQuery()
                    $UpdateCommand.Dispose()
                    
                    # Commit transaction
                    $Transaction.Commit()
                    
                    Write-IntuneInventoryLog -Message "Source code added successfully for $ItemType ID: $ItemId" -Level Info
                    Write-Host "Source code added successfully!" -ForegroundColor Green
                    Write-Host "$ItemType ID: $ItemId" -ForegroundColor Cyan
                    if ($Comments) {
                        Write-Host "Comments: $Comments" -ForegroundColor White
                    }
                    if ($Version) {
                        Write-Host "Version: $Version" -ForegroundColor White
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
            Write-IntuneInventoryLog -Message "Failed to add source code: $($_.Exception.Message)" -Level Error
            throw
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Source code addition process completed" -Level Info -Source "Add-IntuneInventorySourceCode"
    }
}

function Get-IntuneInventorySourceCode {
    <#
    .SYNOPSIS
    Retrieves source code for an inventory item.
    
    .DESCRIPTION
    Gets the current source code and optionally the history of source code changes
    for a specific inventory item.
    
    .PARAMETER ItemId
    The ID of the inventory item.
    
    .PARAMETER ItemType
    The type of item (Application, Script, Remediation).
    
    .PARAMETER IncludeHistory
    Include the complete history of source code changes.
    
    .PARAMETER HistoryOnly
    Return only the source code history, not the current version.
    
    .EXAMPLE
    Get-IntuneInventorySourceCode -ItemId "12345678-1234-1234-1234-123456789012" -ItemType "Script"
    
    .EXAMPLE
    Get-IntuneInventorySourceCode -ItemId "12345678-1234-1234-1234-123456789012" -ItemType "Application" -IncludeHistory
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ItemId,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Application', 'Script', 'Remediation')]
        [string]$ItemType,
        
        [Parameter()]
        [switch]$IncludeHistory,
        
        [Parameter()]
        [switch]$HistoryOnly
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Retrieving source code for $ItemType ID: $ItemId" -Level Info -Source "Get-IntuneInventorySourceCode"
        
        # Verify database connection
        if (-not $Script:DatabaseConnection -or -not (Test-DatabaseConnection -DatabaseConnection $Script:DatabaseConnection)) {
            throw "Database connection is not available. Please run Connect-IntuneInventory first."
        }
    }
    
    process {
        try {
            $Result = @{}
            $TableName = "${ItemType}s"
            
            if (-not $HistoryOnly) {
                # Get current source code
                $CurrentCommand = $Script:DatabaseConnection.CreateCommand()
                $CurrentCommand.CommandText = @"
SELECT DisplayName, HasSourceCode, SourceCode, SourceCodeAdded, LastUpdated
FROM $TableName 
WHERE Id = @ItemId;
"@
                $null = $CurrentCommand.Parameters.AddWithValue("@ItemId", $ItemId)
                
                $Reader = $CurrentCommand.ExecuteReader()
                if ($Reader.Read()) {
                    $Result.Current = @{
                        ItemId = $ItemId
                        ItemType = $ItemType
                        DisplayName = $Reader["DisplayName"]
                        HasSourceCode = [bool]$Reader["HasSourceCode"]
                        SourceCode = $Reader["SourceCode"]
                        SourceCodeAdded = $Reader["SourceCodeAdded"]
                        LastUpdated = $Reader["LastUpdated"]
                    }
                }
                else {
                    throw "$ItemType with ID '$ItemId' not found in inventory database."
                }
                $Reader.Close()
                $CurrentCommand.Dispose()
            }
            
            if ($IncludeHistory -or $HistoryOnly) {
                # Get source code history
                $HistoryCommand = $Script:DatabaseConnection.CreateCommand()
                $HistoryCommand.CommandText = @"
SELECT Id, SourceCode, AddedBy, AddedDate, Comments, Version
FROM SourceCodeHistory 
WHERE ItemId = @ItemId AND ItemType = @ItemType
ORDER BY AddedDate DESC;
"@
                $null = $HistoryCommand.Parameters.AddWithValue("@ItemId", $ItemId)
                $null = $HistoryCommand.Parameters.AddWithValue("@ItemType", $ItemType)
                
                $HistoryReader = $HistoryCommand.ExecuteReader()
                $History = @()
                while ($HistoryReader.Read()) {
                    $History += @{
                        HistoryId = $HistoryReader["Id"]
                        SourceCode = $HistoryReader["SourceCode"]
                        AddedBy = $HistoryReader["AddedBy"]
                        AddedDate = $HistoryReader["AddedDate"]
                        Comments = $HistoryReader["Comments"]
                        Version = $HistoryReader["Version"]
                    }
                }
                $HistoryReader.Close()
                $HistoryCommand.Dispose()
                
                $Result.History = $History
            }
            
            return $Result
        }
        catch {
            Write-IntuneInventoryLog -Message "Failed to retrieve source code: $($_.Exception.Message)" -Level Error
            throw
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Source code retrieval completed" -Level Info -Source "Get-IntuneInventorySourceCode"
    }
}

function Get-IntuneInventoryItem {
    <#
    .SYNOPSIS
    Retrieves detailed information about a specific inventory item.
    
    .DESCRIPTION
    Gets comprehensive information about an application, script, or remediation
    from the inventory database including metadata, source code, and assignments.
    
    .PARAMETER ItemId
    The ID of the inventory item to retrieve.
    
    .PARAMETER ItemType
    The type of item to retrieve (Application, Script, Remediation).
    
    .PARAMETER IncludeSourceCode
    Include source code information in the result.
    
    .PARAMETER IncludeAssignments
    Include assignment information in the result.
    
    .EXAMPLE
    Get-IntuneInventoryItem -ItemId "12345678-1234-1234-1234-123456789012" -ItemType "Application"
    
    .EXAMPLE
    Get-IntuneInventoryItem -ItemId "12345678-1234-1234-1234-123456789012" -ItemType "Script" -IncludeSourceCode -IncludeAssignments
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ItemId,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Application', 'Script', 'Remediation')]
        [string]$ItemType,
        
        [Parameter()]
        [switch]$IncludeSourceCode,
        
        [Parameter()]
        [switch]$IncludeAssignments
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Retrieving $ItemType ID: $ItemId" -Level Info -Source "Get-IntuneInventoryItem"
        
        # Verify database connection
        if (-not $Script:DatabaseConnection -or -not (Test-DatabaseConnection -DatabaseConnection $Script:DatabaseConnection)) {
            throw "Database connection is not available. Please run Connect-IntuneInventory first."
        }
    }
    
    process {
        try {
            $TableName = "${ItemType}s"
            
            # Get main item information
            $ItemCommand = $Script:DatabaseConnection.CreateCommand()
            $ItemCommand.CommandText = "SELECT * FROM $TableName WHERE Id = @ItemId;"
            $null = $ItemCommand.Parameters.AddWithValue("@ItemId", $ItemId)
            
            $Reader = $ItemCommand.ExecuteReader()
            if (-not $Reader.Read()) {
                $Reader.Close()
                $ItemCommand.Dispose()
                throw "$ItemType with ID '$ItemId' not found in inventory database."
            }
            
            # Build result object from database row
            $Result = @{}
            for ($i = 0; $i -lt $Reader.FieldCount; $i++) {
                $FieldName = $Reader.GetName($i)
                $Result[$FieldName] = if ($Reader.IsDBNull($i)) { $null } else { $Reader.GetValue($i) }
            }
            $Reader.Close()
            $ItemCommand.Dispose()
            
            # Add source code information if requested
            if ($IncludeSourceCode -and $Result.HasSourceCode) {
                try {
                    $SourceCodeInfo = Get-IntuneInventorySourceCode -ItemId $ItemId -ItemType $ItemType -IncludeHistory
                    $Result.SourceCodeDetails = $SourceCodeInfo
                }
                catch {
                    Write-IntuneInventoryLog -Message "Could not retrieve source code details: $($_.Exception.Message)" -Level Warning
                }
            }
            
            # Add assignment information if requested
            if ($IncludeAssignments) {
                try {
                    $AssignmentCommand = $Script:DatabaseConnection.CreateCommand()
                    $AssignmentCommand.CommandText = @"
SELECT * FROM Assignments 
WHERE ItemId = @ItemId AND ItemType = @ItemType;
"@
                    $null = $AssignmentCommand.Parameters.AddWithValue("@ItemId", $ItemId)
                    $null = $AssignmentCommand.Parameters.AddWithValue("@ItemType", $ItemType)
                    
                    $AssignmentReader = $AssignmentCommand.ExecuteReader()
                    $Assignments = @()
                    while ($AssignmentReader.Read()) {
                        $Assignment = @{}
                        for ($i = 0; $i -lt $AssignmentReader.FieldCount; $i++) {
                            $FieldName = $AssignmentReader.GetName($i)
                            $Assignment[$FieldName] = if ($AssignmentReader.IsDBNull($i)) { $null } else { $AssignmentReader.GetValue($i) }
                        }
                        $Assignments += $Assignment
                    }
                    $AssignmentReader.Close()
                    $AssignmentCommand.Dispose()
                    
                    $Result.Assignments = $Assignments
                }
                catch {
                    Write-IntuneInventoryLog -Message "Could not retrieve assignment details: $($_.Exception.Message)" -Level Warning
                }
            }
            
            return [PSCustomObject]$Result
        }
        catch {
            Write-IntuneInventoryLog -Message "Failed to retrieve inventory item: $($_.Exception.Message)" -Level Error
            throw
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Inventory item retrieval completed" -Level Info -Source "Get-IntuneInventoryItem"
    }
}
