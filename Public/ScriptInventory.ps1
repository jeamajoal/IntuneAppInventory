function Invoke-IntuneScriptInventory {
    <#
    .SYNOPSIS
    Inventories Intune PowerShell scripts and stores them in the database.
    
    .DESCRIPTION
    Retrieves all Intune PowerShell scripts from Microsoft Graph and stores comprehensive
    information including metadata, script content, and assignments in the SQLite database.
    
    .PARAMETER IncludeAssignments
    Include assignment information for each script.
    
    .PARAMETER Force
    Force a complete re-inventory even if scripts already exist in the database.
    
    .EXAMPLE
    Invoke-IntuneScriptInventory
    
    .EXAMPLE
    Invoke-IntuneScriptInventory -IncludeAssignments -Force
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [switch]$IncludeAssignments,
        
        [Parameter()]
        [switch]$Force
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting Intune script inventory" -Level Info -Source "Invoke-IntuneScriptInventory"
        
        # Verify connections
        if (-not $Script:DatabaseConnection -or -not (Test-DatabaseConnection -DatabaseConnection $Script:DatabaseConnection)) {
            throw "Database connection is not available. Please run Connect-IntuneInventory first."
        }
        
        if (-not (Test-GraphAuthentication)) {
            throw "Microsoft Graph authentication is not valid. Please run Connect-IntuneInventory first."
        }
    }
    
    process {
        $RunId = $null
        $StartTime = Get-Date
        $ItemsProcessed = 0
        $ErrorCount = 0
        $ErrorMessages = @()
        
        try {
            # Create inventory run record
            $UserContext = Get-CurrentUserContext
            $Command = $Script:DatabaseConnection.CreateCommand()
            $Command.CommandText = @"
INSERT INTO InventoryRuns (RunType, StartTime, Status, TenantId, UserPrincipalName)
VALUES (@RunType, @StartTime, @Status, @TenantId, @UserPrincipalName);
SELECT last_insert_rowid();
"@
            $null = $Command.Parameters.AddWithValue("@RunType", "ScriptInventory")
            $null = $Command.Parameters.AddWithValue("@StartTime", $StartTime.ToString("yyyy-MM-dd HH:mm:ss"))
            $null = $Command.Parameters.AddWithValue("@Status", "Running")
            $null = $Command.Parameters.AddWithValue("@TenantId", $UserContext.TenantId)
            $null = $Command.Parameters.AddWithValue("@UserPrincipalName", $UserContext.UserPrincipalName)
            
            $RunId = $Command.ExecuteScalar()
            $Command.Dispose()
            
            Write-IntuneInventoryLog -Message "Started script inventory run ID: $RunId" -Level Info
            
            # Get all device PowerShell scripts using Graph API
            Write-IntuneInventoryLog -Message "Retrieving Intune PowerShell scripts..." -Level Info
            $Scripts = Get-GraphRequestAll -Uri "v1.0/deviceManagement/deviceManagementScripts"
            
            Write-IntuneInventoryLog -Message "Found $($Scripts.Count) scripts to process" -Level Info
            
            foreach ($Script in $Scripts) {
                try {
                    if ($PSCmdlet.ShouldProcess($Script.displayName, "Inventory Intune script")) {
                        Write-IntuneInventoryLog -Message "Processing script: $($Script.displayName)" -Level Verbose
                        
                        # Get script content
                        $ScriptContent = ""
                        $SourceCode = ""
                        try {
                            $ScriptDetail = Invoke-GraphRequest -Uri "v1.0/deviceManagement/deviceManagementScripts/$($Script.id)" -Method GET
                            if ($ScriptDetail.scriptContent) {
                                # Script content is base64 encoded
                                $ScriptBytes = [System.Convert]::FromBase64String($ScriptDetail.scriptContent)
                                $ScriptContent = [System.Text.Encoding]::UTF8.GetString($ScriptBytes)
                                $SourceCode = $ScriptContent
                            }
                        }
                        catch {
                            Write-IntuneInventoryLog -Message "Could not retrieve script content for $($Script.displayName): $($_.Exception.Message)" -Level Warning
                        }
                        
                        # Insert or update script record
                        $InsertCommand = $Script:DatabaseConnection.CreateCommand()
                        $InsertCommand.CommandText = @"
INSERT OR REPLACE INTO Scripts (
    Id, DisplayName, Description, ScriptContent, CreatedDateTime, 
    LastModifiedDateTime, RunAsAccount, FileName, RoleScopeTagIds,
    HasSourceCode, SourceCode, SourceCodeAdded, InventoryDate, LastUpdated
) VALUES (
    @Id, @DisplayName, @Description, @ScriptContent, @CreatedDateTime,
    @LastModifiedDateTime, @RunAsAccount, @FileName, @RoleScopeTagIds,
    @HasSourceCode, @SourceCode, @SourceCodeAdded, @InventoryDate, @LastUpdated
);
"@
                        
                        # Add parameters
                        $HasSourceCode = if ([string]::IsNullOrWhiteSpace($SourceCode)) { 0 } else { 1 }
                        
                        $null = $InsertCommand.Parameters.AddWithValue("@Id", $Script.id)
                        $null = $InsertCommand.Parameters.AddWithValue("@DisplayName", $Script.displayName ?? "")
                        $null = $InsertCommand.Parameters.AddWithValue("@Description", $Script.description ?? "")
                        $null = $InsertCommand.Parameters.AddWithValue("@ScriptContent", $ScriptContent)
                        $null = $InsertCommand.Parameters.AddWithValue("@CreatedDateTime", $Script.createdDateTime ?? "")
                        $null = $InsertCommand.Parameters.AddWithValue("@LastModifiedDateTime", $Script.lastModifiedDateTime ?? "")
                        $null = $InsertCommand.Parameters.AddWithValue("@RunAsAccount", $Script.runAsAccount ?? "")
                        $null = $InsertCommand.Parameters.AddWithValue("@FileName", $Script.fileName ?? "")
                        $null = $InsertCommand.Parameters.AddWithValue("@RoleScopeTagIds", ($Script.roleScopeTagIds -join ", ") ?? "")
                        $null = $InsertCommand.Parameters.AddWithValue("@HasSourceCode", $HasSourceCode)
                        $null = $InsertCommand.Parameters.AddWithValue("@SourceCode", $SourceCode)
                        $null = $InsertCommand.Parameters.AddWithValue("@SourceCodeAdded", $HasSourceCode -eq 1 ? $StartTime.ToString("yyyy-MM-dd HH:mm:ss") : "")
                        $null = $InsertCommand.Parameters.AddWithValue("@InventoryDate", $StartTime.ToString("yyyy-MM-dd HH:mm:ss"))
                        $null = $InsertCommand.Parameters.AddWithValue("@LastUpdated", $StartTime.ToString("yyyy-MM-dd HH:mm:ss"))
                        
                        $null = $InsertCommand.ExecuteNonQuery()
                        $InsertCommand.Dispose()
                        
                        # Process assignments if requested
                        if ($IncludeAssignments) {
                            try {
                                $Assignments = Get-GraphRequestAll -Uri "v1.0/deviceManagement/deviceManagementScripts/$($Script.id)/assignments"
                                foreach ($Assignment in $Assignments) {
                                    # Process assignment (similar to application assignments)
                                    Write-IntuneInventoryLog -Message "Processing assignment for script: $($Script.displayName)" -Level Verbose
                                    # Add assignment processing logic here
                                }
                            }
                            catch {
                                Write-IntuneInventoryLog -Message "Could not retrieve assignments for script $($Script.displayName): $($_.Exception.Message)" -Level Warning
                                $ErrorCount++
                                $ErrorMessages += "Assignment error for script $($Script.displayName): $($_.Exception.Message)"
                            }
                        }
                        
                        $ItemsProcessed++
                        
                        if ($ItemsProcessed % 5 -eq 0) {
                            Write-Progress -Activity "Inventorying Scripts" -Status "Processed $ItemsProcessed of $($Scripts.Count)" -PercentComplete (($ItemsProcessed / $Scripts.Count) * 100)
                        }
                    }
                }
                catch {
                    Write-IntuneInventoryLog -Message "Error processing script $($Script.displayName): $($_.Exception.Message)" -Level Error
                    $ErrorCount++
                    $ErrorMessages += "Error processing script $($Script.displayName): $($_.Exception.Message)"
                }
            }
            
            Write-Progress -Activity "Inventorying Scripts" -Completed
            
            # Update inventory run record
            $EndTime = Get-Date
            $UpdateCommand = $Script:DatabaseConnection.CreateCommand()
            $UpdateCommand.CommandText = @"
UPDATE InventoryRuns 
SET EndTime = @EndTime, Status = @Status, ItemsProcessed = @ItemsProcessed, 
    ErrorCount = @ErrorCount, ErrorMessages = @ErrorMessages
WHERE Id = @RunId;
"@
            $null = $UpdateCommand.Parameters.AddWithValue("@EndTime", $EndTime.ToString("yyyy-MM-dd HH:mm:ss"))
            $null = $UpdateCommand.Parameters.AddWithValue("@Status", "Completed")
            $null = $UpdateCommand.Parameters.AddWithValue("@ItemsProcessed", $ItemsProcessed)
            $null = $UpdateCommand.Parameters.AddWithValue("@ErrorCount", $ErrorCount)
            $null = $UpdateCommand.Parameters.AddWithValue("@ErrorMessages", ($ErrorMessages -join "; "))
            $null = $UpdateCommand.Parameters.AddWithValue("@RunId", $RunId)
            
            $null = $UpdateCommand.ExecuteNonQuery()
            $UpdateCommand.Dispose()
            
            Write-IntuneInventoryLog -Message "Script inventory completed. Processed: $ItemsProcessed, Errors: $ErrorCount" -Level Info
            Write-Host "Script inventory completed!" -ForegroundColor Green
            Write-Host "Scripts processed: $ItemsProcessed" -ForegroundColor Cyan
            if ($ErrorCount -gt 0) {
                Write-Host "Errors encountered: $ErrorCount" -ForegroundColor Yellow
            }
        }
        catch {
            # Update inventory run with error status
            if ($RunId) {
                try {
                    $ErrorCommand = $Script:DatabaseConnection.CreateCommand()
                    $ErrorCommand.CommandText = @"
UPDATE InventoryRuns 
SET EndTime = @EndTime, Status = @Status, ItemsProcessed = @ItemsProcessed, 
    ErrorCount = @ErrorCount, ErrorMessages = @ErrorMessages
WHERE Id = @RunId;
"@
                    $null = $ErrorCommand.Parameters.AddWithValue("@EndTime", (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
                    $null = $ErrorCommand.Parameters.AddWithValue("@Status", "Failed")
                    $null = $ErrorCommand.Parameters.AddWithValue("@ItemsProcessed", $ItemsProcessed)
                    $null = $ErrorCommand.Parameters.AddWithValue("@ErrorCount", $ErrorCount + 1)
                    $null = $ErrorCommand.Parameters.AddWithValue("@ErrorMessages", ($ErrorMessages + $_.Exception.Message) -join "; ")
                    $null = $ErrorCommand.Parameters.AddWithValue("@RunId", $RunId)
                    
                    $null = $ErrorCommand.ExecuteNonQuery()
                    $ErrorCommand.Dispose()
                }
                catch {
                    Write-IntuneInventoryLog -Message "Failed to update inventory run with error status: $($_.Exception.Message)" -Level Warning
                }
            }
            
            Write-IntuneInventoryLog -Message "Script inventory failed: $($_.Exception.Message)" -Level Error
            throw
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Script inventory process completed" -Level Info -Source "Invoke-IntuneScriptInventory"
    }
}

function Invoke-IntuneRemediationInventory {
    <#
    .SYNOPSIS
    Inventories Intune remediation scripts (proactive remediations) and stores them in the database.
    
    .DESCRIPTION
    Retrieves all Intune remediation scripts from Microsoft Graph and stores comprehensive
    information including metadata, detection and remediation script content, and assignments.
    
    .PARAMETER IncludeAssignments
    Include assignment information for each remediation.
    
    .PARAMETER Force
    Force a complete re-inventory even if remediations already exist in the database.
    
    .EXAMPLE
    Invoke-IntuneRemediationInventory
    
    .EXAMPLE
    Invoke-IntuneRemediationInventory -IncludeAssignments -Force
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [switch]$IncludeAssignments,
        
        [Parameter()]
        [switch]$Force
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting Intune remediation inventory" -Level Info -Source "Invoke-IntuneRemediationInventory"
        
        # Verify connections
        if (-not $Script:DatabaseConnection -or -not (Test-DatabaseConnection -DatabaseConnection $Script:DatabaseConnection)) {
            throw "Database connection is not available. Please run Connect-IntuneInventory first."
        }
        
        if (-not (Test-GraphAuthentication)) {
            throw "Microsoft Graph authentication is not valid. Please run Connect-IntuneInventory first."
        }
    }
    
    process {
        $RunId = $null
        $StartTime = Get-Date
        $ItemsProcessed = 0
        $ErrorCount = 0
        $ErrorMessages = @()
        
        try {
            # Create inventory run record
            $UserContext = Get-CurrentUserContext
            $Command = $Script:DatabaseConnection.CreateCommand()
            $Command.CommandText = @"
INSERT INTO InventoryRuns (RunType, StartTime, Status, TenantId, UserPrincipalName)
VALUES (@RunType, @StartTime, @Status, @TenantId, @UserPrincipalName);
SELECT last_insert_rowid();
"@
            $null = $Command.Parameters.AddWithValue("@RunType", "RemediationInventory")
            $null = $Command.Parameters.AddWithValue("@StartTime", $StartTime.ToString("yyyy-MM-dd HH:mm:ss"))
            $null = $Command.Parameters.AddWithValue("@Status", "Running")
            $null = $Command.Parameters.AddWithValue("@TenantId", $UserContext.TenantId)
            $null = $Command.Parameters.AddWithValue("@UserPrincipalName", $UserContext.UserPrincipalName)
            
            $RunId = $Command.ExecuteScalar()
            $Command.Dispose()
            
            Write-IntuneInventoryLog -Message "Started remediation inventory run ID: $RunId" -Level Info
            
            # Get all device health scripts (remediations) using Graph API
            Write-IntuneInventoryLog -Message "Retrieving Intune remediation scripts..." -Level Info
            $Remediations = Get-GraphRequestAll -Uri "v1.0/deviceManagement/deviceHealthScripts"
            
            Write-IntuneInventoryLog -Message "Found $($Remediations.Count) remediations to process" -Level Info
            
            foreach ($Remediation in $Remediations) {
                try {
                    if ($PSCmdlet.ShouldProcess($Remediation.displayName, "Inventory Intune remediation")) {
                        Write-IntuneInventoryLog -Message "Processing remediation: $($Remediation.displayName)" -Level Verbose
                        
                        # Get remediation script content
                        $DetectionScriptContent = ""
                        $RemediationScriptContent = ""
                        $SourceCode = ""
                        
                        try {
                            # Get detection script
                            if ($Remediation.detectionScriptContent) {
                                $DetectionBytes = [System.Convert]::FromBase64String($Remediation.detectionScriptContent)
                                $DetectionScriptContent = [System.Text.Encoding]::UTF8.GetString($DetectionBytes)
                            }
                            
                            # Get remediation script
                            if ($Remediation.remediationScriptContent) {
                                $RemediationBytes = [System.Convert]::FromBase64String($Remediation.remediationScriptContent)
                                $RemediationScriptContent = [System.Text.Encoding]::UTF8.GetString($RemediationBytes)
                            }
                            
                            # Combine scripts for source code
                            if ($DetectionScriptContent -or $RemediationScriptContent) {
                                $SourceCode = @{
                                    DetectionScript = $DetectionScriptContent
                                    RemediationScript = $RemediationScriptContent
                                } | ConvertTo-SafeJson
                            }
                        }
                        catch {
                            Write-IntuneInventoryLog -Message "Could not retrieve script content for remediation $($Remediation.displayName): $($_.Exception.Message)" -Level Warning
                        }
                        
                        # Insert or update remediation record
                        $InsertCommand = $Script:DatabaseConnection.CreateCommand()
                        $InsertCommand.CommandText = @"
INSERT OR REPLACE INTO Remediations (
    Id, DisplayName, Description, Publisher, Version, CreatedDateTime, 
    LastModifiedDateTime, DetectionScriptContent, RemediationScriptContent, 
    RunAsAccount, RoleScopeTagIds, HasSourceCode, SourceCode, SourceCodeAdded, 
    InventoryDate, LastUpdated
) VALUES (
    @Id, @DisplayName, @Description, @Publisher, @Version, @CreatedDateTime,
    @LastModifiedDateTime, @DetectionScriptContent, @RemediationScriptContent,
    @RunAsAccount, @RoleScopeTagIds, @HasSourceCode, @SourceCode, @SourceCodeAdded,
    @InventoryDate, @LastUpdated
);
"@
                        
                        # Add parameters
                        $HasSourceCode = if ([string]::IsNullOrWhiteSpace($SourceCode)) { 0 } else { 1 }
                        
                        $null = $InsertCommand.Parameters.AddWithValue("@Id", $Remediation.id)
                        $null = $InsertCommand.Parameters.AddWithValue("@DisplayName", $Remediation.displayName ?? "")
                        $null = $InsertCommand.Parameters.AddWithValue("@Description", $Remediation.description ?? "")
                        $null = $InsertCommand.Parameters.AddWithValue("@Publisher", $Remediation.publisher ?? "")
                        $null = $InsertCommand.Parameters.AddWithValue("@Version", $Remediation.version ?? "")
                        $null = $InsertCommand.Parameters.AddWithValue("@CreatedDateTime", $Remediation.createdDateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "")
                        $null = $InsertCommand.Parameters.AddWithValue("@LastModifiedDateTime", $Remediation.lastModifiedDateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "")
                        $null = $InsertCommand.Parameters.AddWithValue("@DetectionScriptContent", $DetectionScriptContent)
                        $null = $InsertCommand.Parameters.AddWithValue("@RemediationScriptContent", $RemediationScriptContent)
                        $null = $InsertCommand.Parameters.AddWithValue("@RunAsAccount", $Remediation.runAsAccount ?? "")
                        $null = $InsertCommand.Parameters.AddWithValue("@RoleScopeTagIds", ($Remediation.roleScopeTagIds -join ", ") ?? "")
                        $null = $InsertCommand.Parameters.AddWithValue("@HasSourceCode", $HasSourceCode)
                        $null = $InsertCommand.Parameters.AddWithValue("@SourceCode", $SourceCode)
                        $null = $InsertCommand.Parameters.AddWithValue("@SourceCodeAdded", $HasSourceCode -eq 1 ? $StartTime.ToString("yyyy-MM-dd HH:mm:ss") : "")
                        $null = $InsertCommand.Parameters.AddWithValue("@InventoryDate", $StartTime.ToString("yyyy-MM-dd HH:mm:ss"))
                        $null = $InsertCommand.Parameters.AddWithValue("@LastUpdated", $StartTime.ToString("yyyy-MM-dd HH:mm:ss"))
                        
                        $null = $InsertCommand.ExecuteNonQuery()
                        $InsertCommand.Dispose()
                        
                        # Process assignments if requested
                        if ($IncludeAssignments) {
                            try {
                                $Assignments = Get-GraphRequestAll -Uri "v1.0/deviceManagement/deviceHealthScripts/$($Remediation.id)/assignments"
                                foreach ($Assignment in $Assignments) {
                                    # Process assignment (similar to other assignments)
                                    Write-IntuneInventoryLog -Message "Processing assignment for remediation: $($Remediation.displayName)" -Level Verbose
                                    # Add assignment processing logic here
                                }
                            }
                            catch {
                                Write-IntuneInventoryLog -Message "Could not retrieve assignments for remediation $($Remediation.displayName): $($_.Exception.Message)" -Level Warning
                                $ErrorCount++
                                $ErrorMessages += "Assignment error for remediation $($Remediation.displayName): $($_.Exception.Message)"
                            }
                        }
                        
                        $ItemsProcessed++
                        
                        if ($ItemsProcessed % 5 -eq 0) {
                            Write-Progress -Activity "Inventorying Remediations" -Status "Processed $ItemsProcessed of $($Remediations.Count)" -PercentComplete (($ItemsProcessed / $Remediations.Count) * 100)
                        }
                    }
                }
                catch {
                    Write-IntuneInventoryLog -Message "Error processing remediation $($Remediation.displayName): $($_.Exception.Message)" -Level Error
                    $ErrorCount++
                    $ErrorMessages += "Error processing remediation $($Remediation.displayName): $($_.Exception.Message)"
                }
            }
            
            Write-Progress -Activity "Inventorying Remediations" -Completed
            
            # Update inventory run record
            $EndTime = Get-Date
            $UpdateCommand = $Script:DatabaseConnection.CreateCommand()
            $UpdateCommand.CommandText = @"
UPDATE InventoryRuns 
SET EndTime = @EndTime, Status = @Status, ItemsProcessed = @ItemsProcessed, 
    ErrorCount = @ErrorCount, ErrorMessages = @ErrorMessages
WHERE Id = @RunId;
"@
            $null = $UpdateCommand.Parameters.AddWithValue("@EndTime", $EndTime.ToString("yyyy-MM-dd HH:mm:ss"))
            $null = $UpdateCommand.Parameters.AddWithValue("@Status", "Completed")
            $null = $UpdateCommand.Parameters.AddWithValue("@ItemsProcessed", $ItemsProcessed)
            $null = $UpdateCommand.Parameters.AddWithValue("@ErrorCount", $ErrorCount)
            $null = $UpdateCommand.Parameters.AddWithValue("@ErrorMessages", ($ErrorMessages -join "; "))
            $null = $UpdateCommand.Parameters.AddWithValue("@RunId", $RunId)
            
            $null = $UpdateCommand.ExecuteNonQuery()
            $UpdateCommand.Dispose()
            
            Write-IntuneInventoryLog -Message "Remediation inventory completed. Processed: $ItemsProcessed, Errors: $ErrorCount" -Level Info
            Write-Host "Remediation inventory completed!" -ForegroundColor Green
            Write-Host "Remediations processed: $ItemsProcessed" -ForegroundColor Cyan
            if ($ErrorCount -gt 0) {
                Write-Host "Errors encountered: $ErrorCount" -ForegroundColor Yellow
            }
        }
        catch {
            # Update inventory run with error status
            if ($RunId) {
                try {
                    $ErrorCommand = $Script:DatabaseConnection.CreateCommand()
                    $ErrorCommand.CommandText = @"
UPDATE InventoryRuns 
SET EndTime = @EndTime, Status = @Status, ItemsProcessed = @ItemsProcessed, 
    ErrorCount = @ErrorCount, ErrorMessages = @ErrorMessages
WHERE Id = @RunId;
"@
                    $null = $ErrorCommand.Parameters.AddWithValue("@EndTime", (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
                    $null = $ErrorCommand.Parameters.AddWithValue("@Status", "Failed")
                    $null = $ErrorCommand.Parameters.AddWithValue("@ItemsProcessed", $ItemsProcessed)
                    $null = $ErrorCommand.Parameters.AddWithValue("@ErrorCount", $ErrorCount + 1)
                    $null = $ErrorCommand.Parameters.AddWithValue("@ErrorMessages", ($ErrorMessages + $_.Exception.Message) -join "; ")
                    $null = $ErrorCommand.Parameters.AddWithValue("@RunId", $RunId)
                    
                    $null = $ErrorCommand.ExecuteNonQuery()
                    $ErrorCommand.Dispose()
                }
                catch {
                    Write-IntuneInventoryLog -Message "Failed to update inventory run with error status: $($_.Exception.Message)" -Level Warning
                }
            }
            
            Write-IntuneInventoryLog -Message "Remediation inventory failed: $($_.Exception.Message)" -Level Error
            throw
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Remediation inventory process completed" -Level Info -Source "Invoke-IntuneRemediationInventory"
    }
}
