function Invoke-IntuneScriptInventory {
    <#
    .SYNOPSIS
    Inventories Intune device managemen                        try {
                            $ScriptContentUri = "beta/deviceManagement/deviceManagementScripts/$($GraphScript.id)"
                            $ScriptDetails = Invoke-GraphRequest -Uri $ScriptContentUri -Method GETcripts and stores them in JSON storage.
    
    .DESCRIPTION
    Retrieves all Intune device management scripts from Microsoft Graph and stores comprehensive
    information including metadata, assignments, and source code in JSON files.
    
    .PARAMETER IncludeAssignments
    Include assignment information for each script.
    
    .PARAMETER IncludeSourceCode
    Include the actual script content in the inventory.
    
    .PARAMETER Force
    Force a complete re-inventory even if scripts already exist in storage.
    
    .EXAMPLE
    Invoke-IntuneScriptInventory
    
    .EXAMPLE
    Invoke-IntuneScriptInventory -IncludeAssignments -IncludeSourceCode -Force
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$IncludeAssignments,
        
        [Parameter()]
        [switch]$IncludeSourceCode,
        
        [Parameter()]
        [switch]$Force
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting script inventory collection" -Level Info -Source "Invoke-IntuneScriptInventory"
        
        if (-not (Test-IntuneConnection)) {
            throw "Microsoft Graph authentication is not valid. Please run Connect-IntuneInventory first."
        }
        
        if (-not (Test-StorageConnection)) {
            throw "JSON storage connection is not available. Please run Connect-IntuneInventory first."
        }
    }
    
    process {
        $StartTime = Get-Date
        $ItemsProcessed = 0
        $ErrorCount = 0
        $ErrorMessages = @()
        
        try {
            # Create inventory run record
            $RunRecord = @{
                Id = [guid]::NewGuid().ToString()
                RunType = "ScriptInventory"
                StartTime = $StartTime.ToString("yyyy-MM-dd HH:mm:ss")
                Status = "Running"
                TenantId = $Script:GraphConnection.TenantId
                UserPrincipalName = $Script:GraphConnection.UserPrincipalName
                ItemsProcessed = 0
                ErrorCount = 0
                ErrorMessages = @()
                EndTime = $null
                Duration = $null
            }
            
            Add-InventoryItem -ItemType "InventoryRuns" -Item $RunRecord
            Write-IntuneInventoryLog -Message "Created inventory run record: $($RunRecord.Id)" -Level Info
            
            # Get all device management scripts from Graph using beta endpoint
            Write-IntuneInventoryLog -Message "Retrieving scripts from Microsoft Graph (beta endpoint)" -Level Info
            $GraphScripts = Get-GraphRequestAll -Uri "beta/deviceManagement/deviceManagementScripts"
            
            if (-not $GraphScripts -or $GraphScripts.Count -eq 0) {
                Write-IntuneInventoryLog -Message "No scripts found in Intune" -Level Warning
                return @{
                    Success = $true
                    ItemsProcessed = 0
                    Message = "No scripts found"
                    RunId = $RunRecord.Id
                }
            }
            
            Write-IntuneInventoryLog -Message "Found $($GraphScripts.Count) scripts to process" -Level Info
            
            # Clear existing data if Force is specified
            if ($Force) {
                Write-IntuneInventoryLog -Message "Force parameter specified - clearing existing script data" -Level Info
                $ScriptsPath = Join-Path $Script:StorageRoot "scripts.json"
                Save-JsonData -FilePath $ScriptsPath -Data @()
            }
            
            # Process each script
            foreach ($GraphScript in $GraphScripts) {
                try {
                    Write-IntuneInventoryLog -Message "Processing script: $($GraphScript.displayName) (ID: $($GraphScript.id))" -Level Verbose
                    
                    # Get script content if requested
                    $ScriptContent = $null
                    if ($IncludeSourceCode) {
                        try {
                            $ScriptContentUri = "beta/deviceManagement/deviceManagementScripts/$($GraphScript.id)"
                            $ScriptDetails = Invoke-GraphRequest -Uri $ScriptContentUri -Method GET
                            if ($ScriptDetails.scriptContent) {
                                $ScriptContent = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($ScriptDetails.scriptContent))
                            }
                        }
                        catch {
                            Write-IntuneInventoryLog -Message "Failed to retrieve script content for '$($GraphScript.displayName)': $($_.Exception.Message)" -Level Warning
                        }
                    }
                    
                    # Create script record
                    $ScriptRecord = New-Object System.Collections.Hashtable
                    $ScriptRecord['Id'] = $GraphScript.id
                    $ScriptRecord['DisplayName'] = $GraphScript.displayName
                    $ScriptRecord['Description'] = $GraphScript.description
                    $ScriptRecord['ScriptContent'] = $ScriptContent
                    $ScriptRecord['ScriptContentEncoded'] = $GraphScript.scriptContent
                    $ScriptRecord['RunAsAccount'] = $GraphScript.runAsAccount
                    $ScriptRecord['FileName'] = $GraphScript.fileName
                    $ScriptRecord['RoleScopeTagIds'] = $GraphScript.roleScopeTagIds
                    $ScriptRecord['CreatedDateTime'] = $GraphScript.createdDateTime
                    $ScriptRecord['LastModifiedDateTime'] = $GraphScript.lastModifiedDateTime
                    $ScriptRecord['RunCount'] = $GraphScript.runCount
                    $ScriptRecord['SuccessDeviceCount'] = $GraphScript.successDeviceCount
                    $ScriptRecord['ErrorDeviceCount'] = $GraphScript.errorDeviceCount
                    $ScriptRecord['PendingDeviceCount'] = $GraphScript.pendingDeviceCount
                    $ScriptRecord['HasSourceCode'] = if ($ScriptContent) { $true } else { $false }
                    $ScriptRecord['SourceCodePath'] = $null
                    $ScriptRecord['LastInventoried'] = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    $ScriptRecord['InventoryRunId'] = $RunRecord.Id
                    
                    # Add the script to storage
                    Add-InventoryItem -ItemType "Scripts" -Item $ScriptRecord
                    
                    # Process assignments if requested
                    if ($IncludeAssignments) {
                        try {
                            Write-IntuneInventoryLog -Message "Retrieving assignments for script: $($GraphScript.displayName)" -Level Verbose
                            $AssignmentsUri = "beta/deviceManagement/deviceManagementScripts/$($GraphScript.id)/assignments"
                            $ScriptAssignments = Get-GraphRequestAll -Uri $AssignmentsUri
                            
                            if ($ScriptAssignments -and $ScriptAssignments.Count -gt 0) {
                                # Resolve group IDs to display names
                                $ScriptAssignments = Resolve-GroupAssignments -Assignments $ScriptAssignments
                                
                                foreach ($Assignment in $ScriptAssignments) {
                                    $AssignmentRecord = @{
                                        Id = $Assignment.id
                                        ObjectId = $GraphScript.id
                                        ObjectType = "Script"
                                        ObjectName = $GraphScript.displayName
                                        Target = ConvertTo-SafeJson -InputObject $Assignment.target
                                        TargetType = $Assignment.target.'@odata.type'
                                        TargetId = $Assignment.target.deviceAndAppManagementAssignmentFilterId
                                        TargetGroupId = $Assignment.target.groupId
                                        GroupDisplayNames = $Assignment.GroupDisplayNames
                                        TargetDisplayName = $Assignment.TargetDisplayName
                                        FilterType = $Assignment.target.deviceAndAppManagementAssignmentFilterType
                                        LastModifiedDateTime = $Assignment.lastModifiedDateTime
                                        CreatedDateTime = $Assignment.createdDateTime
                                        LastInventoried = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                                        InventoryRunId = $RunRecord.Id
                                    }
                                    
                                    Add-InventoryItem -ItemType "Assignments" -Item $AssignmentRecord
                                }
                                
                                Write-IntuneInventoryLog -Message "Processed $($ScriptAssignments.Count) assignments for script: $($GraphScript.displayName)" -Level Verbose
                            }
                        }
                        catch {
                            $ErrorCount++
                            $ErrorMessage = "Failed to retrieve assignments for script '$($GraphScript.displayName)': $($_.Exception.Message)"
                            $ErrorMessages += $ErrorMessage
                            Write-IntuneInventoryLog -Message $ErrorMessage -Level Warning
                        }
                    }
                    
                    $ItemsProcessed++
                    
                    if ($ItemsProcessed % 10 -eq 0) {
                        Write-IntuneInventoryLog -Message "Processed $ItemsProcessed scripts so far..." -Level Info
                    }
                }
                catch {
                    $ErrorCount++
                    $ErrorMessage = "Failed to process script '$($GraphScript.displayName)': $($_.Exception.Message)"
                    $ErrorMessages += $ErrorMessage
                    Write-IntuneInventoryLog -Message $ErrorMessage -Level Error
                }
            }
            
            # Update inventory run record
            $EndTime = Get-Date
            $Duration = ($EndTime - $StartTime).TotalMinutes
            
            $RunRecord.Status = if ($ErrorCount -eq 0) { "Completed" } else { "CompletedWithErrors" }
            $RunRecord.ItemsProcessed = $ItemsProcessed
            $RunRecord.ErrorCount = $ErrorCount
            $RunRecord.ErrorMessages = $ErrorMessages
            $RunRecord.EndTime = $EndTime.ToString("yyyy-MM-dd HH:mm:ss")
            $RunRecord.Duration = [math]::Round($Duration, 2)
            
            Add-InventoryItem -ItemType "InventoryRuns" -Item $RunRecord
            
            $SuccessMessage = "Script inventory completed successfully. Processed $ItemsProcessed scripts"
            if ($ErrorCount -gt 0) {
                $SuccessMessage += " with $ErrorCount errors"
            }
            
            Write-IntuneInventoryLog -Message $SuccessMessage -Level Info
            
            return @{
                Success = $true
                ItemsProcessed = $ItemsProcessed
                ErrorCount = $ErrorCount
                ErrorMessages = $ErrorMessages
                Duration = $Duration
                RunId = $RunRecord.Id
                Message = $SuccessMessage
            }
        }
        catch {
            $ErrorMessage = "Critical error during script inventory: $($_.Exception.Message)"
            Write-IntuneInventoryLog -Message $ErrorMessage -Level Error
            
            # Update run record with error status
            if ($RunRecord) {
                $RunRecord.Status = "Failed"
                $RunRecord.ErrorCount = $ErrorCount + 1
                $RunRecord.ErrorMessages = $ErrorMessages + $ErrorMessage
                $RunRecord.EndTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                $RunRecord.Duration = [math]::Round(((Get-Date) - $StartTime).TotalMinutes, 2)
                
                try {
                    Add-InventoryItem -ItemType "InventoryRuns" -Item $RunRecord
                }
                catch {
                    Write-IntuneInventoryLog -Message "Failed to update run record: $($_.Exception.Message)" -Level Error
                }
            }
            
            throw $ErrorMessage
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Script inventory operation completed" -Level Info -Source "Invoke-IntuneScriptInventory"
    }
}

function Invoke-IntuneRemediationInventory {
    <#
    .SYNOPSIS
    Inventories Intune remediation scripts and stores them in JSON storage.
    
    .DESCRIPTION
    Retrieves all Intune remediation scripts (proactive remediations) from Microsoft Graph and stores comprehensive
    information including metadata, assignments, and source code in JSON files.
    
    .PARAMETER IncludeAssignments
    Include assignment information for each remediation.
    
    .PARAMETER IncludeSourceCode
    Include the actual script content (detection and remediation) in the inventory.
    
    .PARAMETER Force
    Force a complete re-inventory even if remediations already exist in storage.
    
    .EXAMPLE
    Invoke-IntuneRemediationInventory
    
    .EXAMPLE
    Invoke-IntuneRemediationInventory -IncludeAssignments -IncludeSourceCode -Force
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$IncludeAssignments,
        
        [Parameter()]
        [switch]$IncludeSourceCode,
        
        [Parameter()]
        [switch]$Force
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting remediation inventory collection" -Level Info -Source "Invoke-IntuneRemediationInventory"
        
        if (-not (Test-IntuneConnection)) {
            throw "Microsoft Graph authentication is not valid. Please run Connect-IntuneInventory first."
        }
        
        if (-not (Test-StorageConnection)) {
            throw "JSON storage connection is not available. Please run Connect-IntuneInventory first."
        }
    }
    
    process {
        $StartTime = Get-Date
        $ItemsProcessed = 0
        $ErrorCount = 0
        $ErrorMessages = @()
        
        try {
            # Create inventory run record
            $RunRecord = @{
                Id = [guid]::NewGuid().ToString()
                RunType = "RemediationInventory"
                StartTime = $StartTime.ToString("yyyy-MM-dd HH:mm:ss")
                Status = "Running"
                TenantId = $Script:GraphConnection.TenantId
                UserPrincipalName = $Script:GraphConnection.UserPrincipalName
                ItemsProcessed = 0
                ErrorCount = 0
                ErrorMessages = @()
                EndTime = $null
                Duration = $null
            }
            
            Add-InventoryItem -ItemType "InventoryRuns" -Item $RunRecord
            Write-IntuneInventoryLog -Message "Created inventory run record: $($RunRecord.Id)" -Level Info
            
            # Get all device health scripts (remediations) from Graph using beta endpoint
            Write-IntuneInventoryLog -Message "Retrieving remediations from Microsoft Graph (beta endpoint)" -Level Info
            $GraphRemediations = Get-GraphRequestAll -Uri "beta/deviceManagement/deviceHealthScripts"
            
            if (-not $GraphRemediations -or $GraphRemediations.Count -eq 0) {
                Write-IntuneInventoryLog -Message "No remediations found in Intune" -Level Warning
                return @{
                    Success = $true
                    ItemsProcessed = 0
                    Message = "No remediations found"
                    RunId = $RunRecord.Id
                }
            }
            
            Write-IntuneInventoryLog -Message "Found $($GraphRemediations.Count) remediations to process" -Level Info
            
            # Clear existing data if Force is specified
            if ($Force) {
                Write-IntuneInventoryLog -Message "Force parameter specified - clearing existing remediation data" -Level Info
                $RemediationsPath = Join-Path $Script:StorageRoot "remediations.json"
                Save-JsonData -FilePath $RemediationsPath -Data @()
            }
            
            # Process each remediation
            foreach ($GraphRemediation in $GraphRemediations) {
                try {
                    Write-IntuneInventoryLog -Message "Processing remediation: $($GraphRemediation.displayName) (ID: $($GraphRemediation.id))" -Level Verbose
                    
                    # Get remediation content if requested
                    $DetectionScriptContent = $null
                    $RemediationScriptContent = $null
                    if ($IncludeSourceCode) {
                        try {
                            $RemediationDetailsUri = "beta/deviceManagement/deviceHealthScripts/$($GraphRemediation.id)"
                            $RemediationDetails = Invoke-GraphRequest -Uri $RemediationDetailsUri -Method GET
                            
                            if ($RemediationDetails.detectionScriptContent) {
                                $DetectionScriptContent = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($RemediationDetails.detectionScriptContent))
                            }
                            if ($RemediationDetails.remediationScriptContent) {
                                $RemediationScriptContent = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($RemediationDetails.remediationScriptContent))
                            }
                        }
                        catch {
                            Write-IntuneInventoryLog -Message "Failed to retrieve script content for '$($GraphRemediation.displayName)': $($_.Exception.Message)" -Level Warning
                        }
                    }
                    
                    # Create remediation record
                    $RemediationRecord = New-Object System.Collections.Hashtable
                    $RemediationRecord['Id'] = $GraphRemediation.id
                    $RemediationRecord['DisplayName'] = $GraphRemediation.displayName
                    $RemediationRecord['Description'] = $GraphRemediation.description
                    $RemediationRecord['DetectionScriptContent'] = $DetectionScriptContent
                    $RemediationRecord['RemediationScriptContent'] = $RemediationScriptContent
                    $RemediationRecord['DetectionScriptContentEncoded'] = $GraphRemediation.detectionScriptContent
                    $RemediationRecord['RemediationScriptContentEncoded'] = $GraphRemediation.remediationScriptContent
                    $RemediationRecord['RunAsAccount'] = $GraphRemediation.runAsAccount
                    $RemediationRecord['EnforceSignatureCheck'] = $GraphRemediation.enforceSignatureCheck
                    $RemediationRecord['RunAs32Bit'] = $GraphRemediation.runAs32Bit
                    $RemediationRecord['RoleScopeTagIds'] = $GraphRemediation.roleScopeTagIds
                    $RemediationRecord['IsGlobalScript'] = $GraphRemediation.isGlobalScript
                    $RemediationRecord['HighestAvailableVersion'] = $GraphRemediation.highestAvailableVersion
                    $RemediationRecord['Publisher'] = $GraphRemediation.publisher
                    $RemediationRecord['Version'] = $GraphRemediation.version
                    $RemediationRecord['CreatedDateTime'] = $GraphRemediation.createdDateTime
                    $RemediationRecord['LastModifiedDateTime'] = $GraphRemediation.lastModifiedDateTime
                    $RemediationRecord['RunSummary'] = $GraphRemediation.runSummary
                    $RemediationRecord['DeviceRunStates'] = $GraphRemediation.deviceRunStates
                    $RemediationRecord['HasSourceCode'] = if ($DetectionScriptContent -or $RemediationScriptContent) { $true } else { $false }
                    $RemediationRecord['SourceCodePath'] = $null
                    $RemediationRecord['LastInventoried'] = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    $RemediationRecord['InventoryRunId'] = $RunRecord.Id
                    
                    # Add the remediation to storage
                    Add-InventoryItem -ItemType "Remediations" -Item $RemediationRecord
                    
                    # Process assignments if requested
                    if ($IncludeAssignments) {
                        try {
                            Write-IntuneInventoryLog -Message "Retrieving assignments for remediation: $($GraphRemediation.displayName)" -Level Verbose
                            $AssignmentsUri = "beta/deviceManagement/deviceHealthScripts/$($GraphRemediation.id)/assignments"
                            $RemediationAssignments = Get-GraphRequestAll -Uri $AssignmentsUri
                            
                            if ($RemediationAssignments -and $RemediationAssignments.Count -gt 0) {
                                # Resolve group IDs to display names
                                $RemediationAssignments = Resolve-GroupAssignments -Assignments $RemediationAssignments
                                
                                foreach ($Assignment in $RemediationAssignments) {
                                    $AssignmentRecord = @{
                                        Id = $Assignment.id
                                        ObjectId = $GraphRemediation.id
                                        ObjectType = "Remediation"
                                        ObjectName = $GraphRemediation.displayName
                                        Target = ConvertTo-SafeJson -InputObject $Assignment.target
                                        TargetType = $Assignment.target.'@odata.type'
                                        TargetId = $Assignment.target.deviceAndAppManagementAssignmentFilterId
                                        TargetGroupId = $Assignment.target.groupId
                                        GroupDisplayNames = $Assignment.GroupDisplayNames
                                        TargetDisplayName = $Assignment.TargetDisplayName
                                        FilterType = $Assignment.target.deviceAndAppManagementAssignmentFilterType
                                        RunRemediationScript = $Assignment.runRemediationScript
                                        RunSchedule = $Assignment.runSchedule
                                        LastModifiedDateTime = $Assignment.lastModifiedDateTime
                                        CreatedDateTime = $Assignment.createdDateTime
                                        LastInventoried = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                                        InventoryRunId = $RunRecord.Id
                                    }
                                    
                                    Add-InventoryItem -ItemType "Assignments" -Item $AssignmentRecord
                                }
                                
                                Write-IntuneInventoryLog -Message "Processed $($RemediationAssignments.Count) assignments for remediation: $($GraphRemediation.displayName)" -Level Verbose
                            }
                        }
                        catch {
                            $ErrorCount++
                            $ErrorMessage = "Failed to retrieve assignments for remediation '$($GraphRemediation.displayName)': $($_.Exception.Message)"
                            $ErrorMessages += $ErrorMessage
                            Write-IntuneInventoryLog -Message $ErrorMessage -Level Warning
                        }
                    }
                    
                    $ItemsProcessed++
                    
                    if ($ItemsProcessed % 10 -eq 0) {
                        Write-IntuneInventoryLog -Message "Processed $ItemsProcessed remediations so far..." -Level Info
                    }
                }
                catch {
                    $ErrorCount++
                    $ErrorMessage = "Failed to process remediation '$($GraphRemediation.displayName)': $($_.Exception.Message)"
                    $ErrorMessages += $ErrorMessage
                    Write-IntuneInventoryLog -Message $ErrorMessage -Level Error
                }
            }
            
            # Update inventory run record
            $EndTime = Get-Date
            $Duration = ($EndTime - $StartTime).TotalMinutes
            
            $RunRecord.Status = if ($ErrorCount -eq 0) { "Completed" } else { "CompletedWithErrors" }
            $RunRecord.ItemsProcessed = $ItemsProcessed
            $RunRecord.ErrorCount = $ErrorCount
            $RunRecord.ErrorMessages = $ErrorMessages
            $RunRecord.EndTime = $EndTime.ToString("yyyy-MM-dd HH:mm:ss")
            $RunRecord.Duration = [math]::Round($Duration, 2)
            
            Add-InventoryItem -ItemType "InventoryRuns" -Item $RunRecord
            
            $SuccessMessage = "Remediation inventory completed successfully. Processed $ItemsProcessed remediations"
            if ($ErrorCount -gt 0) {
                $SuccessMessage += " with $ErrorCount errors"
            }
            
            Write-IntuneInventoryLog -Message $SuccessMessage -Level Info
            
            return @{
                Success = $true
                ItemsProcessed = $ItemsProcessed
                ErrorCount = $ErrorCount
                ErrorMessages = $ErrorMessages
                Duration = $Duration
                RunId = $RunRecord.Id
                Message = $SuccessMessage
            }
        }
        catch {
            $ErrorMessage = "Critical error during remediation inventory: $($_.Exception.Message)"
            Write-IntuneInventoryLog -Message $ErrorMessage -Level Error
            
            # Update run record with error status
            if ($RunRecord) {
                $RunRecord.Status = "Failed"
                $RunRecord.ErrorCount = $ErrorCount + 1
                $RunRecord.ErrorMessages = $ErrorMessages + $ErrorMessage
                $RunRecord.EndTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                $RunRecord.Duration = [math]::Round(((Get-Date) - $StartTime).TotalMinutes, 2)
                
                try {
                    Add-InventoryItem -ItemType "InventoryRuns" -Item $RunRecord
                }
                catch {
                    Write-IntuneInventoryLog -Message "Failed to update run record: $($_.Exception.Message)" -Level Error
                }
            }
            
            throw $ErrorMessage
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Remediation inventory operation completed" -Level Info -Source "Invoke-IntuneRemediationInventory"
    }
}
