function Invoke-IntuneApplicationInventory {
    <#
    .SYNOPSIS
    Inventories Intune applications and stores them in JSON storage.
    
    .DESCRIPTION
    Retrieves all Intune applications from Microsoft Graph and stores comprehensive
    information including metadata, assignments, and source code indicators in JSON files.
    
    .PARAMETER IncludeAssignments
    Include assignment information for each application.
    
    .PARAMETER Force
    Force a complete re-inventory even if applications already exist in storage.
    
    .EXAMPLE
    Invoke-IntuneApplicationInventory
    
    .EXAMPLE
    Invoke-IntuneApplicationInventory -IncludeAssignments -Force
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$IncludeAssignments,
        
        [Parameter()]
        [switch]$Force
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting application inventory collection" -Level Info -Source "Invoke-IntuneApplicationInventory"
        
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
                RunType = "ApplicationInventory"
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
            
            # Get all mobile apps from Graph using beta endpoint for more complete data
            Write-IntuneInventoryLog -Message "Retrieving applications from Microsoft Graph (beta endpoint)" -Level Info
            $GraphApps = Get-GraphRequestAll -Uri "beta/deviceAppManagement/mobileApps"
            
            if (-not $GraphApps -or $GraphApps.Count -eq 0) {
                Write-IntuneInventoryLog -Message "No applications found in Intune" -Level Warning
                return @{
                    Success = $true
                    ItemsProcessed = 0
                    Message = "No applications found"
                    RunId = $RunRecord.Id
                }
            }
            
            Write-IntuneInventoryLog -Message "Found $($GraphApps.Count) applications to process" -Level Info
            
            # Clear existing data if Force is specified
            if ($Force) {
                Write-IntuneInventoryLog -Message "Force parameter specified - clearing existing application data" -Level Info
                $ApplicationsPath = Join-Path $Script:StorageRoot "applications.json"
                Save-JsonData -FilePath $ApplicationsPath -Data @()
            }
            
            # Process each application
            foreach ($GraphApp in $GraphApps) {
                try {
                    Write-IntuneInventoryLog -Message "Processing application: $($GraphApp.displayName) (ID: $($GraphApp.id))" -Level Verbose
                    
                    # Create application record with safe property mapping
                    $AppRecord = New-Object System.Collections.Hashtable
                    
                    # Core properties
                    $AppRecord['Id'] = $GraphApp.id
                    $AppRecord['DisplayName'] = $GraphApp.displayName
                    $AppRecord['Description'] = $GraphApp.description
                    $AppRecord['Publisher'] = $GraphApp.publisher
                    $AppRecord['AppType'] = $GraphApp.'@odata.type'
                    $AppRecord['CreatedDateTime'] = $GraphApp.createdDateTime
                    $AppRecord['LastModifiedDateTime'] = $GraphApp.lastModifiedDateTime
                    $AppRecord['IsFeatured'] = $GraphApp.isFeatured
                    $AppRecord['PrivacyInformationUrl'] = $GraphApp.privacyInformationUrl
                    $AppRecord['InformationUrl'] = $GraphApp.informationUrl
                    $AppRecord['Owner'] = $GraphApp.owner
                    $AppRecord['Developer'] = $GraphApp.developer
                    $AppRecord['Notes'] = $GraphApp.notes
                    $AppRecord['UploadState'] = $GraphApp.uploadState
                    $AppRecord['PublishingState'] = $GraphApp.publishingState
                    $AppRecord['IsAssigned'] = $GraphApp.isAssigned
                    $AppRecord['RoleScopeTagIds'] = $GraphApp.roleScopeTagIds
                    $AppRecord['DependentAppCount'] = $GraphApp.dependentAppCount
                    $AppRecord['SupersedingAppCount'] = $GraphApp.supersedingAppCount
                    $AppRecord['SupersededAppCount'] = $GraphApp.supersededAppCount
                    $AppRecord['CommittedContentVersion'] = $GraphApp.committedContentVersion
                    $AppRecord['FileName'] = $GraphApp.fileName
                    $AppRecord['Size'] = $GraphApp.size
                    $AppRecord['InstallCommandLine'] = $GraphApp.installCommandLine
                    $AppRecord['UninstallCommandLine'] = $GraphApp.uninstallCommandLine
                    $AppRecord['ApplicableArchitectures'] = $GraphApp.applicableArchitectures
                    $AppRecord['MinimumFreeDiskSpaceInMB'] = $GraphApp.minimumFreeDiskSpaceInMB
                    $AppRecord['MinimumMemoryInMB'] = $GraphApp.minimumMemoryInMB
                    $AppRecord['MinimumNumberOfProcessors'] = $GraphApp.minimumNumberOfProcessors
                    $AppRecord['MinimumCpuSpeedInMHz'] = $GraphApp.minimumCpuSpeedInMHz
                    $AppRecord['Rules'] = $GraphApp.rules
                    $AppRecord['InstallExperience'] = $GraphApp.installExperience
                    $AppRecord['ReturnCodes'] = $GraphApp.returnCodes
                    $AppRecord['MsiInformation'] = $GraphApp.msiInformation
                    $AppRecord['SetupFilePath'] = $GraphApp.setupFilePath
                    $AppRecord['MinimumSupportedWindowsRelease'] = $GraphApp.minimumSupportedWindowsRelease
                    $AppRecord['Categories'] = $GraphApp.categories
                    $AppRecord['AllowAvailableUninstall'] = $GraphApp.allowAvailableUninstall
                    $AppRecord['InstalledDateTime'] = $GraphApp.installedDateTime
                    $AppRecord['PackageId'] = $GraphApp.packageId
                    $AppRecord['AppId'] = $GraphApp.appId
                    $AppRecord['BundleId'] = $GraphApp.bundleId
                    $AppRecord['MinimumSupportedOperatingSystem'] = $GraphApp.minimumSupportedOperatingSystem
                    $AppRecord['ProductKey'] = $GraphApp.productKey
                    $AppRecord['LicenseType'] = $GraphApp.licenseType
                    $AppRecord['ProductIds'] = $GraphApp.productIds
                    $AppRecord['SkuIds'] = $GraphApp.skuIds
                    $AppRecord['ProductVersion'] = $GraphApp.productVersion
                    $AppRecord['DisplayVersion'] = $GraphApp.displayVersion
                    $AppRecord['IconUrl'] = if ($GraphApp.largeIcon) { $GraphApp.largeIcon.value } else { $null }
                    
                    # Custom inventory properties
                    $AppRecord['HasSourceCode'] = $false
                    $AppRecord['SourceCodePath'] = $null
                    $AppRecord['LastInventoried'] = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    $AppRecord['InventoryRunId'] = $RunRecord.Id
                    
                    # Add the application to storage
                    Add-InventoryItem -ItemType "Applications" -Item $AppRecord
                    
                    # Process assignments if requested
                    if ($IncludeAssignments) {
                        try {
                            Write-IntuneInventoryLog -Message "Retrieving assignments for application: $($GraphApp.displayName)" -Level Verbose
                            $AssignmentsUri = "beta/deviceAppManagement/mobileApps/$($GraphApp.id)/assignments"
                            $AppAssignments = Get-GraphRequestAll -Uri $AssignmentsUri
                            
                            if ($AppAssignments -and $AppAssignments.Count -gt 0) {
                                # Resolve group IDs to display names
                                $AppAssignments = Resolve-GroupAssignments -Assignments $AppAssignments
                                
                                foreach ($Assignment in $AppAssignments) {
                                    $AssignmentRecord = @{
                                        Id = $Assignment.id
                                        ObjectId = $GraphApp.id
                                        ObjectType = "Application"
                                        ObjectName = $GraphApp.displayName
                                        Intent = $Assignment.intent
                                        Settings = $Assignment.settings
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
                                
                                Write-IntuneInventoryLog -Message "Processed $($AppAssignments.Count) assignments for application: $($GraphApp.displayName)" -Level Verbose
                            }
                        }
                        catch {
                            $ErrorCount++
                            $ErrorMessage = "Failed to retrieve assignments for application '$($GraphApp.displayName)': $($_.Exception.Message)"
                            $ErrorMessages += $ErrorMessage
                            Write-IntuneInventoryLog -Message $ErrorMessage -Level Warning
                        }
                    }
                    
                    $ItemsProcessed++
                    
                    if ($ItemsProcessed % 10 -eq 0) {
                        Write-IntuneInventoryLog -Message "Processed $ItemsProcessed applications so far..." -Level Info
                    }
                }
                catch {
                    $ErrorCount++
                    $ErrorMessage = "Failed to process application '$($GraphApp.displayName)': $($_.Exception.Message)"
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
            
            $SuccessMessage = "Application inventory completed successfully. Processed $ItemsProcessed applications"
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
            $ErrorMessage = "Critical error during application inventory: $($_.Exception.Message)"
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
        Write-IntuneInventoryLog -Message "Application inventory operation completed" -Level Info -Source "Invoke-IntuneApplicationInventory"
    }
}
