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

function Get-IntuneLobAppContent {
    <#
    .SYNOPSIS
    Downloads the actual .intunewin files for Win32 LOB applications from Intune.
    
    .DESCRIPTION
    Uses Microsoft Graph API to download the encrypted content files for Win32 LOB applications.
    Downloads the original .intunewin files that contain the application installers.
    
    .PARAMETER AppId
    The ID of the specific application to download. If not specified, downloads all Win32 LOB apps.
    
    .PARAMETER DownloadPath
    The path where downloaded content should be stored. Defaults to storage\downloaded-apps.
    
    .PARAMETER Force
    Overwrite existing downloaded files.
    
    .EXAMPLE
    Get-IntuneLobAppContent
    
    .EXAMPLE
    Get-IntuneLobAppContent -AppId "12345678-1234-1234-1234-123456789012" -DownloadPath "C:\AppBackups"
    
    .EXAMPLE
    Get-IntuneLobAppContent -Force
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$AppId,
        
        [Parameter()]
        [string]$DownloadPath,
        
        [Parameter()]
        [switch]$Force
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting LOB app content download" -Level Info -Source "Get-IntuneLobAppContent"
        
        if (-not (Test-IntuneConnection)) {
            throw "Microsoft Graph authentication is not valid. Please run Connect-IntuneInventory first."
        }
        
        if (-not (Test-StorageConnection)) {
            throw "JSON storage connection is not available. Please run Connect-IntuneInventory first."
        }
        
        # Set default download path
        if (-not $DownloadPath) {
            $DownloadPath = Join-Path $Script:StorageRoot "downloaded-apps"
        }
        
        # Ensure download directory exists
        if (-not (Test-Path $DownloadPath)) {
            New-Item -Path $DownloadPath -ItemType Directory -Force | Out-Null
            Write-IntuneInventoryLog -Message "Created download directory: $DownloadPath" -Level Info
        }
    }
    
    process {
        $StartTime = Get-Date
        $ItemsProcessed = 0
        $ErrorCount = 0
        $ErrorMessages = @()
        $DownloadedFiles = @()
        
        try {
            # Get applications from storage
            $AllApps = Get-InventoryItems -ItemType "Applications"
            
            # Filter to Win32 LOB apps
            if ($AppId) {
                $LobApps = $AllApps | Where-Object { $_.Id -eq $AppId -and $_.AppType -eq '#microsoft.graph.win32LobApp' }
                if (-not $LobApps) {
                    throw "App with ID '$AppId' not found or is not a Win32 LOB app"
                }
            } else {
                $LobApps = $AllApps | Where-Object { $_.AppType -eq '#microsoft.graph.win32LobApp' }
            }
            
            if (-not $LobApps -or $LobApps.Count -eq 0) {
                Write-IntuneInventoryLog -Message "No Win32 LOB apps found to download" -Level Warning
                return @{
                    Success = $true
                    ItemsProcessed = 0
                    Message = "No Win32 LOB apps found"
                }
            }
            
            Write-IntuneInventoryLog -Message "Found $($LobApps.Count) Win32 LOB apps to process" -Level Info
            
            foreach ($App in $LobApps) {
                try {
                    Write-IntuneInventoryLog -Message "Processing app: $($App.DisplayName) (ID: $($App.Id))" -Level Verbose
                    
                    # Create app-specific directory
                    $AppDir = Join-Path $DownloadPath $App.Id
                    if (-not (Test-Path $AppDir)) {
                        New-Item -Path $AppDir -ItemType Directory -Force | Out-Null
                    }
                    
                    # Check if already downloaded
                    $ExistingFiles = Get-ChildItem $AppDir -Filter "*.intunewin" -ErrorAction SilentlyContinue
                    if ($ExistingFiles -and -not $Force) {
                        Write-IntuneInventoryLog -Message "App content already downloaded for '$($App.DisplayName)'. Use -Force to re-download." -Level Info
                        continue
                    }
                    
                    # Get content versions from Graph API
                    Write-IntuneInventoryLog -Message "Retrieving content versions for '$($App.DisplayName)'" -Level Verbose
                    $ContentVersionsUri = "beta/deviceAppManagement/mobileApps/$($App.Id)/contentVersions"
                    $ContentVersions = Get-GraphRequestAll -Uri $ContentVersionsUri
                    
                    if (-not $ContentVersions -or $ContentVersions.Count -eq 0) {
                        Write-IntuneInventoryLog -Message "No content versions found for '$($App.DisplayName)'" -Level Warning
                        $ErrorCount++
                        $ErrorMessages += "No content versions for $($App.DisplayName)"
                        continue
                    }
                    
                    # Get latest version
                    $LatestVersion = $ContentVersions | Sort-Object { [datetime]$_.createdDateTime } -Descending | Select-Object -First 1
                    Write-IntuneInventoryLog -Message "Using content version $($LatestVersion.id) for '$($App.DisplayName)'" -Level Verbose
                    
                    # Get content files
                    $ContentFilesUri = "beta/deviceAppManagement/mobileApps/$($App.Id)/contentVersions/$($LatestVersion.id)/files"
                    $ContentFiles = Get-GraphRequestAll -Uri $ContentFilesUri
                    
                    if (-not $ContentFiles -or $ContentFiles.Count -eq 0) {
                        Write-IntuneInventoryLog -Message "No content files found for '$($App.DisplayName)'" -Level Warning
                        $ErrorCount++
                        $ErrorMessages += "No content files for $($App.DisplayName)"
                        continue
                    }
                    
                    # Download each content file
                    foreach ($ContentFile in $ContentFiles) {
                        try {
                            Write-IntuneInventoryLog -Message "Downloading file: $($ContentFile.name) for '$($App.DisplayName)'" -Level Verbose
                            
                            # Get download URL
                            $DownloadUrlUri = "beta/deviceAppManagement/mobileApps/$($App.Id)/contentVersions/$($LatestVersion.id)/files/$($ContentFile.id)"
                            $FileDetails = Invoke-GraphRequest -Uri $DownloadUrlUri -Method GET
                            
                            if (-not $FileDetails.azureStorageUri) {
                                Write-IntuneInventoryLog -Message "No download URL available for file: $($ContentFile.name)" -Level Warning
                                continue
                            }
                            
                            # Download the file
                            $LocalFilePath = Join-Path $AppDir $ContentFile.name
                            $DownloadUri = $FileDetails.azureStorageUri
                            
                            Write-IntuneInventoryLog -Message "Downloading to: $LocalFilePath" -Level Verbose
                            Invoke-WebRequest -Uri $DownloadUri -OutFile $LocalFilePath -ErrorAction Stop
                            
                            # Verify file size if available
                            if ($ContentFile.size -and (Test-Path $LocalFilePath)) {
                                $ActualSize = (Get-Item $LocalFilePath).Length
                                if ($ActualSize -eq $ContentFile.size) {
                                    Write-IntuneInventoryLog -Message "Successfully downloaded: $($ContentFile.name) ($ActualSize bytes)" -Level Info
                                    $DownloadedFiles += @{
                                        AppName = $App.DisplayName
                                        AppId = $App.Id
                                        FileName = $ContentFile.name
                                        FilePath = $LocalFilePath
                                        FileSize = $ActualSize
                                    }
                                } else {
                                    Write-IntuneInventoryLog -Message "File size mismatch for $($ContentFile.name). Expected: $($ContentFile.size), Actual: $ActualSize" -Level Warning
                                }
                            }
                        }
                        catch {
                            Write-IntuneInventoryLog -Message "Failed to download file '$($ContentFile.name)' for '$($App.DisplayName)': $($_.Exception.Message)" -Level Error
                            $ErrorCount++
                            $ErrorMessages += "Download failed for $($ContentFile.name) in $($App.DisplayName)"
                        }
                    }
                    
                    # Save app metadata
                    $AppMetadata = @{
                        AppInfo = $App
                        ContentVersion = $LatestVersion
                        ContentFiles = $ContentFiles
                        DownloadTimestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    }
                    
                    $MetadataPath = Join-Path $AppDir "app-metadata.json"
                    $AppMetadata | ConvertTo-Json -Depth 10 | Set-Content $MetadataPath
                    
                    $ItemsProcessed++
                    
                    if ($ItemsProcessed % 10 -eq 0) {
                        Write-IntuneInventoryLog -Message "Processed $ItemsProcessed apps so far..." -Level Info
                    }
                }
                catch {
                    $ErrorCount++
                    $ErrorMessage = "Failed to download content for app '$($App.DisplayName)': $($_.Exception.Message)"
                    $ErrorMessages += $ErrorMessage
                    Write-IntuneInventoryLog -Message $ErrorMessage -Level Error
                }
            }
            
            # Generate summary
            $EndTime = Get-Date
            $Duration = ($EndTime - $StartTime).TotalMinutes
            
            $SuccessMessage = "LOB app content download completed. Processed $ItemsProcessed apps, downloaded $($DownloadedFiles.Count) files"
            if ($ErrorCount -gt 0) {
                $SuccessMessage += " with $ErrorCount errors"
            }
            
            Write-IntuneInventoryLog -Message $SuccessMessage -Level Info
            Write-Host "`n=== DOWNLOAD SUMMARY ===" -ForegroundColor Green
            Write-Host "📱 Apps Processed: $ItemsProcessed" -ForegroundColor Cyan
            Write-Host "📁 Files Downloaded: $($DownloadedFiles.Count)" -ForegroundColor Cyan
            Write-Host "💾 Download Location: $DownloadPath" -ForegroundColor Cyan
            Write-Host "⏱️  Duration: $([math]::Round($Duration, 2)) minutes" -ForegroundColor Cyan
            
            if ($DownloadedFiles.Count -gt 0) {
                Write-Host "`n=== DOWNLOADED FILES ===" -ForegroundColor Yellow
                $DownloadedFiles | ForEach-Object {
                    Write-Host "  $($_.AppName): $($_.FileName) ($([math]::Round($_.FileSize / 1MB, 2)) MB)" -ForegroundColor White
                }
            }
            
            if ($ErrorCount -gt 0) {
                Write-Host "`n⚠️  ERRORS: $ErrorCount" -ForegroundColor Red
                $ErrorMessages | ForEach-Object { Write-Host "   - $_" -ForegroundColor Red }
            }
            
            return @{
                Success = $true
                ItemsProcessed = $ItemsProcessed
                FilesDownloaded = $DownloadedFiles.Count
                DownloadedFiles = $DownloadedFiles
                ErrorCount = $ErrorCount
                ErrorMessages = $ErrorMessages
                Duration = [math]::Round($Duration, 2)
                DownloadPath = $DownloadPath
                Message = $SuccessMessage
            }
        }
        catch {
            $ErrorMessage = "Critical error during LOB app content download: $($_.Exception.Message)"
            Write-IntuneInventoryLog -Message $ErrorMessage -Level Error
            throw $ErrorMessage
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "LOB app content download operation completed" -Level Info -Source "Get-IntuneLobAppContent"
    }
}
