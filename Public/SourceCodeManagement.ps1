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
    Attempts to retrieve LOB application content information from Intune.
    
    .DESCRIPTION
    This function attempts to access content metadata for Win32 LOB applications.
    
    IMPORTANT LIMITATIONS:
    - Microsoft Graph API does not provide direct access to download .intunewin file content
    - Content is encrypted and requires device-specific certificates for decryption  
    - This function can only retrieve content version metadata, not actual files
    
    ALTERNATIVE APPROACHES FOR CONTENT RECOVERY:
    - Use Intune Management Extension logs on enrolled devices
    - See: https://msendpointmgr.com/2019/01/18/how-to-decode-intune-win32-app-packages/
    - Tools: IntuneWinAppUtilDecoder and Get-DecryptInfoFromSideCarLogFiles.ps1
    
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
    
    .NOTES
    This function demonstrates the Graph API approach but is limited by Microsoft's security model.
    For actual content recovery, consider the device-based log parsing approach documented at:
    https://msendpointmgr.com/2019/01/18/how-to-decode-intune-win32-app-packages/
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

function Get-IntuneWin32AppRecoveryHelp {
    <#
    .SYNOPSIS
    Provides comprehensive guidance for recovering source files from Intune Win32 applications.
    
    .DESCRIPTION
    Since the Microsoft Graph API does not allow direct download of .intunewin content,
    this function provides information about multiple alternative approaches for recovering 
    application source files from Intune-managed devices.
    
    REQUIREMENTS:
    - PowerShell 5.1 compatible tools only
    - Standard character encoding (no special ASCII/Unicode characters)
    - Enrolled device with target applications installed
    
    .PARAMETER ShowTools
    Display information about available tools for content recovery.
    
    .PARAMETER ShowUrls
    Display relevant URLs and resources.
    
    .PARAMETER ShowMethods
    Display detailed information about all available recovery methods.
    
    .PARAMETER CheckEnvironment
    Check if the current environment supports the various recovery methods.
    
    .EXAMPLE
    Get-IntuneWin32AppRecoveryHelp
    
    .EXAMPLE  
    Get-IntuneWin32AppRecoveryHelp -ShowTools -ShowUrls -ShowMethods
    
    .EXAMPLE
    Get-IntuneWin32AppRecoveryHelp -CheckEnvironment
    
    .NOTES
    This guidance combines research from Oliver Kieselbach, Bilal el Haddouchi, and the MSEndpointMgr community.
    #>
    [CmdletBinding()]
    param(
        [switch]$ShowTools,
        [switch]$ShowUrls,
        [switch]$ShowMethods,
        [switch]$CheckEnvironment
    )
    
    Write-Host "`n=== INTUNE WIN32 APP CONTENT RECOVERY GUIDANCE ===" -ForegroundColor Cyan
    Write-Host "`nThe Microsoft Graph API does not provide direct access to download .intunewin file content" -ForegroundColor Yellow
    Write-Host "due to security restrictions. Content is encrypted and requires device-specific certificates." -ForegroundColor Yellow
    
    if ($CheckEnvironment) {
        Write-Host "`n=== ENVIRONMENT CHECK ===" -ForegroundColor Magenta
        
        # Check if running on enrolled device
        $IntuneRegPath = "HKLM:\SOFTWARE\Microsoft\Enrollments"
        $IsEnrolled = Test-Path $IntuneRegPath
        Write-Host "   Intune Enrollment: " -NoNewline -ForegroundColor White
        if ($IsEnrolled) {
            Write-Host "DETECTED" -ForegroundColor Green
        } else {
            Write-Host "NOT DETECTED" -ForegroundColor Red
        }
        
        # Check for IME installation
        $IMEPath = "C:\Program Files (x86)\Microsoft Intune Management Extension"
        $HasIME = Test-Path $IMEPath
        Write-Host "   Intune Management Extension: " -NoNewline -ForegroundColor White
        if ($HasIME) {
            Write-Host "INSTALLED" -ForegroundColor Green
        } else {
            Write-Host "NOT INSTALLED" -ForegroundColor Red
        }
        
        # Check for log files
        $LogPath = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log"
        $HasLogs = Test-Path $LogPath
        Write-Host "   IME Log Files: " -NoNewline -ForegroundColor White
        if ($HasLogs) {
            Write-Host "AVAILABLE" -ForegroundColor Green
            $LogSize = (Get-Item $LogPath -ErrorAction SilentlyContinue).Length / 1MB
            Write-Host "     Log Size: $([math]::Round($LogSize, 2)) MB" -ForegroundColor Cyan
        } else {
            Write-Host "NOT AVAILABLE" -ForegroundColor Red
        }
        
        # Check staging folder
        $StagingPath = "C:\Program Files (x86)\Microsoft Intune Management Extension\Content\Staging"
        $HasStaging = Test-Path $StagingPath
        Write-Host "   Staging Folder: " -NoNewline -ForegroundColor White
        if ($HasStaging) {
            Write-Host "EXISTS" -ForegroundColor Green
            $StagingItems = Get-ChildItem $StagingPath -ErrorAction SilentlyContinue
            Write-Host "     Current Items: $($StagingItems.Count)" -ForegroundColor Cyan
        } else {
            Write-Host "NOT FOUND" -ForegroundColor Red
        }
        
        # Check PowerShell version
        Write-Host "   PowerShell Version: " -NoNewline -ForegroundColor White
        if ($PSVersionTable.PSVersion.Major -eq 5) {
            Write-Host "$($PSVersionTable.PSVersion) (COMPATIBLE)" -ForegroundColor Green
        } else {
            Write-Host "$($PSVersionTable.PSVersion) (USE POWERSHELL 5.1)" -ForegroundColor Yellow
        }
        
        # Check admin privileges
        $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
        Write-Host "   Administrator Rights: " -NoNewline -ForegroundColor White
        if ($IsAdmin) {
            Write-Host "AVAILABLE" -ForegroundColor Green
        } else {
            Write-Host "REQUIRED FOR SOME METHODS" -ForegroundColor Yellow
        }
    }
    
    if ($ShowMethods) {
        Write-Host "`n=== AVAILABLE RECOVERY METHODS ===" -ForegroundColor Magenta
        
        Write-Host "`nMETHOD 1: Real-Time Staging Capture (Recommended)" -ForegroundColor Green
        Write-Host "   * Temporarily denies SYSTEM permissions on IMECache folder" -ForegroundColor White
        Write-Host "   * Captures apps during installation in staging folder" -ForegroundColor White
        Write-Host "   * Provides source files in ZIP format" -ForegroundColor White
        Write-Host "   * Requires: Admin rights, enrolled device, app reinstall" -ForegroundColor Cyan
        Write-Host "   * Success Rate: HIGH (if app can be reinstalled)" -ForegroundColor Green
        
        Write-Host "`nMETHOD 2: Log Parsing Approach (Historical)" -ForegroundColor Green
        Write-Host "   * Extracts download URLs and decryption keys from logs" -ForegroundColor White
        Write-Host "   * Downloads encrypted .bin files from Azure Storage" -ForegroundColor White
        Write-Host "   * Requires decryption with IntuneWinAppUtilDecoder" -ForegroundColor White
        Write-Host "   * Requires: Recent app installations in logs" -ForegroundColor Cyan
        Write-Host "   * Success Rate: MEDIUM (depends on log retention)" -ForegroundColor Yellow
        
        Write-Host "`nMETHOD 3: Certificate-Based Decryption" -ForegroundColor Green
        Write-Host "   * Uses device certificates for content decryption" -ForegroundColor White
        Write-Host "   * Advanced method requiring certificate extraction" -ForegroundColor White
        Write-Host "   * Combines with downloaded encrypted content" -ForegroundColor White
        Write-Host "   * Requires: Deep technical knowledge, certificate access" -ForegroundColor Cyan
        Write-Host "   * Success Rate: LOW (complex implementation)" -ForegroundColor Red
    }
    
    Write-Host "`nPRIMARY APPROACH: Real-Time Staging Capture" -ForegroundColor Green
    Write-Host "   1. Run staging capture script as Administrator" -ForegroundColor White
    Write-Host "   2. Temporarily deny SYSTEM permissions on IMECache" -ForegroundColor White
    Write-Host "   3. Install target app through Company Portal" -ForegroundColor White
    Write-Host "   4. App will show 'Failed to install' but files are captured" -ForegroundColor White
    Write-Host "   5. Copy ZIP files from staging folder" -ForegroundColor White
    Write-Host "   6. Restore SYSTEM permissions" -ForegroundColor White
    Write-Host "   7. Retry app installation if needed" -ForegroundColor White
    
    Write-Host "`nALTERNATIVE APPROACH: Log Parsing Method" -ForegroundColor Green
    Write-Host "   1. Use an enrolled device that has the apps installed" -ForegroundColor White
    Write-Host "   2. Parse Intune Management Extension logs for download URLs and decryption info" -ForegroundColor White
    Write-Host "   3. Use specialized tools to decrypt the content" -ForegroundColor White
    
    Write-Host "`nCRITICAL PATHS:" -ForegroundColor Green
    Write-Host "   IME Logs: C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\" -ForegroundColor Cyan
    Write-Host "   Staging: C:\Program Files (x86)\Microsoft Intune Management Extension\Content\Staging\" -ForegroundColor Cyan
    Write-Host "   IMECache: C:\Windows\IMECache\" -ForegroundColor Cyan
    
    if ($ShowTools) {
        Write-Host "`nRECOMMENDED TOOLS:" -ForegroundColor Green
        Write-Host "   STAGING CAPTURE METHOD:" -ForegroundColor Yellow
        Write-Host "   * BackupWin32Apps.ps1 - Real-time staging capture script" -ForegroundColor White
        Write-Host "   * 7-Zip or WinRAR - For extracting captured ZIP files" -ForegroundColor White
        Write-Host "   * PowerShell 5.1 - For running capture scripts" -ForegroundColor White
        
        Write-Host "`n   LOG PARSING METHOD:" -ForegroundColor Yellow
        Write-Host "   * Get-DecryptInfoFromSideCarLogFiles.ps1 - Extracts info from logs" -ForegroundColor White
        Write-Host "   * IntuneWinAppUtilDecoder.exe - Decrypts .intunewin files" -ForegroundColor White
        Write-Host "   * PowerShell 5.1 - For log parsing and automation" -ForegroundColor White
        
        Write-Host "`n   CERTIFICATE METHOD:" -ForegroundColor Yellow
        Write-Host "   * Certificate extraction tools (advanced)" -ForegroundColor White
        Write-Host "   * Custom decryption scripts" -ForegroundColor White
        Write-Host "   * Cryptographic libraries" -ForegroundColor White
    }
    
    if ($ShowUrls) {
        Write-Host "`nUSEFUL RESOURCES:" -ForegroundColor Green
        Write-Host "   STAGING CAPTURE RESOURCES:" -ForegroundColor Yellow
        Write-Host "   * Blog: https://www.bilalelhaddouchi.nl/index.php/2022/03/23/extract-win32-apps/" -ForegroundColor Cyan
        Write-Host "   * Script: https://github.com/Mr-Famous/ExtractWin32" -ForegroundColor Cyan
        
        Write-Host "`n   LOG PARSING RESOURCES:" -ForegroundColor Yellow
        Write-Host "   * Blog: https://msendpointmgr.com/2019/01/18/how-to-decode-intune-win32-app-packages/" -ForegroundColor Cyan
        Write-Host "   * GitHub: https://github.com/okieselbach/Intune" -ForegroundColor Cyan
        Write-Host "   * Script: https://github.com/okieselbach/Intune/blob/master/Get-DecryptInfoFromSideCarLogFiles.ps1" -ForegroundColor Cyan
        Write-Host "   * Decoder: https://github.com/okieselbach/Intune/tree/master/IntuneWinAppUtilDecoder" -ForegroundColor Cyan
        
        Write-Host "`n   GENERAL RESOURCES:" -ForegroundColor Yellow
        Write-Host "   * MSEndpointMgr Community: https://msendpointmgr.com/" -ForegroundColor Cyan
        Write-Host "   * Oliver Kieselbach Blog: https://oliverkieselbach.com/" -ForegroundColor Cyan
    }
    
    Write-Host "`nIMPORTANT NOTES:" -ForegroundColor Yellow
    Write-Host "   * These approaches require an enrolled device with apps installed/installable" -ForegroundColor White
    Write-Host "   * The device must have necessary certificates for decryption" -ForegroundColor White
    Write-Host "   * Staging method requires app reinstallation" -ForegroundColor White
    Write-Host "   * Log parsing depends on recent app installation history" -ForegroundColor White
    Write-Host "   * These methods are for legitimate recovery purposes only" -ForegroundColor White
    Write-Host "   * Tools must be compatible with PowerShell 5.1" -ForegroundColor White
    Write-Host "   * Use standard characters only - NO special ASCII/Unicode characters" -ForegroundColor White
    Write-Host "   * Always test in non-production environment first" -ForegroundColor White
    Write-Host "   * Some methods may temporarily affect app functionality" -ForegroundColor White
    
    Write-Host "`nRECOMMENDATION:" -ForegroundColor Magenta
    Write-Host "   Start with Method 1 (Staging Capture) for best results." -ForegroundColor White
    Write-Host "   Use Method 2 (Log Parsing) if apps were recently installed." -ForegroundColor White
    Write-Host "   Consider Method 3 (Certificate) only for advanced scenarios." -ForegroundColor White
    
    Write-Host "`nTIP: Run with -ShowTools, -ShowUrls, -ShowMethods, and -CheckEnvironment for complete information`n" -ForegroundColor Magenta
}

function Test-IntuneAppExtractionEnvironment {
    <#
    .SYNOPSIS
    Tests if the current environment supports Intune app extraction methods.
    
    .DESCRIPTION
    Performs comprehensive checks to determine which app extraction methods
    are available and provides recommendations based on the environment.
    
    .PARAMETER Detailed
    Show detailed information about each check.
    
    .EXAMPLE
    Test-IntuneAppExtractionEnvironment
    
    .EXAMPLE
    Test-IntuneAppExtractionEnvironment -Detailed
    
    .NOTES
    PowerShell 5.1 compatible function for environment assessment.
    #>
    [CmdletBinding()]
    param(
        [switch]$Detailed
    )
    
    Write-Host "`n=== INTUNE APP EXTRACTION ENVIRONMENT TEST ===" -ForegroundColor Cyan
    
    $Results = @{
        OverallCompatible = $true
        StagingMethod = $false
        LogParsingMethod = $false
        Recommendations = @()
        Issues = @()
    }
    
    # Test 1: PowerShell Version
    Write-Host "`nTesting PowerShell compatibility..." -ForegroundColor Yellow
    if ($PSVersionTable.PSVersion.Major -eq 5) {
        Write-Host "   PowerShell 5.1: COMPATIBLE" -ForegroundColor Green
        if ($Detailed) {
            Write-Host "     Version: $($PSVersionTable.PSVersion)" -ForegroundColor Cyan
            Write-Host "     Edition: $($PSVersionTable.PSEdition)" -ForegroundColor Cyan
        }
    } else {
        Write-Host "   PowerShell Version: INCOMPATIBLE (Use PowerShell 5.1)" -ForegroundColor Red
        $Results.Issues += "PowerShell 5.1 required for extraction tools"
        $Results.OverallCompatible = $false
    }
    
    # Test 2: Administrative Rights
    Write-Host "`nTesting administrative privileges..." -ForegroundColor Yellow
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if ($IsAdmin) {
        Write-Host "   Administrator Rights: AVAILABLE" -ForegroundColor Green
        $Results.StagingMethod = $true
    } else {
        Write-Host "   Administrator Rights: NOT AVAILABLE" -ForegroundColor Red
        $Results.Issues += "Administrator rights required for staging capture method"
        Write-Host "     Note: Required for staging capture method" -ForegroundColor Yellow
    }
    
    # Test 3: Intune Enrollment
    Write-Host "`nTesting Intune enrollment..." -ForegroundColor Yellow
    try {
        $IntuneRegPath = "HKLM:\SOFTWARE\Microsoft\Enrollments"
        $Enrollments = Get-ChildItem $IntuneRegPath -ErrorAction SilentlyContinue | Where-Object {
            $Props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            $Props.ProviderID -eq "MS DM Server"
        }
        
        if ($Enrollments) {
            Write-Host "   Intune Enrollment: DETECTED" -ForegroundColor Green
            if ($Detailed) {
                Write-Host "     Enrolled Device Count: $($Enrollments.Count)" -ForegroundColor Cyan
            }
        } else {
            Write-Host "   Intune Enrollment: NOT DETECTED" -ForegroundColor Red
            $Results.Issues += "Device must be enrolled in Intune"
            $Results.OverallCompatible = $false
        }
    } catch {
        Write-Host "   Intune Enrollment: UNABLE TO VERIFY" -ForegroundColor Yellow
        $Results.Issues += "Unable to verify Intune enrollment status"
    }
    
    # Test 4: Intune Management Extension
    Write-Host "`nTesting Intune Management Extension..." -ForegroundColor Yellow
    $IMEPath = "C:\Program Files (x86)\Microsoft Intune Management Extension"
    $IMEService = Get-Service -Name "IntuneManagementExtension" -ErrorAction SilentlyContinue
    
    if (Test-Path $IMEPath) {
        Write-Host "   IME Installation: FOUND" -ForegroundColor Green
        if ($Detailed) {
            Write-Host "     Path: $IMEPath" -ForegroundColor Cyan
        }
    } else {
        Write-Host "   IME Installation: NOT FOUND" -ForegroundColor Red
        $Results.Issues += "Intune Management Extension not installed"
        $Results.OverallCompatible = $false
    }
    
    if ($IMEService) {
        Write-Host "   IME Service: $($IMEService.Status)" -ForegroundColor Green
        if ($Detailed) {
            Write-Host "     Service Name: $($IMEService.Name)" -ForegroundColor Cyan
            Write-Host "     Display Name: $($IMEService.DisplayName)" -ForegroundColor Cyan
        }
    } else {
        Write-Host "   IME Service: NOT FOUND" -ForegroundColor Red
        $Results.Issues += "Intune Management Extension service not found"
    }
    
    # Test 5: Log Files
    Write-Host "`nTesting log file availability..." -ForegroundColor Yellow
    $LogPath = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log"
    if (Test-Path $LogPath) {
        $LogFile = Get-Item $LogPath
        $LogSizeMB = [math]::Round($LogFile.Length / 1MB, 2)
        $LogAge = (Get-Date) - $LogFile.LastWriteTime
        
        Write-Host "   Log Files: AVAILABLE" -ForegroundColor Green
        if ($Detailed) {
            Write-Host "     Size: $LogSizeMB MB" -ForegroundColor Cyan
            Write-Host "     Last Modified: $($LogFile.LastWriteTime)" -ForegroundColor Cyan
            Write-Host "     Age: $([math]::Round($LogAge.TotalHours, 1)) hours" -ForegroundColor Cyan
        }
        
        if ($LogAge.TotalDays -lt 7) {
            $Results.LogParsingMethod = $true
            Write-Host "     Recent Activity: GOOD FOR LOG PARSING" -ForegroundColor Green
        } else {
            Write-Host "     Recent Activity: OLD (may have limited data)" -ForegroundColor Yellow
            $Results.Recommendations += "Install some Win32 apps to generate fresh log data"
        }
    } else {
        Write-Host "   Log Files: NOT FOUND" -ForegroundColor Red
        $Results.Issues += "IME log files not found"
    }
    
    # Test 6: Staging Folder
    Write-Host "`nTesting staging folder access..." -ForegroundColor Yellow
    $StagingPath = "C:\Program Files (x86)\Microsoft Intune Management Extension\Content\Staging"
    if (Test-Path $StagingPath) {
        Write-Host "   Staging Folder: EXISTS" -ForegroundColor Green
        
        try {
            $StagingItems = Get-ChildItem $StagingPath -ErrorAction SilentlyContinue
            if ($Detailed) {
                Write-Host "     Path: $StagingPath" -ForegroundColor Cyan
                Write-Host "     Current Items: $($StagingItems.Count)" -ForegroundColor Cyan
            }
            
            # Test write access (required for staging method)
            $TestFile = Join-Path $StagingPath "test_write_access.tmp"
            try {
                "test" | Out-File $TestFile -ErrorAction Stop
                Remove-Item $TestFile -ErrorAction SilentlyContinue
                Write-Host "     Write Access: AVAILABLE" -ForegroundColor Green
            } catch {
                Write-Host "     Write Access: RESTRICTED" -ForegroundColor Yellow
                $Results.Recommendations += "May need elevated permissions for staging folder access"
            }
        } catch {
            Write-Host "     Access: RESTRICTED" -ForegroundColor Yellow
        }
    } else {
        Write-Host "   Staging Folder: NOT FOUND" -ForegroundColor Red
        $Results.Issues += "Staging folder not accessible"
    }
    
    # Test 7: Required Tools Availability
    Write-Host "`nTesting tool availability..." -ForegroundColor Yellow
    
    # Check for 7-Zip
    $SevenZipPaths = @(
        "C:\Program Files\7-Zip\7z.exe",
        "C:\Program Files (x86)\7-Zip\7z.exe"
    )
    $SevenZipFound = $SevenZipPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    
    if ($SevenZipFound) {
        Write-Host "   7-Zip: FOUND" -ForegroundColor Green
        if ($Detailed) {
            Write-Host "     Path: $SevenZipFound" -ForegroundColor Cyan
        }
    } else {
        Write-Host "   7-Zip: NOT FOUND" -ForegroundColor Yellow
        $Results.Recommendations += "Install 7-Zip for file extraction"
    }
    
    # Generate Summary
    Write-Host "`n=== ASSESSMENT SUMMARY ===" -ForegroundColor Magenta
    
    if ($Results.OverallCompatible) {
        Write-Host "Overall Environment: COMPATIBLE" -ForegroundColor Green
    } else {
        Write-Host "Overall Environment: NEEDS ATTENTION" -ForegroundColor Red
    }
    
    Write-Host "`nAvailable Methods:" -ForegroundColor White
    if ($Results.StagingMethod) {
        Write-Host "   [+] Staging Capture Method: READY" -ForegroundColor Green
    } else {
        Write-Host "   [-] Staging Capture Method: NOT READY" -ForegroundColor Red
    }
    
    if ($Results.LogParsingMethod) {
        Write-Host "   [+] Log Parsing Method: READY" -ForegroundColor Green
    } else {
        Write-Host "   [-] Log Parsing Method: LIMITED/NOT READY" -ForegroundColor Yellow
    }
    
    if ($Results.Issues.Count -gt 0) {
        Write-Host "`nIssues to Address:" -ForegroundColor Red
        $Results.Issues | ForEach-Object { Write-Host "   * $_" -ForegroundColor Red }
    }
    
    if ($Results.Recommendations.Count -gt 0) {
        Write-Host "`nRecommendations:" -ForegroundColor Yellow
        $Results.Recommendations | ForEach-Object { Write-Host "   * $_" -ForegroundColor Yellow }
    }
    
    Write-Host "`nNext Steps:" -ForegroundColor Cyan
    if ($Results.StagingMethod) {
        Write-Host "   1. Use 'Invoke-IntuneStagingCapture' to capture apps during installation" -ForegroundColor White
    }
    if ($Results.LogParsingMethod) {
        Write-Host "   2. Use 'Get-IntuneLoggedApps' to find apps in logs" -ForegroundColor White
    }
    Write-Host "   3. Run 'Get-IntuneWin32AppRecoveryHelp -ShowMethods' for detailed guidance" -ForegroundColor White
    
    return $Results
}

function Get-IntuneLoggedApps {
    <#
    .SYNOPSIS
    Scans Intune Management Extension logs for Win32 app installation records.
    
    .DESCRIPTION
    Parses the IME log files to find Win32 applications that have been installed
    and extracts the download URLs and decryption information when available.
    
    REQUIREMENTS:
    - PowerShell 5.1 compatible
    - Access to IME log files
    - Recent app installation activity
    
    .PARAMETER LogPath
    Path to the IME log file. Defaults to standard location.
    
    .PARAMETER DaysBack
    Number of days back to search in logs. Default is 30.
    
    .PARAMETER AppName
    Filter results to specific app name pattern.
    
    .PARAMETER ShowUrls
    Include download URLs in the output (if found).
    
    .PARAMETER ShowKeys
    Include decryption keys in the output (if found).
    
    .EXAMPLE
    Get-IntuneLoggedApps
    
    .EXAMPLE
    Get-IntuneLoggedApps -DaysBack 7 -ShowUrls -ShowKeys
    
    .EXAMPLE
    Get-IntuneLoggedApps -AppName "*Chrome*"
    
    .NOTES
    This function implements the log parsing method for app recovery.
    Requires recent app installation activity to find useful data.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$LogPath = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log",
        
        [Parameter()]
        [int]$DaysBack = 30,
        
        [Parameter()]
        [string]$AppName = "*",
        
        [Parameter()]
        [switch]$ShowUrls,
        
        [Parameter()]
        [switch]$ShowKeys
    )
    
    Write-Host "`n=== INTUNE LOGGED APPS SCANNER ===" -ForegroundColor Cyan
    
    if (-not (Test-Path $LogPath)) {
        Write-Host "Log file not found: $LogPath" -ForegroundColor Red
        Write-Host "Ensure Intune Management Extension is installed and has recent activity." -ForegroundColor Yellow
        return
    }
    
    Write-Host "Scanning log file: $LogPath" -ForegroundColor White
    
    try {
        $LogContent = Get-Content $LogPath -ErrorAction Stop
        $CutoffDate = (Get-Date).AddDays(-$DaysBack)
        
        Write-Host "Searching $($LogContent.Count) log lines for app installation records..." -ForegroundColor Yellow
        Write-Host "Date range: $($CutoffDate.ToString('yyyy-MM-dd')) to $(Get-Date -Format 'yyyy-MM-dd')" -ForegroundColor Cyan
        
        $Apps = @()
        $CurrentApp = $null
        
        foreach ($Line in $LogContent) {
            # Look for app download start patterns
            if ($Line -match "Download.*content.*for application.*name=(.+?),.*id=(.+?),") {
                $AppDisplayName = $Matches[1].Trim()
                $AppId = $Matches[2].Trim()
                
                if ($AppDisplayName -like $AppName) {
                    $CurrentApp = @{
                        DisplayName = $AppDisplayName
                        Id = $AppId
                        DownloadUrl = $null
                        EncryptionKey = $null
                        IV = $null
                        LogEntries = @()
                        FoundDate = $null
                    }
                    
                    # Try to extract date from log line
                    if ($Line -match "\[(.+?)\]") {
                        try {
                            $LogDate = [DateTime]::Parse($Matches[1])
                            $CurrentApp.FoundDate = $LogDate
                            
                            if ($LogDate -lt $CutoffDate) {
                                $CurrentApp = $null
                                continue
                            }
                        } catch {
                            # Continue if date parsing fails
                        }
                    }
                }
            }
            
            # Look for download URLs
            if ($CurrentApp -and $Line -match "https://[^\""\s]+\.bin") {
                $CurrentApp.DownloadUrl = $Matches[0]
                $CurrentApp.LogEntries += "URL: $Line"
            }
            
            # Look for encryption keys
            if ($CurrentApp -and $Line -match "encryptionKey.*?([A-Za-z0-9+/=]{20,})") {
                $CurrentApp.EncryptionKey = $Matches[1]
                $CurrentApp.LogEntries += "KEY: $Line"
            }
            
            # Look for IV values
            if ($CurrentApp -and $Line -match "iv.*?([A-Za-z0-9+/=]{10,})") {
                $CurrentApp.IV = $Matches[1]
                $CurrentApp.LogEntries += "IV: $Line"
            }
            
            # If we have complete info for current app, save it
            if ($CurrentApp -and $CurrentApp.DownloadUrl -and $CurrentApp.EncryptionKey) {
                $Apps += $CurrentApp
                $CurrentApp = $null
            }
        }
        
        # Add any remaining incomplete app
        if ($CurrentApp) {
            $Apps += $CurrentApp
        }
        
        Write-Host "`nFound $($Apps.Count) app installation records" -ForegroundColor Green
        
        if ($Apps.Count -eq 0) {
            Write-Host "`nNo app installation records found in the specified time period." -ForegroundColor Yellow
            Write-Host "Try increasing -DaysBack parameter or ensure apps have been installed recently." -ForegroundColor Yellow
            return
        }
        
        # Display results
        Write-Host "`n=== DISCOVERED APPLICATIONS ===" -ForegroundColor Magenta
        
        for ($i = 0; $i -lt $Apps.Count; $i++) {
            $App = $Apps[$i]
            Write-Host "`n[$($i + 1)] $($App.DisplayName)" -ForegroundColor Green
            Write-Host "    App ID: $($App.Id)" -ForegroundColor Cyan
            
            if ($App.FoundDate) {
                Write-Host "    Log Date: $($App.FoundDate.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
            }
            
            $HasCompleteInfo = $App.DownloadUrl -and $App.EncryptionKey
            Write-Host "    Recovery Data: " -NoNewline -ForegroundColor White
            if ($HasCompleteInfo) {
                Write-Host "COMPLETE (URL + Key)" -ForegroundColor Green
            } elseif ($App.DownloadUrl) {
                Write-Host "PARTIAL (URL only)" -ForegroundColor Yellow
            } else {
                Write-Host "INCOMPLETE" -ForegroundColor Red
            }
            
            if ($ShowUrls -and $App.DownloadUrl) {
                Write-Host "    Download URL: $($App.DownloadUrl)" -ForegroundColor Yellow
            }
            
            if ($ShowKeys) {
                if ($App.EncryptionKey) {
                    Write-Host "    Encryption Key: $($App.EncryptionKey)" -ForegroundColor Yellow
                }
                if ($App.IV) {
                    Write-Host "    IV: $($App.IV)" -ForegroundColor Yellow
                }
            }
        }
        
        # Show usage instructions
        Write-Host "`n=== NEXT STEPS ===" -ForegroundColor Magenta
        
        $CompleteApps = $Apps | Where-Object { $_.DownloadUrl -and $_.EncryptionKey }
        if ($CompleteApps.Count -gt 0) {
            Write-Host "Apps with complete recovery data: $($CompleteApps.Count)" -ForegroundColor Green
            Write-Host "`nTo download and decrypt these apps:" -ForegroundColor White
            Write-Host "1. Download the .bin files from the URLs" -ForegroundColor Cyan
            Write-Host "2. Use IntuneWinAppUtilDecoder.exe with the encryption keys" -ForegroundColor Cyan
            Write-Host "3. Extract the decrypted .intunewin files with 7-Zip" -ForegroundColor Cyan
            
            Write-Host "`nExample command:" -ForegroundColor Yellow
            $ExampleApp = $CompleteApps[0]
            Write-Host "IntuneWinAppUtilDecoder.exe `"downloaded_file.bin`" /key:`"$($ExampleApp.EncryptionKey)`"" -ForegroundColor White
            if ($ExampleApp.IV) {
                Write-Host "   /iv:`"$($ExampleApp.IV)`"" -ForegroundColor White
            }
        }
        
        $PartialApps = $Apps | Where-Object { $_.DownloadUrl -and -not $_.EncryptionKey }
        if ($PartialApps.Count -gt 0) {
            Write-Host "`nApps with partial data (URL only): $($PartialApps.Count)" -ForegroundColor Yellow
            Write-Host "These may require additional log searching or alternative methods." -ForegroundColor Yellow
        }
        
        Write-Host "`nTip: Use -ShowUrls and -ShowKeys switches to see the actual recovery data." -ForegroundColor Cyan
        
        return $Apps
        
    } catch {
        Write-Host "Error reading log file: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

function Invoke-IntuneStagingCapture {
    <#
    .SYNOPSIS
    Implements the staging folder capture method for Intune Win32 apps.
    
    .DESCRIPTION
    This function temporarily modifies SYSTEM permissions to prevent cleanup
    of staging files during app installation, allowing capture of source files.
    
    WARNING: This method temporarily affects system security permissions.
    Use only in controlled environments for legitimate recovery purposes.
    
    REQUIREMENTS:
    - Administrator privileges
    - PowerShell 5.1
    - Enrolled Intune device
    - Apps available for installation
    
    .PARAMETER OutputPath
    Directory where captured files will be saved. Default: C:\IntuneAppBackup
    
    .PARAMETER WhatIf
    Show what would be done without making changes.
    
    .EXAMPLE
    Invoke-IntuneStagingCapture
    
    .EXAMPLE
    Invoke-IntuneStagingCapture -OutputPath "D:\AppBackups" -WhatIf
    
    .NOTES
    This implements the method described by Bilal el Haddouchi.
    Always test in non-production environments first.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string]$OutputPath = "C:\IntuneAppBackup",
        
        [Parameter()]
        [switch]$WhatIf
    )
    
    # Check prerequisites
    Write-Host "`n=== INTUNE STAGING CAPTURE ===" -ForegroundColor Cyan
    
    # Check admin rights
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if (-not $IsAdmin) {
        Write-Host "ERROR: Administrator privileges required for this operation." -ForegroundColor Red
        Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Yellow
        return
    }
    
    # Check paths
    $IMECachePath = "C:\Windows\IMECache"
    $StagingPath = "C:\Program Files (x86)\Microsoft Intune Management Extension\Content\Staging"
    
    if (-not (Test-Path $IMECachePath)) {
        Write-Host "ERROR: IMECache folder not found: $IMECachePath" -ForegroundColor Red
        return
    }
    
    if (-not (Test-Path $StagingPath)) {
        Write-Host "ERROR: Staging folder not found: $StagingPath" -ForegroundColor Red
        return
    }
    
    Write-Host "Prerequisites check: PASSED" -ForegroundColor Green
    Write-Host "IMECache Path: $IMECachePath" -ForegroundColor Cyan
    Write-Host "Staging Path: $StagingPath" -ForegroundColor Cyan
    Write-Host "Output Path: $OutputPath" -ForegroundColor Cyan
    
    if ($WhatIf) {
        Write-Host "`n=== WHAT IF MODE ===" -ForegroundColor Yellow
        Write-Host "Would perform the following operations:" -ForegroundColor White
        Write-Host "1. Create output directory: $OutputPath" -ForegroundColor White
        Write-Host "2. Deny SYSTEM permissions on: $IMECachePath" -ForegroundColor White
        Write-Host "3. Wait for user to install app via Company Portal" -ForegroundColor White
        Write-Host "4. Copy ZIP files from staging folder to output" -ForegroundColor White
        Write-Host "5. Restore SYSTEM permissions on: $IMECachePath" -ForegroundColor White
        Write-Host "`nNo actual changes would be made." -ForegroundColor Yellow
        return
    }
    
    # Warning prompt
    Write-Host "`n=== WARNING ===" -ForegroundColor Red
    Write-Host "This operation will temporarily modify SYSTEM permissions on the IMECache folder." -ForegroundColor Yellow
    Write-Host "During this time, app installations may fail until permissions are restored." -ForegroundColor Yellow
    Write-Host "Only proceed if you understand the implications." -ForegroundColor Yellow
    
    $UserConfirm = Read-Host "`nType 'YES' to continue or anything else to cancel"
    if ($UserConfirm -ne "YES") {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        return
    }
    
    try {
        # Create output directory
        if (-not (Test-Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
            Write-Host "`nCreated output directory: $OutputPath" -ForegroundColor Green
        }
        
        # Get SYSTEM SID and account
        $SystemSID = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-18")
        $SystemAccount = $SystemSID.Translate([System.Security.Principal.NTAccount])
        $SystemName = $SystemAccount.Value
        
        Write-Host "`nSTEP 1: Modifying SYSTEM permissions on IMECache..." -ForegroundColor Yellow
        
        # Get current ACL
        $OriginalACL = Get-ACL $IMECachePath
        $CurrentACL = Get-ACL $IMECachePath
        
        # Remove existing SYSTEM permissions
        $CurrentACL.Access | Where-Object { $_.IdentityReference -eq $SystemName } | ForEach-Object {
            $CurrentACL.RemoveAccessRule($_) | Out-Null
        }
        
        # Add DENY rule for SYSTEM
        $DenyRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $SystemName, 
            "FullControl", 
            "Deny"
        )
        $CurrentACL.SetAccessRule($DenyRule)
        
        # Apply the modified ACL
        Set-ACL $IMECachePath $CurrentACL
        Write-Host "   SYSTEM permissions denied on IMECache folder" -ForegroundColor Green
        
        Write-Host "`nSTEP 2: Ready for app installation" -ForegroundColor Yellow
        Write-Host "NOW:" -ForegroundColor Red
        Write-Host "1. Open Company Portal" -ForegroundColor White
        Write-Host "2. Install the target Win32 application" -ForegroundColor White
        Write-Host "3. The app will show 'Failed to install' status" -ForegroundColor White
        Write-Host "4. Wait for the failure, then press ENTER here to continue" -ForegroundColor White
        
        Read-Host "`nPress ENTER after the app installation shows as FAILED"
        
        Write-Host "`nSTEP 3: Capturing staging files..." -ForegroundColor Yellow
        
        # Find ZIP files in staging folder
        $ZipFiles = Get-ChildItem $StagingPath -Recurse -Filter "*.zip" -ErrorAction SilentlyContinue
        
        if ($ZipFiles.Count -eq 0) {
            Write-Host "   No ZIP files found in staging folder" -ForegroundColor Red
            Write-Host "   This may indicate the app hasn't been processed yet or method failed" -ForegroundColor Yellow
        } else {
            Write-Host "   Found $($ZipFiles.Count) ZIP file(s) in staging folder" -ForegroundColor Green
            
            foreach ($ZipFile in $ZipFiles) {
                $DestPath = Join-Path $OutputPath $ZipFile.Name
                Copy-Item $ZipFile.FullName $DestPath -Force
                Write-Host "   Copied: $($ZipFile.Name) -> $OutputPath" -ForegroundColor Cyan
            }
        }
        
        # Also capture any other interesting files
        $AllFiles = Get-ChildItem $StagingPath -Recurse -File -ErrorAction SilentlyContinue
        if ($AllFiles.Count -gt $ZipFiles.Count) {
            Write-Host "   Found additional files in staging:" -ForegroundColor Yellow
            $OtherFiles = $AllFiles | Where-Object { $_.Extension -ne ".zip" }
            foreach ($File in $OtherFiles) {
                $RelativePath = $File.FullName.Replace($StagingPath, "").TrimStart("\")
                $DestDir = Join-Path $OutputPath (Split-Path $RelativePath -Parent)
                if (-not (Test-Path $DestDir)) {
                    New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
                }
                $DestPath = Join-Path $OutputPath $RelativePath
                Copy-Item $File.FullName $DestPath -Force
                Write-Host "   Captured: $RelativePath" -ForegroundColor Cyan
            }
        }
        
        Write-Host "`nSTEP 4: Restoring SYSTEM permissions..." -ForegroundColor Yellow
        
        # Restore original permissions
        Set-ACL $IMECachePath $OriginalACL
        Write-Host "   SYSTEM permissions restored on IMECache folder" -ForegroundColor Green
        
        Write-Host "`n=== CAPTURE COMPLETE ===" -ForegroundColor Green
        Write-Host "Captured files location: $OutputPath" -ForegroundColor Cyan
        
        # Show results
        $CapturedFiles = Get-ChildItem $OutputPath -Recurse -File
        if ($CapturedFiles.Count -gt 0) {
            Write-Host "`nCaptured files:" -ForegroundColor White
            $CapturedFiles | ForEach-Object {
                $SizeMB = [math]::Round($_.Length / 1MB, 2)
                Write-Host "   $($_.Name) ($SizeMB MB)" -ForegroundColor Cyan
            }
            
            Write-Host "`nNext steps:" -ForegroundColor Yellow
            Write-Host "1. Extract ZIP files using 7-Zip or similar tool" -ForegroundColor White
            Write-Host "2. You can now retry the app installation in Company Portal" -ForegroundColor White
            Write-Host "3. The app should install normally now that permissions are restored" -ForegroundColor White
        } else {
            Write-Host "`nNo files were captured. This may indicate:" -ForegroundColor Yellow
            Write-Host "- The app installation didn't reach the staging phase" -ForegroundColor White
            Write-Host "- The app was already installed" -ForegroundColor White
            Write-Host "- The timing of the permission change was incorrect" -ForegroundColor White
            Write-Host "Try uninstalling the app first, then repeat the process" -ForegroundColor White
        }
        
    } catch {
        Write-Host "`nERROR during staging capture: $($_.Exception.Message)" -ForegroundColor Red
        
        # Always try to restore permissions on error
        try {
            Write-Host "Attempting to restore SYSTEM permissions..." -ForegroundColor Yellow
            Set-ACL $IMECachePath $OriginalACL
            Write-Host "SYSTEM permissions restored" -ForegroundColor Green
        } catch {
            Write-Host "CRITICAL: Failed to restore SYSTEM permissions!" -ForegroundColor Red
            Write-Host "Manual intervention may be required to restore IMECache permissions" -ForegroundColor Red
        }
        
        throw
    }
}
