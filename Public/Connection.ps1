function Initialize-IntuneInventoryStorage {
    <#
    .SYNOPSIS
    Initializes the JSON storage system for Intune inventory data.
    
    .DESCRIPTION
    Creates the directory structure and storage files for JSON-based inventory data.
    This replaces the SQLite database with a more user-friendly JSON file system.
    
    .PARAMETER StoragePath
    The full path where the storage directory should be created.
    If not specified, defaults to a directory in the user's AppData folder.
    
    .PARAMETER Force
    If specified, reinitializes the storage even if it already exists.
    
    .EXAMPLE
    Initialize-IntuneInventoryStorage
    
    .EXAMPLE
    Initialize-IntuneInventoryStorage -StoragePath "C:\IntuneInventory"
    
    .EXAMPLE
    Initialize-IntuneInventoryStorage -Force
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string]$StoragePath = "$env:APPDATA\IntuneInventory",
        
        [Parameter()]
        [switch]$Force
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting JSON storage initialization" -Level Info -Source "Initialize-IntuneInventoryStorage"
    }
    
    process {
        try {
            # Check if storage exists and handle Force parameter
            if ((Test-Path -Path $StoragePath) -and -not $Force) {
                Write-IntuneInventoryLog -Message "Storage directory already exists at: $StoragePath" -Level Info
                
                # Test existing storage
                $TestResult = Initialize-JsonStorage -StoragePath $StoragePath
                if ($TestResult) {
                    Write-IntuneInventoryLog -Message "Existing storage is valid and accessible" -Level Info
                    
                    # Show current statistics
                    $Stats = Get-StorageStatistics
                    if ($Stats) {
                        Write-Host "`n[Storage Statistics]" -ForegroundColor Yellow
                        Write-Host "   Location: $($Stats.StorageRoot)" -ForegroundColor White
                        Write-Host "   Applications: $($Stats.Applications)" -ForegroundColor White
                        Write-Host "   Scripts: $($Stats.Scripts)" -ForegroundColor White
                        Write-Host "   Remediations: $($Stats.Remediations)" -ForegroundColor White
                        Write-Host "   Inventory Runs: $($Stats.InventoryRuns)" -ForegroundColor White
                    }
                    
                    return $StoragePath
                }
            }
            
            if ($Force -and (Test-Path -Path $StoragePath)) {
                Write-IntuneInventoryLog -Message "Force specified - removing existing storage" -Level Warning
                if ($PSCmdlet.ShouldProcess($StoragePath, "Remove existing storage directory")) {
                    Remove-Item -Path $StoragePath -Recurse -Force
                }
            }
            
            # Initialize new storage
            Write-IntuneInventoryLog -Message "Creating new JSON storage at: $StoragePath" -Level Info
            $InitResult = Initialize-JsonStorage -StoragePath $StoragePath
            
            if ($InitResult) {
                Write-Host "`n[JSON Storage Initialized Successfully!]" -ForegroundColor Green
                Write-Host "   Location: $StoragePath" -ForegroundColor Cyan
                Write-Host "   Type: JSON-based (user-friendly)" -ForegroundColor Cyan
                Write-Host "   Files: applications.json, scripts.json, remediations.json, etc." -ForegroundColor Cyan
                Write-Host "`n[Storage Structure]" -ForegroundColor Yellow
                Write-Host "   $StoragePath\" -ForegroundColor White
                Write-Host "   |-- metadata.json (storage information)" -ForegroundColor Gray
                Write-Host "   |-- applications.json (app inventory)" -ForegroundColor Gray
                Write-Host "   |-- scripts.json (script inventory)" -ForegroundColor Gray
                Write-Host "   |-- remediations.json (remediation inventory)" -ForegroundColor Gray
                Write-Host "   |-- assignments.json (assignment data)" -ForegroundColor Gray
                Write-Host "   |-- inventory-runs.json (run history)" -ForegroundColor Gray
                Write-Host "   |-- reports/ (generated reports)" -ForegroundColor Gray
                Write-Host "   +-- source-code/ (source code files)" -ForegroundColor Gray
                Write-Host "`n[Ready to connect!] Use: Connect-IntuneInventory" -ForegroundColor Green
                
                return $StoragePath
            }
            else {
                throw "Failed to initialize JSON storage system"
            }
        }
        catch {
            $ErrorMessage = "Storage initialization failed: $($_.Exception.Message)"
            Write-IntuneInventoryLog -Message $ErrorMessage -Level Error
            throw $ErrorMessage
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Storage initialization completed" -Level Info -Source "Initialize-IntuneInventoryStorage"
    }
}

function Connect-IntuneInventory {
    <#
    .SYNOPSIS
    Connects to Microsoft Graph and initializes the JSON storage system.
    
    .DESCRIPTION
    Establishes connection to Microsoft Graph using production credentials and sets up
    the JSON-based storage system for inventory data.
    
    .PARAMETER UseWriteCredentials
    Use write-enabled Graph API credentials instead of read-only.
    
    .PARAMETER StoragePath
    Custom path for JSON storage. Defaults to user's AppData folder.
    
    .PARAMETER Force
    Force reconnection even if already connected.
    
    .EXAMPLE
    Connect-IntuneInventory
    
    .EXAMPLE
    Connect-IntuneInventory -UseWriteCredentials -StoragePath "C:\IntuneInventory"
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$UseWriteCredentials,
        
        [Parameter()]
        [string]$StoragePath = "$env:APPDATA\IntuneInventory",
        
        [Parameter()]
        [switch]$Force
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting Intune inventory connection using production credentials" -Level Info -Source "Connect-IntuneInventory"
        
        # Log automation initiation to production global log
        Write-ProductionLog -Message "IntuneInventory automation initiated by $env:USERNAME" -IsInitiation
    }
    
    process {
        try {
            # Check if already connected (unless force)
            if (-not $Force -and (Test-GraphToken) -and (Test-StorageConnection)) {
                Write-IntuneInventoryLog -Message "Already connected to Graph API and storage system" -Level Info
                
                # Show current status
                $ConnectionInfo = Get-IntuneConnectionInfo
                if ($ConnectionInfo) {
                    Write-Host "`n[Already Connected!]" -ForegroundColor Green
                    Write-Host "   Tenant: $($ConnectionInfo.TenantInfo.TenantName)" -ForegroundColor Cyan
                    Write-Host "   Token Expires: $($ConnectionInfo.GraphToken.Expires)" -ForegroundColor Cyan
                    Write-Host "   Storage: $($ConnectionInfo.StorageInfo.StorageRoot)" -ForegroundColor Cyan
                }
                return $true
            }
            
            # Acquire Graph API token
            Write-IntuneInventoryLog -Message "Acquiring Microsoft Graph access token..." -Level Info
            $TokenResult = Get-GraphAccessToken -UseWriteCredentials:$UseWriteCredentials -ForceRefresh:$Force
            
            if (-not $TokenResult.TokenAcquired) {
                throw "Failed to acquire Graph API token"
            }
            
            # Test Graph connection by getting organization info
            Write-IntuneInventoryLog -Message "Verifying Graph API connection..." -Level Info
            $OrgInfo = Get-GraphRequestAll -Uri "v1.0/organization"
            
            if ($OrgInfo -and (@($OrgInfo).Count -gt 0)) {
                $TenantName = $OrgInfo[0].displayName
                $TenantId = $OrgInfo[0].id
                Write-IntuneInventoryLog -Message "Connected to tenant: $TenantName ($TenantId)" -Level Info
                
                # Store connection info
                $Script:ConnectionInfo = @{
                    TenantName = $TenantName
                    TenantId = $TenantId
                    ClientId = if ($UseWriteCredentials) { $Script:ProductionClientIdWrite } else { $Script:ProductionClientId }
                    AuthMethod = "ClientCredentials"
                    ConnectedAt = Get-Date
                    TokenExpiry = $Script:TokenExpiry
                    UseWriteCredentials = $UseWriteCredentials.IsPresent
                }
            }
            else {
                throw "Failed to retrieve organization information from Graph API"
            }
            
            # Initialize JSON storage system
            Write-IntuneInventoryLog -Message "Initializing JSON storage system..." -Level Info
            $StorageResult = Initialize-JsonStorage -StoragePath $StoragePath
            
            if (-not $StorageResult) {
                throw "Failed to initialize JSON storage system"
            }
            
            Write-IntuneInventoryLog -Message "Storage initialized at: $StoragePath" -Level Info
            
            # Display connection summary
            Write-Host "`n[IntuneInventory Connected Successfully!]" -ForegroundColor Green
            Write-Host "   Tenant: $TenantName" -ForegroundColor Cyan
            Write-Host "   Storage: $StoragePath" -ForegroundColor Cyan
            Write-Host "   Token Expires: $($Script:TokenExpiry)" -ForegroundColor Cyan
            Write-Host "   Credentials: $(if ($UseWriteCredentials) { 'Write-Enabled' } else { 'Read-Only' })" -ForegroundColor Cyan
            
            # Show storage statistics
            $Stats = Get-StorageStatistics
            if ($Stats) {
                Write-Host "`n[Current Inventory]" -ForegroundColor Yellow
                Write-Host "   Applications: $($Stats.Applications)" -ForegroundColor White
                Write-Host "   Scripts: $($Stats.Scripts)" -ForegroundColor White
                Write-Host "   Remediations: $($Stats.Remediations)" -ForegroundColor White
                Write-Host "   Inventory Runs: $($Stats.InventoryRuns)" -ForegroundColor White
            }
            
            Write-Host "`n[Ready to run inventory operations!]" -ForegroundColor Green
            return $true
        }
        catch {
            $ErrorMessage = "Connection failed: $($_.Exception.Message)"
            Write-IntuneInventoryLog -Message $ErrorMessage -Level Error
            
            # Log critical failure to production global log
            Write-ProductionLog -Message "IntuneInventory connection failed: $($_.Exception.Message)" -IsFailure
            
            throw $ErrorMessage
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Connection process completed" -Level Info -Source "Connect-IntuneInventory"
    }
}

function Disconnect-IntuneInventory {
    <#
    .SYNOPSIS
    Disconnects from Graph API and cleans up resources.
    
    .DESCRIPTION
    Properly disconnects from Microsoft Graph API and cleans up authentication tokens
    and storage connections.
    #>
    [CmdletBinding()]
    param()
    
    begin {
        Write-IntuneInventoryLog -Message "Starting disconnect process" -Level Info -Source "Disconnect-IntuneInventory"
    }
    
    process {
        try {
            # Clear Graph API authentication
            $Script:GraphToken = $null
            $Script:TokenExpiry = $null
            $Script:GraphHeaders = @{}
            $Script:ConnectionInfo = $null
            
            # Clear storage variables
            $Script:StorageRoot = $null
            $Script:StoragePaths = @{}
            $Script:InventoryData = @{}
            
            Write-IntuneInventoryLog -Message "Successfully disconnected from Graph API and cleared storage cache" -Level Info
            Write-Host "[Disconnected from IntuneInventory successfully]" -ForegroundColor Green
            
            return $true
        }
        catch {
            Write-IntuneInventoryLog -Message "Error during disconnect: $($_.Exception.Message)" -Level Warning
            return $false
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Disconnect process completed" -Level Info -Source "Disconnect-IntuneInventory"
    }
}

function Test-IntuneConnection {
    <#
    .SYNOPSIS
    Tests the current connection status.
    
    .DESCRIPTION
    Verifies that both Graph API authentication and storage system are properly connected.
    #>
    [CmdletBinding()]
    param()
    
    try {
        $GraphConnected = Test-GraphToken
        $StorageConnected = Test-StorageConnection
        
        $ConnectionStatus = @{
            GraphAPIConnected = $GraphConnected
            StorageConnected = $StorageConnected
            FullyConnected = ($GraphConnected -and $StorageConnected)
            ConnectionInfo = $Script:ConnectionInfo
            StorageRoot = $Script:StorageRoot
            TokenExpiry = $Script:TokenExpiry
        }
        
        return [PSCustomObject]$ConnectionStatus
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to test connection: $($_.Exception.Message)" -Level Error
        return $null
    }
}

function Get-IntuneConnectionInfo {
    <#
    .SYNOPSIS
    Gets detailed information about the current connection.
    
    .DESCRIPTION
    Returns comprehensive information about the Graph API connection and storage system.
    #>
    [CmdletBinding()]
    param()
    
    try {
        $ConnectionTest = Test-IntuneConnection
        
        if ($ConnectionTest -and $ConnectionTest.FullyConnected) {
            $Info = @{
                Status = "Connected"
                TenantInfo = $Script:ConnectionInfo
                StorageInfo = Get-StorageStatistics
                GraphToken = @{
                    IsValid = Test-GraphToken
                    Expires = $Script:TokenExpiry
                    TimeUntilExpiry = if ($Script:TokenExpiry) { 
                        $Script:TokenExpiry - (Get-Date) 
                    } else { 
                        $null 
                    }
                }
            }
        }
        else {
            $Info = @{
                Status = "Disconnected"
                GraphAPI = if ($ConnectionTest) { $ConnectionTest.GraphAPIConnected } else { $false }
                Storage = if ($ConnectionTest) { $ConnectionTest.StorageConnected } else { $false }
                Message = "Use Connect-IntuneInventory to establish connection"
            }
        }
        
        return [PSCustomObject]$Info
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to get connection info: $($_.Exception.Message)" -Level Error
        return $null
    }
}

function Invoke-CompleteIntuneInventory {
    <#
    .SYNOPSIS
    Runs a complete Intune inventory collection including all apps, scripts, and remediations with source code and assignments.
    
    .DESCRIPTION
    This is a comprehensive helper function that:
    1. Connects to Microsoft Graph and initializes storage
    2. Runs application inventory with Force and assignments
    3. Runs device management script inventory with Force, source code, and assignments
    4. Runs remediation script inventory with Force, source code, and assignments
    5. Provides a summary of all collected data
    
    .PARAMETER StoragePath
    Custom path for JSON storage. Defaults to user's AppData folder if not specified.
    
    .PARAMETER SkipAssignments
    Skip assignment collection (assignments are included by default).
    
    .EXAMPLE
    Invoke-CompleteIntuneInventory
    
    .EXAMPLE
    Invoke-CompleteIntuneInventory -StoragePath "C:\IntuneInventory" 
    
    .EXAMPLE
    Invoke-CompleteIntuneInventory -SkipAssignments
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$StoragePath = "$env:APPDATA\IntuneInventory",
        
        [Parameter()]
        [switch]$SkipAssignments
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting complete Intune inventory collection" -Level Info -Source "Invoke-CompleteIntuneInventory"
        $StartTime = Get-Date
    }
    
    process {
        try {
            $Results = @{
                Success = $true
                StartTime = $StartTime.ToString("yyyy-MM-dd HH:mm:ss")
                ConnectionResult = $null
                ApplicationResult = $null
                ScriptResult = $null
                RemediationResult = $null
                Summary = @{}
                ErrorMessages = @()
                EndTime = $null
                Duration = $null
            }
            
            # Step 1: Connect to Intune
            Write-Host "=== STEP 1: CONNECTING TO INTUNE ===" -ForegroundColor Green
            try {
                $ConnectionResult = Connect-IntuneInventory -StoragePath $StoragePath -Force
                $Results.ConnectionResult = $ConnectionResult
                Write-Host "✅ Connection established successfully" -ForegroundColor Green
            }
            catch {
                $ErrorMessage = "Failed to connect to Intune: $($_.Exception.Message)"
                $Results.ErrorMessages += $ErrorMessage
                $Results.Success = $false
                Write-Host "❌ Connection failed: $ErrorMessage" -ForegroundColor Red
                throw $ErrorMessage
            }
            
            # Step 2: Application Inventory
            Write-Host "`n=== STEP 2: APPLICATION INVENTORY ===" -ForegroundColor Green
            try {
                $AppParams = @{
                    Force = $true
                }
                if (-not $SkipAssignments) { $AppParams.IncludeAssignments = $true }
                
                $ApplicationResult = Invoke-IntuneApplicationInventory @AppParams
                $Results.ApplicationResult = $ApplicationResult
                Write-Host "✅ Applications: $($ApplicationResult.ItemsProcessed) collected" -ForegroundColor Green
            }
            catch {
                $ErrorMessage = "Application inventory failed: $($_.Exception.Message)"
                $Results.ErrorMessages += $ErrorMessage
                Write-Host "❌ Application inventory failed: $ErrorMessage" -ForegroundColor Red
            }
            
            # Step 3: Device Management Scripts Inventory
            Write-Host "`n=== STEP 3: DEVICE MANAGEMENT SCRIPTS INVENTORY ===" -ForegroundColor Green
            try {
                $ScriptParams = @{
                    Force = $true
                    IncludeSourceCode = $true
                }
                if (-not $SkipAssignments) { $ScriptParams.IncludeAssignments = $true }
                
                $ScriptResult = Invoke-IntuneScriptInventory @ScriptParams
                $Results.ScriptResult = $ScriptResult
                Write-Host "✅ Scripts: $($ScriptResult.ItemsProcessed) collected with source code" -ForegroundColor Green
            }
            catch {
                $ErrorMessage = "Script inventory failed: $($_.Exception.Message)"
                $Results.ErrorMessages += $ErrorMessage
                Write-Host "❌ Script inventory failed: $ErrorMessage" -ForegroundColor Red
            }
            
            # Step 4: Remediation Scripts Inventory
            Write-Host "`n=== STEP 4: REMEDIATION SCRIPTS INVENTORY ===" -ForegroundColor Green
            try {
                $RemediationParams = @{
                    Force = $true
                    IncludeSourceCode = $true
                }
                if (-not $SkipAssignments) { $RemediationParams.IncludeAssignments = $true }
                
                $RemediationResult = Invoke-IntuneRemediationInventory @RemediationParams
                $Results.RemediationResult = $RemediationResult
                Write-Host "✅ Remediations: $($RemediationResult.ItemsProcessed) collected with source code" -ForegroundColor Green
            }
            catch {
                $ErrorMessage = "Remediation inventory failed: $($_.Exception.Message)"
                $Results.ErrorMessages += $ErrorMessage
                Write-Host "❌ Remediation inventory failed: $ErrorMessage" -ForegroundColor Red
            }
            
            # Step 5: Generate Summary
            Write-Host "`n=== INVENTORY SUMMARY ===" -ForegroundColor Green
            $EndTime = Get-Date
            $Duration = ($EndTime - $StartTime).TotalMinutes
            
            $Results.EndTime = $EndTime.ToString("yyyy-MM-dd HH:mm:ss")
            $Results.Duration = [math]::Round($Duration, 2)
            
            $Summary = @{
                TotalApplications = if ($ApplicationResult) { $ApplicationResult.ItemsProcessed } else { 0 }
                TotalScripts = if ($ScriptResult) { $ScriptResult.ItemsProcessed } else { 0 }
                TotalRemediations = if ($RemediationResult) { $RemediationResult.ItemsProcessed } else { 0 }
                TotalErrors = $Results.ErrorMessages.Count
                Duration = $Results.Duration
                StoragePath = $Script:StorageRoot
            }
            $Results.Summary = $Summary
            
            Write-Host "📊 Applications: $($Summary.TotalApplications)" -ForegroundColor Cyan
            Write-Host "📄 Device Management Scripts: $($Summary.TotalScripts)" -ForegroundColor Cyan
            Write-Host "🔧 Remediation Scripts: $($Summary.TotalRemediations)" -ForegroundColor Cyan
            if (-not $SkipAssignments) {
                Write-Host "🎯 Assignments: Included for all items" -ForegroundColor Cyan
            } else {
                Write-Host "⚠️  Assignments: Skipped" -ForegroundColor Yellow
            }
            Write-Host "⏱️  Total Duration: $($Summary.Duration) minutes" -ForegroundColor Cyan
            Write-Host "💾 Storage Location: $($Summary.StoragePath)" -ForegroundColor Cyan
            
            if ($Results.ErrorMessages.Count -gt 0) {
                Write-Host "⚠️  Errors: $($Results.ErrorMessages.Count)" -ForegroundColor Yellow
                foreach ($ErrorMsg in $Results.ErrorMessages) {
                    Write-Host "   - $ErrorMsg" -ForegroundColor Yellow
                }
            }
            
            $TotalItems = $Summary.TotalApplications + $Summary.TotalScripts + $Summary.TotalRemediations
            Write-Host "`n🎉 Complete inventory finished! Total items collected: $TotalItems" -ForegroundColor Green
            
            return $Results
        }
        catch {
            $ErrorMessage = "Critical error during complete inventory: $($_.Exception.Message)"
            Write-IntuneInventoryLog -Message $ErrorMessage -Level Error
            Write-Host "❌ Complete inventory failed: $ErrorMessage" -ForegroundColor Red
            
            $Results.Success = $false
            $Results.ErrorMessages += $ErrorMessage
            $Results.EndTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            $Results.Duration = [math]::Round(((Get-Date) - $StartTime).TotalMinutes, 2)
            
            return $Results
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Complete Intune inventory operation finished" -Level Info -Source "Invoke-CompleteIntuneInventory"
    }
}
