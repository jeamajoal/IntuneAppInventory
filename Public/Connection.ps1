function Initialize-IntuneInventoryDatabase {
    <#
    .SYNOPSIS
    Initializes the SQLite database for Intune inventory storage.
    
    .DESCRIPTION
    Creates or updates the SQLite database schema for storing Intune applications,
    scripts, remediations, and related inventory data.
    
    .PARAMETER DatabasePath
    The full path where the SQLite database file should be created or updated.
    If not specified, defaults to a database in the user's AppData folder.
    
    .PARAMETER Force
    If specified, recreates the database even if it already exists.
    
    .EXAMPLE
    Initialize-IntuneInventoryDatabase
    
    .EXAMPLE
    Initialize-IntuneInventoryDatabase -DatabasePath "C:\IntuneInventory\inventory.db"
    
    .EXAMPLE
    Initialize-IntuneInventoryDatabase -Force
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string]$DatabasePath = (Join-Path -Path $env:APPDATA -ChildPath "IntuneInventory\inventory.db"),
        
        [Parameter()]
        [switch]$Force
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting database initialization" -Level Info -Source "Initialize-IntuneInventoryDatabase"
    }
    
    process {
        try {
            # Check if database exists and handle Force parameter
            if ((Test-Path -Path $DatabasePath) -and -not $Force) {
                Write-IntuneInventoryLog -Message "Database already exists at: $DatabasePath" -Level Info
                
                # Test existing database
                $TestConnection = Get-DatabaseConnection -DatabasePath $DatabasePath
                if (Test-DatabaseConnection -DatabaseConnection $TestConnection) {
                    Write-IntuneInventoryLog -Message "Existing database is valid and accessible" -Level Info
                    Close-DatabaseConnection -DatabaseConnection $TestConnection
                    
                    # Update script variables
                    $Script:DatabasePath = $DatabasePath
                    return
                }
                else {
                    Write-IntuneInventoryLog -Message "Existing database is corrupted, will recreate" -Level Warning
                    Close-DatabaseConnection -DatabaseConnection $TestConnection
                    $Force = $true
                }
            }
            
            if ($Force -and (Test-Path -Path $DatabasePath)) {
                if ($PSCmdlet.ShouldProcess($DatabasePath, "Delete existing database")) {
                    Remove-Item -Path $DatabasePath -Force
                    Write-IntuneInventoryLog -Message "Existing database deleted" -Level Info
                }
            }
            
            if ($PSCmdlet.ShouldProcess($DatabasePath, "Create Intune inventory database")) {
                # Create database connection
                $DatabaseConnection = Get-DatabaseConnection -DatabasePath $DatabasePath
                
                # Initialize schema
                Initialize-DatabaseSchema -DatabaseConnection $DatabaseConnection
                
                # Test the new database
                if (Test-DatabaseConnection -DatabaseConnection $DatabaseConnection) {
                    Write-IntuneInventoryLog -Message "Database initialized successfully at: $DatabasePath" -Level Info
                    
                    # Update script variables
                    $Script:DatabasePath = $DatabasePath
                    $Script:DatabaseConnection = $DatabaseConnection
                    
                    # Create initial inventory run record
                    $UserContext = Get-CurrentUserContext
                    $Command = $DatabaseConnection.CreateCommand()
                    $Command.CommandText = @"
INSERT INTO InventoryRuns (RunType, StartTime, Status, TenantId, UserPrincipalName)
VALUES (@RunType, @StartTime, @Status, @TenantId, @UserPrincipalName);
"@
                    $null = $Command.Parameters.AddWithValue("@RunType", "DatabaseInitialization")
                    $null = $Command.Parameters.AddWithValue("@StartTime", (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
                    $null = $Command.Parameters.AddWithValue("@Status", "Completed")
                    $null = $Command.Parameters.AddWithValue("@TenantId", $UserContext.TenantId)
                    $null = $Command.Parameters.AddWithValue("@UserPrincipalName", $UserContext.UserPrincipalName)
                    
                    $null = $Command.ExecuteNonQuery()
                    $Command.Dispose()
                    
                    Write-Host "Database initialized successfully!" -ForegroundColor Green
                    Write-Host "Database location: $DatabasePath" -ForegroundColor Cyan
                }
                else {
                    throw "Database validation failed after initialization"
                }
            }
        }
        catch {
            Write-IntuneInventoryLog -Message "Database initialization failed: $($_.Exception.Message)" -Level Error
            throw
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Database initialization completed" -Level Info -Source "Initialize-IntuneInventoryDatabase"
    }
}

function Connect-IntuneInventory {
    <#
    .SYNOPSIS
    Establishes connection to Microsoft Graph for Intune inventory operations using production credentials.
    
    .DESCRIPTION
    Authenticates to Microsoft Graph using the production-standard authentication pattern with
    the required permissions for inventorying Intune applications, scripts, and remediations.
    Also establishes database connection.
    
    .PARAMETER UseWriteCredentials
    Use write-enabled credentials instead of read-only credentials for operations that require write access.
    
    .PARAMETER DatabasePath
    The path to the SQLite database. If not specified, uses the module default.
    
    .EXAMPLE
    Connect-IntuneInventory
    
    .EXAMPLE
    Connect-IntuneInventory -UseWriteCredentials
    
    .EXAMPLE
    Connect-IntuneInventory -DatabasePath "C:\IntuneInventory\inventory.db"
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$UseWriteCredentials,
        
        [Parameter()]
        [string]$DatabasePath
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting Intune inventory connection using production credentials" -Level Info -Source "Connect-IntuneInventory"
    }
    
    process {
        try {
            # Acquire token using production standard
            Write-IntuneInventoryLog -Message "Acquiring Microsoft Graph access token..." -Level Info
            $TokenResult = Get-GraphAccessToken -UseWriteCredentials:$UseWriteCredentials
            
            if (-not $TokenResult.TokenAcquired) {
                throw "Failed to acquire Microsoft Graph access token"
            }
            
            # Verify authentication and permissions
            if (-not (Test-GraphAuthentication)) {
                throw "Microsoft Graph authentication or permissions verification failed"
            }
            
            # Get organization info for connection details
            $OrgInfo = Get-GraphRequestAll -Uri "v1.0/organization"
            $OrgDetails = $OrgInfo[0]
            
            Write-IntuneInventoryLog -Message "Connected to tenant: $($OrgDetails.displayName) ($($Script:ProductionTenantId))" -Level Info
            
            # Store connection information
            $Script:ConnectionInfo = @{
                TenantId = $Script:ProductionTenantId
                ClientId = if ($UseWriteCredentials) { $Script:ProductionClientIdWrite } else { $Script:ProductionClientId }
                UserPrincipalName = $OrgDetails.displayName
                ConnectedAt = Get-Date
                AuthMethod = "ClientCredentials"
                TokenExpiry = $Script:TokenExpiry
                UseWriteCredentials = $UseWriteCredentials.IsPresent
            }
            
            # Establish database connection
            if ($DatabasePath) {
                $Script:DatabasePath = $DatabasePath
            }
            elseif (-not $Script:DatabasePath) {
                $Script:DatabasePath = Join-Path -Path $env:APPDATA -ChildPath "IntuneInventory\inventory.db"
            }
            
            # Ensure database is initialized
            if (-not (Test-Path -Path $Script:DatabasePath)) {
                Write-IntuneInventoryLog -Message "Database not found, initializing..." -Level Info
                Initialize-IntuneInventoryDatabase -DatabasePath $Script:DatabasePath
            }
            else {
                # Connect to existing database
                $Script:DatabaseConnection = Get-DatabaseConnection -DatabasePath $Script:DatabasePath
                if (-not (Test-DatabaseConnection -DatabaseConnection $Script:DatabaseConnection)) {
                    throw "Unable to connect to existing database at: $Script:DatabasePath"
                }
                Write-IntuneInventoryLog -Message "Connected to database: $Script:DatabasePath" -Level Info
            }
            
            Write-Host "Successfully connected to Intune inventory system!" -ForegroundColor Green
            Write-Host "Tenant: $($OrgDetails.displayName)" -ForegroundColor Cyan
            Write-Host "Organization: $($Script:ProductionTenantId)" -ForegroundColor Cyan
            Write-Host "Client ID: $($Script:ConnectionInfo.ClientId)" -ForegroundColor Cyan
            Write-Host "Database: $Script:DatabasePath" -ForegroundColor Cyan
            Write-Host "Token expires: $($Script:TokenExpiry.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Yellow
            if ($UseWriteCredentials) {
                Write-Host "Using WRITE credentials - Full Intune management access" -ForegroundColor Yellow
            } else {
                Write-Host "Using READ credentials - Read-only access" -ForegroundColor Green
            }
        }
        catch {
            Write-IntuneInventoryLog -Message "Connection failed: $($_.Exception.Message)" -Level Error
            throw
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Intune inventory connection completed" -Level Info -Source "Connect-IntuneInventory"
    }
}

function Disconnect-IntuneInventory {
    <#
    .SYNOPSIS
    Disconnects from Microsoft Graph and closes database connections.
    
    .DESCRIPTION
    Properly clears the Graph token cache and closes the SQLite database connection
    to clean up resources using production-standard cleanup procedures.
    
    .EXAMPLE
    Disconnect-IntuneInventory
    #>
    [CmdletBinding()]
    param()
    
    begin {
        Write-IntuneInventoryLog -Message "Starting Intune inventory disconnection" -Level Info -Source "Disconnect-IntuneInventory"
    }
    
    process {
        try {
            # Close database connection
            if ($Script:DatabaseConnection) {
                Close-DatabaseConnection -DatabaseConnection $Script:DatabaseConnection
                $Script:DatabaseConnection = $null
                Write-IntuneInventoryLog -Message "Database connection closed" -Level Info
            }
            
            # Clear Graph token and cache (production pattern)
            $Script:GraphToken = $null
            $Script:TokenExpiry = $null
            $Script:GraphHeaders = @{}
            $Script:ConnectionInfo = $null
            
            Write-IntuneInventoryLog -Message "Graph authentication cleared" -Level Info
            
            Write-Host "Disconnected from Intune inventory system" -ForegroundColor Yellow
        }
        catch {
            Write-IntuneInventoryLog -Message "Disconnection error: $($_.Exception.Message)" -Level Error
            throw
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Intune inventory disconnection completed" -Level Info -Source "Disconnect-IntuneInventory"
    }
}
