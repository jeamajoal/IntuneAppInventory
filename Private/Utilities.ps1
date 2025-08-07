# Module-scoped variables for Graph API - Production Pattern
$Script:GraphToken = $null
$Script:TokenExpiry = $null
$Script:GraphHeaders = @{}
$Script:GraphBaseUrl = "https://graph.microsoft.com/v1.0"
$Script:ConnectionInfo = $null

# Production credential configuration
$Script:CredRoot = 'C:\Scheduled_Task_Resources\KeysCreds'
$Script:ProductionTenantId = '06b476ce-d8bc-4355-a477-0392dd2dc025'
$Script:ProductionClientId = 'ab9240e1-f604-40e0-a193-c3ac9d817077'
$Script:ProductionClientIdWrite = '23582b3b-26a4-4ec6-84ec-6877195c22bf'

function Get-GraphAccessToken {
    <#
    .SYNOPSIS
    Acquires an access token for Microsoft Graph API using production credentials.
    
    .DESCRIPTION
    Gets an access token using the production Graph API app registration and credentials.
    Uses the standardized authentication pattern from production automation scripts.
    
    .PARAMETER UseWriteCredentials
    Use the write-enabled app registration instead of read-only credentials.
    
    .PARAMETER ForceRefresh
    Force acquisition of a new token even if cached token is valid.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$UseWriteCredentials,
        
        [Parameter()]
        [switch]$ForceRefresh
    )
    
    begin {
        Write-Verbose "Get-GraphAccessToken: Starting production token acquisition"
        
        # Check if we have a valid token unless force refresh
        if (-not $ForceRefresh -and $Script:GraphToken -and $Script:TokenExpiry) {
            if ($Script:TokenExpiry -gt (Get-Date).AddMinutes(5)) {
                Write-Verbose "Using cached token (expires: $($Script:TokenExpiry))"
                return @{
                    TokenAcquired = $true
                    Source = "Cache"
                    ExpiresAt = $Script:TokenExpiry
                }
            }
        }
    }
    
    process {
        try {
            # Use production credentials
            $TenantId = $Script:ProductionTenantId
            $ClientId = if ($UseWriteCredentials) { $Script:ProductionClientIdWrite } else { $Script:ProductionClientId }
            
            # Load client secret from production location
            $SecretFile = if ($UseWriteCredentials) { 'graph-write.txt' } else { 'graph.txt' }
            $SecretPath = Join-Path $Script:CredRoot $SecretFile
            
            if (-not (Test-Path $SecretPath)) {
                throw "Cannot find credential file: $SecretPath"
            }
            
            $ClientSecret = Get-Content $SecretPath
            
            Write-Verbose "Using production credentials: TenantId=$TenantId, ClientId=$ClientId"
            
            # Build OAuth2 request - production standard
            $Body = @{
                Grant_Type    = "client_credentials"
                Scope         = "https://graph.microsoft.com/.default"
                Client_Id     = $ClientId
                Client_Secret = $ClientSecret
            }
            
            $TokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
            
            # Make token request
            Write-Verbose "Requesting token from $TokenEndpoint"
            $Response = Invoke-RestMethod -Method Post -Uri $TokenEndpoint -Body $Body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
            
            # Calculate expiry with buffer
            $ExpiryTime = (Get-Date).AddSeconds($Response.expires_in - 300) # 5 min buffer
            
            # Update module variables
            $Script:GraphToken = $Response.access_token
            $Script:TokenExpiry = $ExpiryTime
            $Script:GraphHeaders = @{
                'Authorization' = "Bearer $($Response.access_token)"
                'Content-Type' = 'application/json'
                'Accept' = 'application/json'
            }
            
            Write-Verbose "Token acquired successfully (expires: $ExpiryTime)"
            return @{
                TokenAcquired = $true
                Source = "OAuth"
                ExpiresAt = $ExpiryTime
                Scope = $Response.scope
            }
        }
        catch {
            $ErrorBody = $null
            if ($_.Exception.Response) {
                $Reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $Reader.BaseStream.Position = 0
                $ErrorBody = $Reader.ReadToEnd() | ConvertFrom-Json -ErrorAction SilentlyContinue
            }
            
            if ($ErrorBody -and $ErrorBody.error) {
                throw "OAuth token acquisition failed: $($ErrorBody.error) - $($ErrorBody.error_description)"
            }
            else {
                throw "OAuth token acquisition failed: $_"
            }
        }
        finally {
            # Clear sensitive variables
            if ($ClientSecret) {
                Clear-Variable ClientSecret -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

function Test-GraphToken {
    <#
    .SYNOPSIS
    Tests if the current Graph token is valid.
    
    .DESCRIPTION
    Validates the current Graph API token by checking expiry and optionally making a test API call.
    #>
    [CmdletBinding()]
    param()
    
    try {
        # Check if token exists
        if ([string]::IsNullOrWhiteSpace($Script:GraphToken)) {
            return $false
        }
        
        # Check expiry
        if ($Script:TokenExpiry -and $Script:TokenExpiry -gt (Get-Date).AddMinutes(5)) {
            # Test with organization endpoint
            $OrgUri = "$Script:GraphBaseUrl/organization"
            $OrgInfo = Invoke-RestMethod -Uri $OrgUri -Headers $Script:GraphHeaders -Method Get -ErrorAction Stop
            return ($OrgInfo.value -and $OrgInfo.value.Count -gt 0)
        }
        
        return $false
    }
    catch {
        Write-Verbose "Token validation failed: $($_.Exception.Message)"
        return $false
    }
}

function Get-GraphRequestAll {
    <#
    .SYNOPSIS
    Makes paginated GET requests to Microsoft Graph API using production standard.
    
    .DESCRIPTION
    Production-standard function for making GET requests to Microsoft Graph API with
    automatic pagination, retry logic, and error handling. Based on Get-SoundMgGraphRequestAll-v2.
    
    .PARAMETER Uri
    The Graph API endpoint URI (relative or absolute).
    
    .PARAMETER Headers
    Additional headers to include in the request.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        
        [Parameter()]
        [hashtable]$Headers = @{}
    )
    
    try {
        # Ensure we have a valid token
        if (-not (Test-GraphToken)) {
            throw "No valid Graph token available. Please authenticate first."
        }
        
        # Make the initial request - ensure URI is well-formed
        if ([System.Uri]::IsWellFormedUriString($Uri, [System.UriKind]::Absolute)) {
            # URI is already a complete URL
            $requestUri = $Uri
        } elseif ([System.Uri]::IsWellFormedUriString("https://graph.microsoft.com/$Uri", [System.UriKind]::Absolute)) {
            # URI is a relative path, prepend base URL
            $requestUri = "https://graph.microsoft.com/$Uri"
        } else {
            throw "Invalid or malformed URI: $Uri"
        }
        
        # Prepare headers - merge default auth headers with custom headers
        $requestHeaders = @{
            'Authorization' = "Bearer $Script:GraphToken"
            'Content-Type' = 'application/json'
        }
        
        # Add any additional headers (like ConsistencyLevel for advanced queries)
        foreach ($key in $Headers.Keys) {
            $requestHeaders[$key] = $Headers[$key]
        }
        
        # Initialize variables for pagination
        $nextLink = $requestUri
        $allResults = @()
        $retryCount = 0
        $maxRetries = 3
        $retryDelaySeconds = 5
        $pageCount = 0
        
        while ($nextLink) {
            try {
                Write-Verbose "Making Graph API request to: $nextLink (Page: $($pageCount + 1))"
                
                # Make the Graph API request with headers
                $response = Invoke-RestMethod -Uri $nextLink -Headers $requestHeaders -Method Get
                
                # Add results to collection
                if ($response.value) {
                    $allResults += $response.value
                    Write-Verbose "Retrieved $($response.value.Count) items from page $($pageCount + 1)"
                }
                
                # Move to next page - handle multiple possible property names
                $nextLink = $null
                if ($response.PSObject.Properties['@odata.nextLink']) {
                    $nextLink = $response.'@odata.nextLink'
                } elseif ($response.PSObject.Properties['odata.nextLink']) {
                    $nextLink = $response.'odata.nextLink'
                } elseif ($response.PSObject.Properties['nextLink']) {
                    $nextLink = $response.nextLink
                }
                
                $pageCount++
                $retryCount = 0  # Reset retry count on successful request
                
                if (-not $nextLink) {
                    Write-Verbose "No more pages to retrieve."
                    break
                }
            } catch {
                $retryCount++
                Write-Warning "Graph API request failed (Attempt $retryCount/$maxRetries) for URL: $nextLink - $($_.Exception.Message)"
                
                if ($retryCount -ge $maxRetries) {
                    throw "Failed to fetch data from $nextLink after $maxRetries retries. Last error: $($_.Exception.Message)"
                }
                
                Write-Verbose "Retrying request to: $nextLink in $retryDelaySeconds seconds"
                Start-Sleep -Seconds $retryDelaySeconds
                # Note: $nextLink stays the same for retry
            }
        }
        
        Write-Verbose "Completed Graph API pagination. Total pages: $pageCount, Total items: $($allResults.Count)"
        return $allResults
    } catch {
        Write-Error "Error making Graph API request: $($_.Exception.Message)"
        throw
    }
}

function Invoke-GraphRequest {
    <#
    .SYNOPSIS
    Makes HTTP requests to Microsoft Graph API with robust error handling.
    
    .DESCRIPTION
    Production-standard function for making various HTTP requests to Microsoft Graph API
    with comprehensive retry logic and error handling. Based on Invoke-RobustGraphRequest.
    
    .PARAMETER Uri
    The Graph API endpoint URI.
    
    .PARAMETER Method
    HTTP method (GET, POST, PUT, DELETE, PATCH).
    
    .PARAMETER Body
    The request body for POST/PUT/PATCH requests.
    
    .PARAMETER All
    Automatically handle pagination and return all results (GET requests only).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        
        [Parameter()]
        [Microsoft.PowerShell.Commands.WebRequestMethod]$Method = 'GET',
        
        [Parameter()]
        [object]$Body,
        
        [Parameter()]
        [switch]$All
    )
    
    try {
        # For GET requests with All switch, use pagination function
        if ($Method -eq 'GET' -and $All) {
            return Get-GraphRequestAll -Uri $Uri
        }
        
        # Ensure we have a valid token
        if (-not (Test-GraphToken)) {
            throw "No valid Graph token available. Please authenticate first."
        }
        
        # Construct full URI if needed
        if ([System.Uri]::IsWellFormedUriString($Uri, [System.UriKind]::Absolute)) {
            # URI is already a complete URL
            $requestUri = $Uri
        } elseif ([System.Uri]::IsWellFormedUriString("https://graph.microsoft.com/$Uri", [System.UriKind]::Absolute)) {
            # URI is a relative path, prepend base URL
            $requestUri = "https://graph.microsoft.com/$Uri"
        } else {
            throw "Invalid or malformed URI: $Uri"
        }

        # Prepare request parameters
        $RequestParams = @{
            Uri = $requestUri
            Method = $Method
            Headers = $Script:GraphHeaders
            ErrorAction = 'Stop'
        }
        
        if ($Body) {
            if ($Body -is [string]) {
                $RequestParams.Body = $Body
            }
            else {
                $RequestParams.Body = ($Body | ConvertTo-Json -Depth 10 -Compress)
            }
        }
        
        # Make request with retry logic
        $maxRetries = 3
        $retryDelay = 5
        
        for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
            try {
                $Response = Invoke-RestMethod @RequestParams
                Write-Verbose "Graph API request completed successfully"
                return $Response
            }
            catch {
                $statusCode = $null
                if ($_.Exception.Response) {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }
                
                Write-Verbose "Graph API request attempt $attempt failed. Status: $statusCode"
                
                # Check for non-retryable errors
                if ($statusCode -ge 400 -and $statusCode -le 499) {
                    # Client errors - don't retry
                    throw "Graph API request failed with status $statusCode : $($_.Exception.Message)"
                }
                
                if ($attempt -ge $maxRetries) {
                    throw "Graph API request failed after $maxRetries retries: $($_.Exception.Message)"
                }
                
                Write-Verbose "Retrying Graph API request in $retryDelay seconds..."
                Start-Sleep -Seconds $retryDelay
            }
        }
    }
    catch {
        $ErrorMessage = "Graph API request failed: $($_.Exception.Message)"
        Write-Error $ErrorMessage
        throw
    }
}
function Write-IntuneInventoryLog {
    <#
    .SYNOPSIS
    Writes log messages for the Intune Inventory module using production logging standard.
    
    .DESCRIPTION
    Centralized logging function that writes messages to various outputs.
    Note: This does NOT use Write-GlobalLog - that should only be used for automation initiation and failures.
    
    .PARAMETER Message
    The log message to write.
    
    .PARAMETER Level
    The log level (Info, Warning, Error, Debug, Verbose).
    
    .PARAMETER Source
    The source component or function generating the log.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet('Info', 'Warning', 'Error', 'Debug', 'Verbose')]
        [string]$Level = 'Info',
        
        [Parameter()]
        [string]$Source = 'IntuneInventory'
    )
    
    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $LogMessage = "[$Timestamp] [$Level] [$Source] $Message"
    
    # Standard output based on level
    switch ($Level) {
        'Error' { Write-Error $LogMessage }
        'Warning' { Write-Warning $LogMessage }
        'Debug' { Write-Debug $LogMessage }
        'Verbose' { Write-Verbose $LogMessage }
        default { Write-Host $LogMessage }
    }
}

function Write-ProductionLog {
    <#
    .SYNOPSIS
    Writes to the production global logging system following production patterns.
    
    .DESCRIPTION
    Used ONLY for automation initiation and critical failures, following the production
    standard of using Write-GlobalLog sparingly (once for initiation, once for failures).
    
    .PARAMETER Message
    The message to write to global log.
    
    .PARAMETER IsInitiation
    Indicates this is an automation initiation log.
    
    .PARAMETER IsFailure
    Indicates this is a critical failure log.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter()]
        [switch]$IsInitiation,
        
        [Parameter()]
        [switch]$IsFailure
    )
    
    # Only log to global system for initiation or failures
    if ($IsInitiation -or $IsFailure) {
        try {
            if (Get-Command Write-GlobalLog -ErrorAction SilentlyContinue) {
                Write-GlobalLog -Message $Message -scriptName "IntuneInventory"
                Write-IntuneInventoryLog -Message "Global log written: $Message" -Level Verbose
            }
            else {
                Write-IntuneInventoryLog -Message "Write-GlobalLog not available. Message would have been: $Message" -Level Verbose
            }
        }
        catch {
            Write-IntuneInventoryLog -Message "Failed to write global log: $($_.Exception.Message)" -Level Warning
        }
    }
}

function ConvertTo-SafeJson {
    <#
    .SYNOPSIS
    Converts an object to JSON with error handling.
    
    .DESCRIPTION
    Safely converts PowerShell objects to JSON strings with proper error handling
    and depth control.
    
    .PARAMETER InputObject
    The object to convert to JSON.
    
    .PARAMETER Depth
    The maximum depth for JSON conversion.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,
        
        [Parameter()]
        [int]$Depth = 10
    )
    
    try {
        if ($null -eq $InputObject) {
            return $null
        }
        
        return ($InputObject | ConvertTo-Json -Depth $Depth -Compress)
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to convert object to JSON: $($_.Exception.Message)" -Level Warning
        return $null
    }
}

function ConvertFrom-SafeJson {
    <#
    .SYNOPSIS
    Converts JSON to an object with error handling.
    
    .DESCRIPTION
    Safely converts JSON strings to PowerShell objects with proper error handling.
    
    .PARAMETER JsonString
    The JSON string to convert.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$JsonString
    )
    
    try {
        if ([string]::IsNullOrWhiteSpace($JsonString)) {
            return $null
        }
        
        return ($JsonString | ConvertFrom-Json)
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to convert JSON to object: $($_.Exception.Message)" -Level Warning
        return $null
    }
}

function Get-CurrentUserContext {
    <#
    .SYNOPSIS
    Gets the current user context for audit purposes.
    
    .DESCRIPTION
    Returns information about the current user session for logging and audit trails.
    Uses production-standard authentication context.
    #>
    [CmdletBinding()]
    param()
    
    try {
        if ($Script:ConnectionInfo) {
            return @{
                UserPrincipalName = $Script:ConnectionInfo.UserPrincipalName
                TenantId = $Script:ConnectionInfo.TenantId
                ClientId = $Script:ConnectionInfo.ClientId
                AuthMethod = $Script:ConnectionInfo.AuthMethod
            }
        }
        else {
            # Return production standard context
            return @{
                UserPrincipalName = "Production Service Account"
                TenantId = $Script:ProductionTenantId
                ClientId = $Script:ProductionClientId
                AuthMethod = "ClientCredentials"
            }
        }
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to get user context: $($_.Exception.Message)" -Level Warning
        return @{
            UserPrincipalName = $env:USERNAME
            TenantId = $Script:ProductionTenantId
            ClientId = $Script:ProductionClientId
            AuthMethod = "Unknown"
        }
    }
}

function Resolve-GroupAssignments {
    <#
    .SYNOPSIS
    Resolves group IDs in assignment targets to include display names.
    
    .DESCRIPTION
    Takes assignment data and resolves any group IDs to their display names,
    adding GroupDisplayNames array to the assignment record.
    
    .PARAMETER Assignments
    Array of assignment objects to process.
    
    .EXAMPLE
    $EnhancedAssignments = Resolve-GroupAssignments -Assignments $AppAssignments
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Assignments
    )
    
    begin {
        Write-Verbose "Starting group ID resolution for $($Assignments.Count) assignments"
        
        # Cache for group lookups to avoid duplicate API calls
        $Script:GroupCache = @{}
    }
    
    process {
        try {
            foreach ($Assignment in $Assignments) {
                # Initialize group display names array
                $GroupDisplayNames = @()
                
                # Check if assignment has a group ID
                if ($Assignment.target -and $Assignment.target.groupId) {
                    $GroupId = $Assignment.target.groupId
                    
                    # Check cache first
                    if ($Script:GroupCache.ContainsKey($GroupId)) {
                        $GroupDisplayNames += $Script:GroupCache[$GroupId]
                        Write-Verbose "Using cached group name for ${GroupId}: $($Script:GroupCache[$GroupId])"
                    }
                    else {
                        # Look up group from Graph API
                        try {
                            Write-Verbose "Looking up group ID: $GroupId"
                            $GroupInfo = Invoke-GraphRequest -Uri "v1.0/groups/$GroupId" -Method GET
                            
                            if ($GroupInfo -and $GroupInfo.displayName) {
                                $DisplayName = $GroupInfo.displayName
                                $GroupDisplayNames += $DisplayName
                                $Script:GroupCache[$GroupId] = $DisplayName
                                Write-Verbose "Resolved group $GroupId to: $DisplayName"
                            }
                            else {
                                Write-Warning "Could not resolve group ID: $GroupId"
                                $GroupDisplayNames += "Unknown Group ($GroupId)"
                                $Script:GroupCache[$GroupId] = "Unknown Group ($GroupId)"
                            }
                        }
                        catch {
                            Write-Warning "Failed to resolve group ID $GroupId : $($_.Exception.Message)"
                            $GroupDisplayNames += "Error Resolving ($GroupId)"
                            $Script:GroupCache[$GroupId] = "Error Resolving ($GroupId)"
                        }
                    }
                }
                
                # Handle special target types
                $TargetDisplayName = switch ($Assignment.target.'@odata.type') {
                    '#microsoft.graph.allLicensedUsersAssignmentTarget' { 'All Users' }
                    '#microsoft.graph.allDevicesAssignmentTarget' { 'All Devices' }
                    '#microsoft.graph.exclusionGroupAssignmentTarget' { "Exclude: $($GroupDisplayNames -join ', ')" }
                    '#microsoft.graph.groupAssignmentTarget' { $GroupDisplayNames -join ', ' }
                    default { 
                        if ($GroupDisplayNames.Count -gt 0) { 
                            $GroupDisplayNames -join ', ' 
                        } else { 
                            'Unknown Target' 
                        }
                    }
                }
                
                # Add resolved information to assignment
                $Assignment | Add-Member -MemberType NoteProperty -Name 'GroupDisplayNames' -Value $GroupDisplayNames -Force
                $Assignment | Add-Member -MemberType NoteProperty -Name 'TargetDisplayName' -Value $TargetDisplayName -Force
            }
            
            return $Assignments
        }
        catch {
            Write-Error "Failed to resolve group assignments: $($_.Exception.Message)"
            return $Assignments
        }
    }
    
    end {
        Write-Verbose "Completed group ID resolution. Cache contains $($Script:GroupCache.Count) entries"
    }
}
