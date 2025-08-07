# Module-scoped variables for Graph API
$Script:GraphToken = $null
$Script:TokenExpiry = $null
$Script:GraphHeaders = @{}
$Script:GraphBaseUrl = "https://graph.microsoft.com/v1.0"
$Script:TokenCache = @{}
$Script:ConnectionInfo = $null

function Get-GraphAccessToken {
    <#
    .SYNOPSIS
    Acquires an access token for Microsoft Graph API.
    
    .DESCRIPTION
    Gets an access token using client credentials flow for Microsoft Graph API access.
    Supports both client secret and certificate-based authentication.
    
    .PARAMETER TenantId
    The Azure AD tenant ID.
    
    .PARAMETER ClientId
    The application (client) ID.
    
    .PARAMETER ClientSecret
    The client secret for authentication.
    
    .PARAMETER CertificateThumbprint
    The certificate thumbprint for certificate-based authentication.
    
    .PARAMETER ForceRefresh
    Force acquisition of a new token even if cached token is valid.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        
        [Parameter(ParameterSetName = 'ClientSecret', Mandatory = $true)]
        [SecureString]$ClientSecret,
        
        [Parameter(ParameterSetName = 'Certificate')]
        [string]$CertificateThumbprint,
        
        [Parameter()]
        [switch]$ForceRefresh
    )
    
    begin {
        Write-Verbose "Get-GraphAccessToken: Starting token acquisition"
        
        # Check cache unless force refresh
        if (-not $ForceRefresh) {
            $CacheKey = "$TenantId-$ClientId"
            if ($Script:TokenCache.ContainsKey($CacheKey)) {
                $Cached = $Script:TokenCache[$CacheKey]
                if ($Cached.Expiry -gt (Get-Date).AddMinutes(5)) {
                    Write-Verbose "Using cached token (expires: $($Cached.Expiry))"
                    $Script:GraphToken = $Cached.Token
                    $Script:TokenExpiry = $Cached.Expiry
                    $Script:GraphHeaders = @{
                        'Authorization' = "Bearer $($Cached.Token)"
                        'Content-Type' = 'application/json'
                        'Accept' = 'application/json'
                    }
                    return @{
                        TokenAcquired = $true
                        Source = "Cache"
                        ExpiresAt = $Cached.Expiry
                    }
                }
            }
        }
        
        $TokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    }
    
    process {
        try {
            # Build request body
            $Body = @{
                client_id = $ClientId
                scope = "https://graph.microsoft.com/.default"
                grant_type = "client_credentials"
            }
            
            # Handle authentication method
            if ($PSCmdlet.ParameterSetName -eq 'ClientSecret') {
                # Convert SecureString to plain text
                $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
                try {
                    $PlainSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
                    $Body.client_secret = $PlainSecret
                }
                finally {
                    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
                    if ($PlainSecret) { Clear-Variable PlainSecret -Force }
                }
            }
            elseif ($PSCmdlet.ParameterSetName -eq 'Certificate') {
                throw "Certificate authentication not yet implemented"
            }
            
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
            
            # Cache token
            $CacheKey = "$TenantId-$ClientId"
            $Script:TokenCache[$CacheKey] = @{
                Token = $Response.access_token
                Expiry = $ExpiryTime
                TenantId = $TenantId
                ClientId = $ClientId
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
            if ($Body -and $Body.client_secret) {
                $Body.client_secret = $null
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

function Invoke-GraphRequest {
    <#
    .SYNOPSIS
    Makes a request to Microsoft Graph API.
    
    .DESCRIPTION
    Wrapper function for making HTTP requests to Microsoft Graph API with proper error handling
    and automatic pagination support.
    
    .PARAMETER Uri
    The Graph API endpoint URI.
    
    .PARAMETER Method
    The HTTP method (GET, POST, PUT, DELETE, PATCH).
    
    .PARAMETER Body
    The request body for POST/PUT/PATCH requests.
    
    .PARAMETER All
    Automatically handle pagination and return all results.
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
        # Ensure we have a valid token
        if (-not (Test-GraphToken)) {
            throw "No valid Graph token available. Please authenticate first."
        }
        
        # Prepare request parameters
        $RequestParams = @{
            Uri = $Uri
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
        
        # Make initial request
        $Response = Invoke-RestMethod @RequestParams
        
        # Handle pagination if All switch is specified
        if ($All -and $Response.'@odata.nextLink') {
            $AllResults = @()
            if ($Response.value) {
                $AllResults += $Response.value
            }
            else {
                $AllResults += $Response
            }
            
            $NextLink = $Response.'@odata.nextLink'
            while ($NextLink) {
                Write-Verbose "Following pagination link: $NextLink"
                $RequestParams.Uri = $NextLink
                $PageResponse = Invoke-RestMethod @RequestParams
                
                if ($PageResponse.value) {
                    $AllResults += $PageResponse.value
                }
                
                $NextLink = $PageResponse.'@odata.nextLink'
            }
            
            return $AllResults
        }
        
        return $Response
    }
    catch {
        $ErrorMessage = "Graph API request failed: $($_.Exception.Message)"
        
        # Try to extract more detailed error information
        if ($_.Exception.Response) {
            try {
                $ErrorStream = $_.Exception.Response.GetResponseStream()
                $Reader = New-Object System.IO.StreamReader($ErrorStream)
                $ErrorBody = $Reader.ReadToEnd() | ConvertFrom-Json
                
                if ($ErrorBody.error) {
                    $ErrorMessage += " - $($ErrorBody.error.code): $($ErrorBody.error.message)"
                }
            }
            catch {
                # If we can't parse the error response, use the original message
            }
        }
        
        Write-Error $ErrorMessage
        throw
    }
}
    <#
    .SYNOPSIS
    Writes log messages for the Intune Inventory module.
    
    .DESCRIPTION
    Centralized logging function that writes messages to various outputs based on configuration.
    This is a placeholder that will be customized based on user-provided logging examples.
    
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
    
    switch ($Level) {
        'Error' { Write-Error $LogMessage }
        'Warning' { Write-Warning $LogMessage }
        'Debug' { Write-Debug $LogMessage }
        'Verbose' { Write-Verbose $LogMessage }
        default { Write-Host $LogMessage }
    }
}

function Test-GraphAuthentication {
    <#
    .SYNOPSIS
    Tests if Microsoft Graph authentication is valid and has required permissions.
    
    .DESCRIPTION
    Verifies that the current Microsoft Graph session has the necessary permissions
    for Intune inventory operations.
    #>
    [CmdletBinding()]
    param()
    
    try {
        $Context = Get-MgContext
        if (-not $Context) {
            Write-IntuneInventoryLog -Message "No Microsoft Graph authentication context found" -Level Error
            return $false
        }
        
        # Test basic connectivity
        try {
            $null = Get-MgOrganization -ErrorAction Stop
            Write-IntuneInventoryLog -Message "Microsoft Graph connectivity verified" -Level Verbose
        }
        catch {
    function Write-IntuneInventoryLog {
    <#
    .SYNOPSIS
    Writes log messages for the Intune Inventory module.
    
    .DESCRIPTION
    Centralized logging function that writes messages to various outputs based on configuration.
    This is a placeholder that will be customized based on user-provided logging examples.
    
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
    
    switch ($Level) {
        'Error' { Write-Error $LogMessage }
        'Warning' { Write-Warning $LogMessage }
        'Debug' { Write-Debug $LogMessage }
        'Verbose' { Write-Verbose $LogMessage }
        default { Write-Host $LogMessage }
    }
}

function Test-GraphAuthentication {
    <#
    .SYNOPSIS
    Tests if Microsoft Graph authentication is valid and has required permissions.
    
    .DESCRIPTION
    Verifies that the current Microsoft Graph session has the necessary permissions
    for Intune inventory operations using pure Graph API calls.
    #>
    [CmdletBinding()]
    param()
    
    try {
        if (-not (Test-GraphToken)) {
            Write-IntuneInventoryLog -Message "No valid Microsoft Graph token found" -Level Error
            return $false
        }
        
        # Test basic connectivity with organization endpoint
        try {
            $OrgResponse = Invoke-GraphRequest -Uri "$Script:GraphBaseUrl/organization" -Method GET
            Write-IntuneInventoryLog -Message "Microsoft Graph connectivity verified" -Level Verbose
        }
        catch {
            Write-IntuneInventoryLog -Message "Microsoft Graph connectivity test failed: $($_.Exception.Message)" -Level Error
            return $false
        }
        
        # Test Intune-specific endpoints to verify permissions
        $TestEndpoints = @(
            @{ Endpoint = "deviceAppManagement/mobileApps?`$top=1"; Description = "Mobile Apps" },
            @{ Endpoint = "deviceManagement/deviceManagementScripts?`$top=1"; Description = "Device Management Scripts" },
            @{ Endpoint = "deviceManagement/deviceHealthScripts?`$top=1"; Description = "Device Health Scripts" }
        )
        
        $PermissionErrors = @()
        foreach ($Test in $TestEndpoints) {
            try {
                $null = Invoke-GraphRequest -Uri "$Script:GraphBaseUrl/$($Test.Endpoint)" -Method GET
                Write-IntuneInventoryLog -Message "Permission verified for $($Test.Description)" -Level Verbose
            }
            catch {
                $PermissionErrors += "$($Test.Description): $($_.Exception.Message)"
            }
        }
        
        if ($PermissionErrors.Count -gt 0) {
            Write-IntuneInventoryLog -Message "Permission verification failed for: $($PermissionErrors -join '; ')" -Level Warning
            return $false
        }
        
        Write-IntuneInventoryLog -Message "Microsoft Graph authentication and permissions verified" -Level Info
        return $true
    }
    catch {
        Write-IntuneInventoryLog -Message "Authentication test failed: $($_.Exception.Message)" -Level Error
        return $false
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
            return @{
                UserPrincipalName = $env:USERNAME
                TenantId = $null
                ClientId = $null
                AuthMethod = "Unknown"
            }
        }
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to get user context: $($_.Exception.Message)" -Level Warning
        return @{
            UserPrincipalName = $env:USERNAME
            TenantId = $null
            ClientId = $null
            AuthMethod = "Unknown"
        }
    }
}
