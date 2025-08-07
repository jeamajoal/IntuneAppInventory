function Initialize-DatabaseSchema {
    <#
    .SYNOPSIS
    Initializes the SQLite database schema for Intune inventory.
    
    .DESCRIPTION
    Creates the necessary tables and indexes for storing Intune applications, scripts, 
    remediations, assignments, and source code information.
    
    .PARAMETER DatabaseConnection
    The SQLite database connection object.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SQLite.SQLiteConnection]$DatabaseConnection
    )
    
    try {
        Write-Verbose "Creating database schema..."
        
        $SchemaQueries = @(
            # Applications table
            @"
CREATE TABLE IF NOT EXISTS Applications (
    Id TEXT PRIMARY KEY,
    DisplayName TEXT NOT NULL,
    Description TEXT,
    Publisher TEXT,
    AppType TEXT,
    CreatedDateTime TEXT,
    LastModifiedDateTime TEXT,
    PrivacyInformationUrl TEXT,
    InformationUrl TEXT,
    Owner TEXT,
    Developer TEXT,
    Notes TEXT,
    PublishingState TEXT,
    CommittedContentVersion TEXT,
    FileName TEXT,
    Size INTEGER,
    InstallCommandLine TEXT,
    UninstallCommandLine TEXT,
    MinimumSupportedOperatingSystem TEXT,
    HasSourceCode INTEGER DEFAULT 0,
    SourceCode TEXT,
    SourceCodeAdded TEXT,
    InventoryDate TEXT NOT NULL,
    LastUpdated TEXT NOT NULL
);
"@,
            # Scripts table
            @"
CREATE TABLE IF NOT EXISTS Scripts (
    Id TEXT PRIMARY KEY,
    DisplayName TEXT NOT NULL,
    Description TEXT,
    ScriptContent TEXT,
    CreatedDateTime TEXT,
    LastModifiedDateTime TEXT,
    RunAsAccount TEXT,
    FileName TEXT,
    RoleScopeTagIds TEXT,
    HasSourceCode INTEGER DEFAULT 1,
    SourceCode TEXT,
    SourceCodeAdded TEXT,
    InventoryDate TEXT NOT NULL,
    LastUpdated TEXT NOT NULL
);
"@,
            # Remediations table
            @"
CREATE TABLE IF NOT EXISTS Remediations (
    Id TEXT PRIMARY KEY,
    DisplayName TEXT NOT NULL,
    Description TEXT,
    Publisher TEXT,
    Version TEXT,
    CreatedDateTime TEXT,
    LastModifiedDateTime TEXT,
    DetectionScriptContent TEXT,
    RemediationScriptContent TEXT,
    RunAsAccount TEXT,
    RoleScopeTagIds TEXT,
    HasSourceCode INTEGER DEFAULT 1,
    SourceCode TEXT,
    SourceCodeAdded TEXT,
    InventoryDate TEXT NOT NULL,
    LastUpdated TEXT NOT NULL
);
"@,
            # Assignments table
            @"
CREATE TABLE IF NOT EXISTS Assignments (
    Id TEXT PRIMARY KEY,
    ItemId TEXT NOT NULL,
    ItemType TEXT NOT NULL,
    TargetType TEXT,
    TargetGroupId TEXT,
    TargetGroupName TEXT,
    Intent TEXT,
    Settings TEXT,
    CreatedDateTime TEXT,
    LastModifiedDateTime TEXT,
    InventoryDate TEXT NOT NULL,
    FOREIGN KEY (ItemId) REFERENCES Applications(Id) ON DELETE CASCADE,
    FOREIGN KEY (ItemId) REFERENCES Scripts(Id) ON DELETE CASCADE,
    FOREIGN KEY (ItemId) REFERENCES Remediations(Id) ON DELETE CASCADE
);
"@,
            # Inventory runs table
            @"
CREATE TABLE IF NOT EXISTS InventoryRuns (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    RunType TEXT NOT NULL,
    StartTime TEXT NOT NULL,
    EndTime TEXT,
    Status TEXT NOT NULL,
    ItemsProcessed INTEGER DEFAULT 0,
    ErrorCount INTEGER DEFAULT 0,
    ErrorMessages TEXT,
    TenantId TEXT,
    UserPrincipalName TEXT
);
"@,
            # Source code history table
            @"
CREATE TABLE IF NOT EXISTS SourceCodeHistory (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    ItemId TEXT NOT NULL,
    ItemType TEXT NOT NULL,
    SourceCode TEXT NOT NULL,
    AddedBy TEXT,
    AddedDate TEXT NOT NULL,
    Comments TEXT,
    Version TEXT
);
"@
        )
        
        # Index creation queries
        $IndexQueries = @(
            "CREATE INDEX IF NOT EXISTS idx_applications_displayname ON Applications(DisplayName);",
            "CREATE INDEX IF NOT EXISTS idx_applications_apptype ON Applications(AppType);",
            "CREATE INDEX IF NOT EXISTS idx_applications_inventorydate ON Applications(InventoryDate);",
            "CREATE INDEX IF NOT EXISTS idx_scripts_displayname ON Scripts(DisplayName);",
            "CREATE INDEX IF NOT EXISTS idx_scripts_inventorydate ON Scripts(InventoryDate);",
            "CREATE INDEX IF NOT EXISTS idx_remediations_displayname ON Remediations(DisplayName);",
            "CREATE INDEX IF NOT EXISTS idx_remediations_inventorydate ON Remediations(InventoryDate);",
            "CREATE INDEX IF NOT EXISTS idx_assignments_itemid ON Assignments(ItemId);",
            "CREATE INDEX IF NOT EXISTS idx_assignments_itemtype ON Assignments(ItemType);",
            "CREATE INDEX IF NOT EXISTS idx_assignments_targettype ON Assignments(TargetType);",
            "CREATE INDEX IF NOT EXISTS idx_inventoryruns_runtype ON InventoryRuns(RunType);",
            "CREATE INDEX IF NOT EXISTS idx_inventoryruns_starttime ON InventoryRuns(StartTime);",
            "CREATE INDEX IF NOT EXISTS idx_sourcecodehist_itemid ON SourceCodeHistory(ItemId);",
            "CREATE INDEX IF NOT EXISTS idx_sourcecodehist_itemtype ON SourceCodeHistory(ItemType);"
        )
        
        # Execute schema creation
        foreach ($Query in $SchemaQueries) {
            $Command = $DatabaseConnection.CreateCommand()
            $Command.CommandText = $Query
            $null = $Command.ExecuteNonQuery()
            $Command.Dispose()
        }
        
        # Execute index creation
        foreach ($Query in $IndexQueries) {
            $Command = $DatabaseConnection.CreateCommand()
            $Command.CommandText = $Query
            $null = $Command.ExecuteNonQuery()
            $Command.Dispose()
        }
        
        Write-Verbose "Database schema initialized successfully"
    }
    catch {
        Write-Error "Failed to initialize database schema: $($_.Exception.Message)"
        throw
    }
}

function Get-DatabaseConnection {
    <#
    .SYNOPSIS
    Creates and returns a SQLite database connection.
    
    .DESCRIPTION
    Establishes a connection to the SQLite database file and returns the connection object.
    
    .PARAMETER DatabasePath
    The full path to the SQLite database file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DatabasePath
    )
    
    try {
        Write-Verbose "Connecting to database: $DatabasePath"
        
        # Ensure directory exists
        $DatabaseDirectory = Split-Path -Path $DatabasePath -Parent
        if (-not (Test-Path -Path $DatabaseDirectory)) {
            New-Item -Path $DatabaseDirectory -ItemType Directory -Force | Out-Null
        }
        
        # Create connection string
        $ConnectionString = "Data Source=$DatabasePath;Version=3;Journal Mode=WAL;Synchronous=NORMAL;"
        
        # Create and open connection
        $Connection = New-Object System.Data.SQLite.SQLiteConnection($ConnectionString)
        $Connection.Open()
        
        Write-Verbose "Database connection established successfully"
        return $Connection
    }
    catch {
        Write-Error "Failed to connect to database: $($_.Exception.Message)"
        throw
    }
}

function Test-DatabaseConnection {
    <#
    .SYNOPSIS
    Tests if the database connection is valid and responsive.
    
    .DESCRIPTION
    Performs a simple query against the database to verify connectivity.
    
    .PARAMETER DatabaseConnection
    The SQLite database connection object to test.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SQLite.SQLiteConnection]$DatabaseConnection
    )
    
    try {
        if ($DatabaseConnection.State -ne 'Open') {
            return $false
        }
        
        $Command = $DatabaseConnection.CreateCommand()
        $Command.CommandText = "SELECT 1;"
        $Result = $Command.ExecuteScalar()
        $Command.Dispose()
        
        return ($Result -eq 1)
    }
    catch {
        Write-Verbose "Database connection test failed: $($_.Exception.Message)"
        return $false
    }
}

function Close-DatabaseConnection {
    <#
    .SYNOPSIS
    Safely closes a SQLite database connection.
    
    .DESCRIPTION
    Properly closes and disposes of a SQLite database connection object.
    
    .PARAMETER DatabaseConnection
    The SQLite database connection object to close.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.SQLite.SQLiteConnection]$DatabaseConnection
    )
    
    try {
        if ($DatabaseConnection -and $DatabaseConnection.State -eq 'Open') {
            $DatabaseConnection.Close()
            $DatabaseConnection.Dispose()
            Write-Verbose "Database connection closed successfully"
        }
    }
    catch {
        Write-Warning "Error closing database connection: $($_.Exception.Message)"
    }
}
