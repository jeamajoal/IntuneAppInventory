function Get-InventorySummaryReport {
    <#
    .SYNOPSIS
    Generates a summary report of the inventory database.
    #>
    [CmdletBinding()]
    param()
    
    try {
        $Summary = @{}
        
        # Get application counts
        $AppCommand = $Script:DatabaseConnection.CreateCommand()
        $AppCommand.CommandText = @"
SELECT 
    COUNT(*) as TotalApps,
    SUM(CASE WHEN HasSourceCode = 1 THEN 1 ELSE 0 END) as AppsWithSourceCode,
    COUNT(DISTINCT AppType) as UniqueAppTypes
FROM Applications;
"@
        $AppReader = $AppCommand.ExecuteReader()
        if ($AppReader.Read()) {
            $Summary.Applications = @{
                Total = $AppReader["TotalApps"]
                WithSourceCode = $AppReader["AppsWithSourceCode"]
                UniqueTypes = $AppReader["UniqueAppTypes"]
            }
        }
        $AppReader.Close()
        $AppCommand.Dispose()
        
        # Get script counts
        $ScriptCommand = $Script:DatabaseConnection.CreateCommand()
        $ScriptCommand.CommandText = @"
SELECT 
    COUNT(*) as TotalScripts,
    SUM(CASE WHEN HasSourceCode = 1 THEN 1 ELSE 0 END) as ScriptsWithSourceCode
FROM Scripts;
"@
        $ScriptReader = $ScriptCommand.ExecuteReader()
        if ($ScriptReader.Read()) {
            $Summary.Scripts = @{
                Total = $ScriptReader["TotalScripts"]
                WithSourceCode = $ScriptReader["ScriptsWithSourceCode"]
            }
        }
        $ScriptReader.Close()
        $ScriptCommand.Dispose()
        
        # Get remediation counts
        $RemediationCommand = $Script:DatabaseConnection.CreateCommand()
        $RemediationCommand.CommandText = @"
SELECT 
    COUNT(*) as TotalRemediations,
    SUM(CASE WHEN HasSourceCode = 1 THEN 1 ELSE 0 END) as RemediationsWithSourceCode
FROM Remediations;
"@
        $RemediationReader = $RemediationCommand.ExecuteReader()
        if ($RemediationReader.Read()) {
            $Summary.Remediations = @{
                Total = $RemediationReader["TotalRemediations"]
                WithSourceCode = $RemediationReader["RemediationsWithSourceCode"]
            }
        }
        $RemediationReader.Close()
        $RemediationCommand.Dispose()
        
        # Get inventory run information
        $RunCommand = $Script:DatabaseConnection.CreateCommand()
        $RunCommand.CommandText = @"
SELECT 
    COUNT(*) as TotalRuns,
    MAX(StartTime) as LastRun,
    SUM(CASE WHEN Status = 'Completed' THEN 1 ELSE 0 END) as SuccessfulRuns
FROM InventoryRuns;
"@
        $RunReader = $RunCommand.ExecuteReader()
        if ($RunReader.Read()) {
            $Summary.InventoryRuns = @{
                Total = $RunReader["TotalRuns"]
                LastRun = $RunReader["LastRun"]
                Successful = $RunReader["SuccessfulRuns"]
            }
        }
        $RunReader.Close()
        $RunCommand.Dispose()
        
        return [PSCustomObject]$Summary
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to generate summary report: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Get-ApplicationsReport {
    <#
    .SYNOPSIS
    Generates a detailed applications report.
    #>
    [CmdletBinding()]
    param(
        [switch]$IncludeSourceCode,
        [switch]$FilterMissingSourceCode
    )
    
    try {
        $Query = @"
SELECT 
    Id, DisplayName, Description, Publisher, AppType, 
    CreatedDateTime, LastModifiedDateTime, PublishingState,
    HasSourceCode, SourceCodeAdded, InventoryDate
FROM Applications
"@
        
        if ($FilterMissingSourceCode) {
            $Query += " WHERE HasSourceCode = 0"
        }
        
        $Query += " ORDER BY DisplayName;"
        
        $Command = $Script:DatabaseConnection.CreateCommand()
        $Command.CommandText = $Query
        
        $Reader = $Command.ExecuteReader()
        $Applications = @()
        while ($Reader.Read()) {
            $App = @{
                Id = $Reader["Id"]
                DisplayName = $Reader["DisplayName"]
                Description = $Reader["Description"]
                Publisher = $Reader["Publisher"]
                AppType = $Reader["AppType"]
                CreatedDateTime = $Reader["CreatedDateTime"]
                LastModifiedDateTime = $Reader["LastModifiedDateTime"]
                PublishingState = $Reader["PublishingState"]
                HasSourceCode = [bool]$Reader["HasSourceCode"]
                SourceCodeAdded = $Reader["SourceCodeAdded"]
                InventoryDate = $Reader["InventoryDate"]
            }
            
            if ($IncludeSourceCode -and $App.HasSourceCode) {
                $App.SourceCode = $Reader["SourceCode"]
            }
            
            $Applications += [PSCustomObject]$App
        }
        $Reader.Close()
        $Command.Dispose()
        
        return $Applications
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to generate applications report: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Get-ScriptsReport {
    <#
    .SYNOPSIS
    Generates a detailed scripts report.
    #>
    [CmdletBinding()]
    param(
        [switch]$IncludeSourceCode,
        [switch]$FilterMissingSourceCode
    )
    
    try {
        $SelectFields = @"
Id, DisplayName, Description, CreatedDateTime, LastModifiedDateTime,
RunAsAccount, FileName, HasSourceCode, SourceCodeAdded, InventoryDate
"@
        
        if ($IncludeSourceCode) {
            $SelectFields += ", SourceCode"
        }
        
        $Query = "SELECT $SelectFields FROM Scripts"
        
        if ($FilterMissingSourceCode) {
            $Query += " WHERE HasSourceCode = 0"
        }
        
        $Query += " ORDER BY DisplayName;"
        
        $Command = $Script:DatabaseConnection.CreateCommand()
        $Command.CommandText = $Query
        
        $Reader = $Command.ExecuteReader()
        $Scripts = @()
        while ($Reader.Read()) {
            $Script = @{
                Id = $Reader["Id"]
                DisplayName = $Reader["DisplayName"]
                Description = $Reader["Description"]
                CreatedDateTime = $Reader["CreatedDateTime"]
                LastModifiedDateTime = $Reader["LastModifiedDateTime"]
                RunAsAccount = $Reader["RunAsAccount"]
                FileName = $Reader["FileName"]
                HasSourceCode = [bool]$Reader["HasSourceCode"]
                SourceCodeAdded = $Reader["SourceCodeAdded"]
                InventoryDate = $Reader["InventoryDate"]
            }
            
            if ($IncludeSourceCode -and $Script.HasSourceCode) {
                $Script.SourceCode = $Reader["SourceCode"]
            }
            
            $Scripts += [PSCustomObject]$Script
        }
        $Reader.Close()
        $Command.Dispose()
        
        return $Scripts
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to generate scripts report: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Get-RemediationsReport {
    <#
    .SYNOPSIS
    Generates a detailed remediations report.
    #>
    [CmdletBinding()]
    param(
        [switch]$IncludeSourceCode,
        [switch]$FilterMissingSourceCode
    )
    
    try {
        $SelectFields = @"
Id, DisplayName, Description, Publisher, Version, CreatedDateTime, 
LastModifiedDateTime, RunAsAccount, HasSourceCode, SourceCodeAdded, InventoryDate
"@
        
        if ($IncludeSourceCode) {
            $SelectFields += ", SourceCode"
        }
        
        $Query = "SELECT $SelectFields FROM Remediations"
        
        if ($FilterMissingSourceCode) {
            $Query += " WHERE HasSourceCode = 0"
        }
        
        $Query += " ORDER BY DisplayName;"
        
        $Command = $Script:DatabaseConnection.CreateCommand()
        $Command.CommandText = $Query
        
        $Reader = $Command.ExecuteReader()
        $Remediations = @()
        while ($Reader.Read()) {
            $Remediation = @{
                Id = $Reader["Id"]
                DisplayName = $Reader["DisplayName"]
                Description = $Reader["Description"]
                Publisher = $Reader["Publisher"]
                Version = $Reader["Version"]
                CreatedDateTime = $Reader["CreatedDateTime"]
                LastModifiedDateTime = $Reader["LastModifiedDateTime"]
                RunAsAccount = $Reader["RunAsAccount"]
                HasSourceCode = [bool]$Reader["HasSourceCode"]
                SourceCodeAdded = $Reader["SourceCodeAdded"]
                InventoryDate = $Reader["InventoryDate"]
            }
            
            if ($IncludeSourceCode -and $Remediation.HasSourceCode) {
                $Remediation.SourceCode = $Reader["SourceCode"]
            }
            
            $Remediations += [PSCustomObject]$Remediation
        }
        $Reader.Close()
        $Command.Dispose()
        
        return $Remediations
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to generate remediations report: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Get-AssignmentsReport {
    <#
    .SYNOPSIS
    Generates a detailed assignments report.
    #>
    [CmdletBinding()]
    param()
    
    try {
        $Query = @"
SELECT 
    Id, ItemId, ItemType, TargetType, TargetGroupId, TargetGroupName,
    Intent, CreatedDateTime, LastModifiedDateTime, InventoryDate
FROM Assignments
ORDER BY ItemType, ItemId;
"@
        
        $Command = $Script:DatabaseConnection.CreateCommand()
        $Command.CommandText = $Query
        
        $Reader = $Command.ExecuteReader()
        $Assignments = @()
        while ($Reader.Read()) {
            $Assignment = @{
                Id = $Reader["Id"]
                ItemId = $Reader["ItemId"]
                ItemType = $Reader["ItemType"]
                TargetType = $Reader["TargetType"]
                TargetGroupId = $Reader["TargetGroupId"]
                TargetGroupName = $Reader["TargetGroupName"]
                Intent = $Reader["Intent"]
                CreatedDateTime = $Reader["CreatedDateTime"]
                LastModifiedDateTime = $Reader["LastModifiedDateTime"]
                InventoryDate = $Reader["InventoryDate"]
            }
            
            $Assignments += [PSCustomObject]$Assignment
        }
        $Reader.Close()
        $Command.Dispose()
        
        return $Assignments
    }
    catch {
        Write-IntuneInventoryLog -Message "Failed to generate assignments report: $($_.Exception.Message)" -Level Error
        throw
    }
}

function New-HtmlReport {
    <#
    .SYNOPSIS
    Generates an HTML report from the inventory data.
    #>
    [CmdletBinding()]
    param(
        [hashtable]$ReportData,
        [string]$OutputPath
    )
    
    $Html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Intune Inventory Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1, h2 { color: #0078d4; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .summary { background-color: #f9f9f9; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .missing-source { color: #d83b01; font-weight: bold; }
        .has-source { color: #107c10; }
    </style>
</head>
<body>
    <h1>Intune Inventory Report</h1>
    <p>Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    
    <div class="summary">
        <h2>Executive Summary</h2>
        <ul>
            <li>Total Applications: $($ReportData.Summary.Applications.Total) (With Source Code: $($ReportData.Summary.Applications.WithSourceCode))</li>
            <li>Total Scripts: $($ReportData.Summary.Scripts.Total) (With Source Code: $($ReportData.Summary.Scripts.WithSourceCode))</li>
            <li>Total Remediations: $($ReportData.Summary.Remediations.Total) (With Source Code: $($ReportData.Summary.Remediations.WithSourceCode))</li>
            <li>Last Inventory Run: $($ReportData.Summary.InventoryRuns.LastRun)</li>
        </ul>
    </div>
    
    <!-- Additional HTML content would be generated here -->
    
</body>
</html>
"@
    
    $Html | Out-File -FilePath $OutputPath -Encoding UTF8
}

function New-ExecutiveSummary {
    <#
    .SYNOPSIS
    Generates an executive summary text report.
    #>
    [CmdletBinding()]
    param(
        [hashtable]$ReportData,
        [string]$OutputPath
    )
    
    $Summary = @"
INTUNE INVENTORY EXECUTIVE SUMMARY
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
=====================================

OVERVIEW
--------
Total Applications: $($ReportData.Summary.Applications.Total)
Total Scripts: $($ReportData.Summary.Scripts.Total)
Total Remediations: $($ReportData.Summary.Remediations.Total)

SOURCE CODE COVERAGE
-------------------
Applications with Source Code: $($ReportData.Summary.Applications.WithSourceCode) / $($ReportData.Summary.Applications.Total) ($([math]::Round(($ReportData.Summary.Applications.WithSourceCode / [math]::Max($ReportData.Summary.Applications.Total, 1)) * 100, 1))%)
Scripts with Source Code: $($ReportData.Summary.Scripts.WithSourceCode) / $($ReportData.Summary.Scripts.Total) ($([math]::Round(($ReportData.Summary.Scripts.WithSourceCode / [math]::Max($ReportData.Summary.Scripts.Total, 1)) * 100, 1))%)
Remediations with Source Code: $($ReportData.Summary.Remediations.WithSourceCode) / $($ReportData.Summary.Remediations.Total) ($([math]::Round(($ReportData.Summary.Remediations.WithSourceCode / [math]::Max($ReportData.Summary.Remediations.Total, 1)) * 100, 1))%)

RECOMMENDATIONS
--------------
- Review items missing source code and add where possible
- Establish process for maintaining source code with deployments
- Regular inventory updates to maintain currency

INVENTORY HISTORY
----------------
Total Inventory Runs: $($ReportData.Summary.InventoryRuns.Total)
Successful Runs: $($ReportData.Summary.InventoryRuns.Successful)
Last Run: $($ReportData.Summary.InventoryRuns.LastRun)
"@
    
    $Summary | Out-File -FilePath $OutputPath -Encoding UTF8
}
