function Get-IntuneInventoryReport {
    <#
    .SYNOPSIS
    Generates comprehensive reports from the Intune inventory data stored in JSON.
    
    .DESCRIPTION
    Creates various types of reports from the JSON-stored inventory data including
    applications, scripts, remediations, assignments, and summary information.
    
    .PARAMETER ReportType
    The type of report to generate: Summary, Applications, Scripts, Remediations, Assignments, or All.
    
    .PARAMETER OutputFormat
    The output format: Object (default), CSV, JSON, or HTML.
    
    .PARAMETER OutputPath
    Optional path to save the report. If not specified, returns the report object.
    
    .PARAMETER FilterBy
    Optional filter criteria for the report data.
    
    .EXAMPLE
    Get-IntuneInventoryReport -ReportType Summary
    
    .EXAMPLE
    Get-IntuneInventoryReport -ReportType Applications -OutputFormat CSV -OutputPath "C:\Reports\Apps.csv"
    
    .EXAMPLE
    Get-IntuneInventoryReport -ReportType All -OutputFormat HTML -OutputPath "C:\Reports\Complete.html"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Summary', 'Applications', 'Scripts', 'Remediations', 'Assignments', 'InventoryRuns', 'All')]
        [string]$ReportType,
        
        [Parameter()]
        [ValidateSet('Object', 'CSV', 'JSON', 'HTML')]
        [string]$OutputFormat = 'Object',
        
        [Parameter()]
        [string]$OutputPath,
        
        [Parameter()]
        [hashtable]$FilterBy
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting inventory report generation: $ReportType" -Level Info -Source "Get-IntuneInventoryReport"
        
        # Verify storage connection
        if (-not (Test-StorageConnection)) {
            throw "JSON storage connection is not available. Please run Connect-IntuneInventory first."
        }
    }
    
    process {
        try {
            $ReportData = @{}
            $ReportMetadata = @{
                GeneratedAt = Get-Date
                GeneratedBy = $env:USERNAME
                ReportType = $ReportType
                StorageLocation = $Script:StorageRoot
                TenantId = if ($Script:ConnectionInfo) { $Script:ConnectionInfo.TenantId } else { "Unknown" }
                TenantName = if ($Script:ConnectionInfo) { $Script:ConnectionInfo.TenantName } else { "Unknown" }
            }
            
            switch ($ReportType) {
                'Summary' {
                    $ReportData = Get-InventorySummaryReport -FilterBy $FilterBy
                }
                'Applications' {
                    $ReportData = Get-ApplicationsReport -FilterBy $FilterBy
                }
                'Scripts' {
                    $ReportData = Get-ScriptsReport -FilterBy $FilterBy
                }
                'Remediations' {
                    $ReportData = Get-RemediationsReport -FilterBy $FilterBy
                }
                'Assignments' {
                    $ReportData = Get-AssignmentsReport -FilterBy $FilterBy
                }
                'InventoryRuns' {
                    $ReportData = Get-InventoryRunsReport -FilterBy $FilterBy
                }
                'All' {
                    $ReportData = @{
                        Summary = Get-InventorySummaryReport -FilterBy $FilterBy
                        Applications = Get-ApplicationsReport -FilterBy $FilterBy
                        Scripts = Get-ScriptsReport -FilterBy $FilterBy
                        Remediations = Get-RemediationsReport -FilterBy $FilterBy
                        Assignments = Get-AssignmentsReport -FilterBy $FilterBy
                        InventoryRuns = Get-InventoryRunsReport -FilterBy $FilterBy
                    }
                }
            }
            
            # Add metadata to report
            $FullReport = @{
                Metadata = $ReportMetadata
                Data = $ReportData
            }
            
            # Handle output format and path
            if ($OutputPath) {
                $null = Save-ReportToFile -Report $FullReport -Format $OutputFormat -Path $OutputPath
                Write-IntuneInventoryLog -Message "Report saved to: $OutputPath" -Level Info
                return @{ Success = $true; ReportPath = $OutputPath; Format = $OutputFormat }
            }
            else {
                if ($OutputFormat -eq 'Object') { 
                    return $FullReport 
                } else { 
                    return Format-ReportOutput -Report $FullReport -Format $OutputFormat 
                }
            }
        }
        catch {
            $ErrorMessage = "Failed to generate report: $($_.Exception.Message)"
            Write-IntuneInventoryLog -Message $ErrorMessage -Level Error
            throw $ErrorMessage
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Report generation completed" -Level Info -Source "Get-IntuneInventoryReport"
    }
}

function Export-IntuneInventoryReport {
    <#
    .SYNOPSIS
    Exports Intune inventory reports to various formats.
    
    .DESCRIPTION
    Generates and exports comprehensive Intune inventory reports to files in various formats.
    
    .PARAMETER ReportType
    The type of report to export.
    
    .PARAMETER OutputFormat
    The output format for the exported report.
    
    .PARAMETER OutputPath
    The path where the report will be saved.
    
    .PARAMETER IncludeMetadata
    Include metadata information in the exported report.
    
    .EXAMPLE
    Export-IntuneInventoryReport -ReportType Applications -OutputFormat CSV -OutputPath "C:\Reports\Applications.csv"
    
    .EXAMPLE
    Export-IntuneInventoryReport -ReportType All -OutputFormat HTML -OutputPath "C:\Reports\Complete.html" -IncludeMetadata
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Summary', 'Applications', 'Scripts', 'Remediations', 'Assignments', 'InventoryRuns', 'All')]
        [string]$ReportType,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('CSV', 'JSON', 'HTML', 'XML')]
        [string]$OutputFormat,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter()]
        [switch]$IncludeMetadata
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting report export: $ReportType to $OutputFormat" -Level Info -Source "Export-IntuneInventoryReport"
    }
    
    process {
        try {
            # Generate the report
            $Report = Get-IntuneInventoryReport -ReportType $ReportType -OutputFormat 'Object'
            
            # Create output directory if it doesn't exist
            $OutputDir = Split-Path -Path $OutputPath -Parent
            if ($OutputDir -and -not (Test-Path -Path $OutputDir)) {
                $null = New-Item -Path $OutputDir -ItemType Directory -Force
                Write-IntuneInventoryLog -Message "Created output directory: $OutputDir" -Level Info
            }
            
            # Save the report
            $null = Save-ReportToFile -Report $Report -Format $OutputFormat -Path $OutputPath -IncludeMetadata:$IncludeMetadata
            
            Write-IntuneInventoryLog -Message "Report exported successfully to: $OutputPath" -Level Info
            
            return @{
                Success = $true
                ReportPath = $OutputPath
                Format = $OutputFormat
                ReportType = $ReportType
                FileSize = (Get-Item -Path $OutputPath).Length
                ExportedAt = Get-Date
            }
        }
        catch {
            $ErrorMessage = "Failed to export report: $($_.Exception.Message)"
            Write-IntuneInventoryLog -Message $ErrorMessage -Level Error
            throw $ErrorMessage
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Report export completed" -Level Info -Source "Export-IntuneInventoryReport"
    }
}

# Helper functions for generating specific report types
function Get-InventorySummaryReport {
    param([hashtable]$FilterBy)
    
    $Stats = Get-StorageStatistics
    $Summary = @{
        StorageLocation = $Stats.StorageRoot
        TotalApplications = $Stats.Applications
        TotalScripts = $Stats.Scripts
        TotalRemediations = $Stats.Remediations
        TotalAssignments = $Stats.Assignments
        TotalInventoryRuns = $Stats.InventoryRuns
        LastInventoryRun = "Not Available"
    }
    
    # Get last inventory run info
    try {
        $InventoryData = Import-InventoryData -ItemType "InventoryRuns"
        if ($InventoryData -and $InventoryData.Count -gt 0) {
            $LastRun = $InventoryData | Sort-Object LastUpdated -Descending | Select-Object -First 1
            $Summary.LastInventoryRun = "$($LastRun.RunType) at $($LastRun.EndTime)"
        }
    }
    catch {
        Write-IntuneInventoryLog -Message "Could not retrieve last inventory run info: $($_.Exception.Message)" -Level Warning
    }
    
    return $Summary
}

function Get-ApplicationsReport {
    param([hashtable]$FilterBy)
    
    try {
        $Applications = Import-InventoryData -ItemType "Applications"
        if ($FilterBy) {
            $Applications = Apply-ReportFilter -Data $Applications -FilterBy $FilterBy
        }
        return $Applications
    }
    catch {
        Write-IntuneInventoryLog -Message "Could not retrieve applications data: $($_.Exception.Message)" -Level Warning
        return @()
    }
}

function Get-ScriptsReport {
    param([hashtable]$FilterBy)
    
    try {
        $Scripts = Import-InventoryData -ItemType "Scripts"
        if ($FilterBy) {
            $Scripts = Apply-ReportFilter -Data $Scripts -FilterBy $FilterBy
        }
        return $Scripts
    }
    catch {
        Write-IntuneInventoryLog -Message "Could not retrieve scripts data: $($_.Exception.Message)" -Level Warning
        return @()
    }
}

function Get-RemediationsReport {
    param([hashtable]$FilterBy)
    
    try {
        $Remediations = Import-InventoryData -ItemType "Remediations"
        if ($FilterBy) {
            $Remediations = Apply-ReportFilter -Data $Remediations -FilterBy $FilterBy
        }
        return $Remediations
    }
    catch {
        Write-IntuneInventoryLog -Message "Could not retrieve remediations data: $($_.Exception.Message)" -Level Warning
        return @()
    }
}

function Get-AssignmentsReport {
    param([hashtable]$FilterBy)
    
    try {
        $Assignments = Import-InventoryData -ItemType "Assignments"
        if ($FilterBy) {
            $Assignments = Apply-ReportFilter -Data $Assignments -FilterBy $FilterBy
        }
        return $Assignments
    }
    catch {
        Write-IntuneInventoryLog -Message "Could not retrieve assignments data: $($_.Exception.Message)" -Level Warning
        return @()
    }
}

function Get-InventoryRunsReport {
    param([hashtable]$FilterBy)
    
    try {
        $InventoryRuns = Import-InventoryData -ItemType "InventoryRuns"
        if ($FilterBy) {
            $InventoryRuns = Apply-ReportFilter -Data $InventoryRuns -FilterBy $FilterBy
        }
        return $InventoryRuns
    }
    catch {
        Write-IntuneInventoryLog -Message "Could not retrieve inventory runs data: $($_.Exception.Message)" -Level Warning
        return @()
    }
}

function Apply-ReportFilter {
    param(
        [array]$Data,
        [hashtable]$FilterBy
    )
    
    if (-not $FilterBy -or -not $Data) {
        return $Data
    }
    
    $FilteredData = $Data
    
    foreach ($Property in $FilterBy.Keys) {
        $FilteredData = $FilteredData | Where-Object { $_.$Property -like "*$($FilterBy[$Property])*" }
    }
    
    return $FilteredData
}

function Save-ReportToFile {
    param(
        [object]$Report,
        [string]$Format,
        [string]$Path,
        [switch]$IncludeMetadata
    )
    
    $OutputContent = if ($IncludeMetadata) { $Report } else { $Report.Data }
    
    switch ($Format) {
        'JSON' {
            $OutputContent | ConvertTo-SafeJson | Out-File -FilePath $Path -Encoding UTF8
        }
        'CSV' {
            if ($OutputContent -is [array]) {
                $OutputContent | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
            }
            else {
                # For complex objects, flatten them
                $Flattened = @()
                foreach ($Item in $OutputContent.PSObject.Properties) {
                    $Flattened += [PSCustomObject]@{
                        Property = $Item.Name
                        Value = $Item.Value
                    }
                }
                $Flattened | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
            }
        }
        'HTML' {
            $Html = Generate-HtmlReport -Report $OutputContent
            $Html | Out-File -FilePath $Path -Encoding UTF8
        }
        'XML' {
            $OutputContent | Export-Clixml -Path $Path
        }
    }
}

function Generate-HtmlReport {
    param([object]$Report)
    
    $Html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Intune Inventory Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1, h2 { color: #0078d4; }
        table { border-collapse: collapse; width: 100%; margin: 10px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .metadata { background-color: #f9f9f9; padding: 10px; border-radius: 5px; margin-bottom: 20px; }
    </style>
</head>
<body>
    <h1>Intune Inventory Report</h1>
    <div class="metadata">
        <strong>Generated:</strong> $(Get-Date)<br>
        <strong>Generated By:</strong> $env:USERNAME<br>
        <strong>Storage Location:</strong> $($Script:StorageRoot)
    </div>
"@
    
    if ($Report.Metadata) {
        $Html += "<h2>Report Metadata</h2>"
        $Html += ($Report.Metadata | ConvertTo-Html -Fragment)
    }
    
    if ($Report.Data) {
        $Html += "<h2>Report Data</h2>"
        $Html += ($Report.Data | ConvertTo-Html -Fragment)
    }
    else {
        $Html += ($Report | ConvertTo-Html -Fragment)
    }
    
    $Html += "</body></html>"
    
    return $Html
}

function Format-ReportOutput {
    param(
        [object]$Report,
        [string]$Format
    )
    
    switch ($Format) {
        'JSON' {
            return ($Report | ConvertTo-SafeJson)
        }
        'CSV' {
            return ($Report | ConvertTo-Csv -NoTypeInformation)
        }
        'HTML' {
            return (Generate-HtmlReport -Report $Report)
        }
        default {
            return $Report
        }
    }
}
