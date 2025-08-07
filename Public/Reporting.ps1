function Get-IntuneInventoryReport {
    <#
    .SYNOPSIS
    Generates comprehensive reports from the Intune inventory database.
    
    .DESCRIPTION
    Creates various types of reports from the stored Intune inventory data including
    applications, scripts, remediations, and their assignments.
    
    .PARAMETER ReportType
    The type of report to generate (Summary, Applications, Scripts, Remediations, Assignments, All).
    
    .PARAMETER Format
    The output format for the report (Object, Table, CSV, JSON).
    
    .PARAMETER IncludeSourceCode
    Include source code information in the report.
    
    .PARAMETER FilterMissingSourceCode
    Only include items that are missing source code.
    
    .PARAMETER OutputPath
    Path to save the report file (for CSV and JSON formats).
    
    .EXAMPLE
    Get-IntuneInventoryReport -ReportType Summary
    
    .EXAMPLE
    Get-IntuneInventoryReport -ReportType Applications -Format CSV -OutputPath "C:\Reports\Apps.csv"
    
    .EXAMPLE
    Get-IntuneInventoryReport -ReportType All -FilterMissingSourceCode
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet('Summary', 'Applications', 'Scripts', 'Remediations', 'Assignments', 'All')]
        [string]$ReportType = 'Summary',
        
        [Parameter()]
        [ValidateSet('Object', 'Table', 'CSV', 'JSON')]
        [string]$Format = 'Object',
        
        [Parameter()]
        [switch]$IncludeSourceCode,
        
        [Parameter()]
        [switch]$FilterMissingSourceCode,
        
        [Parameter()]
        [string]$OutputPath
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting inventory report generation: $ReportType" -Level Info -Source "Get-IntuneInventoryReport"
        
        # Verify database connection
        if (-not $Script:DatabaseConnection -or -not (Test-DatabaseConnection -DatabaseConnection $Script:DatabaseConnection)) {
            throw "Database connection is not available. Please run Connect-IntuneInventory first."
        }
    }
    
    process {
        try {
            $ReportData = @{}
            
            switch ($ReportType) {
                'Summary' {
                    $ReportData = Get-InventorySummaryReport
                }
                'Applications' {
                    $ReportData = Get-ApplicationsReport -IncludeSourceCode:$IncludeSourceCode -FilterMissingSourceCode:$FilterMissingSourceCode
                }
                'Scripts' {
                    $ReportData = Get-ScriptsReport -IncludeSourceCode:$IncludeSourceCode -FilterMissingSourceCode:$FilterMissingSourceCode
                }
                'Remediations' {
                    $ReportData = Get-RemediationsReport -IncludeSourceCode:$IncludeSourceCode -FilterMissingSourceCode:$FilterMissingSourceCode
                }
                'Assignments' {
                    $ReportData = Get-AssignmentsReport
                }
                'All' {
                    $ReportData = @{
                        Summary = Get-InventorySummaryReport
                        Applications = Get-ApplicationsReport -IncludeSourceCode:$IncludeSourceCode -FilterMissingSourceCode:$FilterMissingSourceCode
                        Scripts = Get-ScriptsReport -IncludeSourceCode:$IncludeSourceCode -FilterMissingSourceCode:$FilterMissingSourceCode
                        Remediations = Get-RemediationsReport -IncludeSourceCode:$IncludeSourceCode -FilterMissingSourceCode:$FilterMissingSourceCode
                        Assignments = Get-AssignmentsReport
                    }
                }
            }
            
            # Format and output the report
            switch ($Format) {
                'Object' {
                    return $ReportData
                }
                'Table' {
                    if ($ReportType -eq 'All') {
                        foreach ($Section in $ReportData.Keys) {
                            Write-Host "`n=== $Section ===" -ForegroundColor Green
                            $ReportData[$Section] | Format-Table -AutoSize
                        }
                    }
                    else {
                        $ReportData | Format-Table -AutoSize
                    }
                }
                'CSV' {
                    if (-not $OutputPath) {
                        $OutputPath = Join-Path -Path $PWD -ChildPath "IntuneInventoryReport_$($ReportType)_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                    }
                    
                    if ($ReportType -eq 'All') {
                        # Save each section to separate CSV files
                        $BaseOutputPath = [System.IO.Path]::GetFileNameWithoutExtension($OutputPath)
                        $Extension = [System.IO.Path]::GetExtension($OutputPath)
                        $Directory = [System.IO.Path]::GetDirectoryName($OutputPath)
                        
                        foreach ($Section in $ReportData.Keys) {
                            $SectionPath = Join-Path -Path $Directory -ChildPath "$BaseOutputPath`_$Section$Extension"
                            $ReportData[$Section] | Export-Csv -Path $SectionPath -NoTypeInformation
                            Write-Host "Report section '$Section' saved to: $SectionPath" -ForegroundColor Cyan
                        }
                    }
                    else {
                        $ReportData | Export-Csv -Path $OutputPath -NoTypeInformation
                        Write-Host "Report saved to: $OutputPath" -ForegroundColor Cyan
                    }
                }
                'JSON' {
                    if (-not $OutputPath) {
                        $OutputPath = Join-Path -Path $PWD -ChildPath "IntuneInventoryReport_$($ReportType)_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
                    }
                    
                    $ReportData | ConvertTo-Json -Depth 10 | Out-File -FilePath $OutputPath -Encoding UTF8
                    Write-Host "Report saved to: $OutputPath" -ForegroundColor Cyan
                }
            }
        }
        catch {
            Write-IntuneInventoryLog -Message "Report generation failed: $($_.Exception.Message)" -Level Error
            throw
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Report generation completed" -Level Info -Source "Get-IntuneInventoryReport"
    }
}

function Export-IntuneInventoryReport {
    <#
    .SYNOPSIS
    Exports comprehensive Intune inventory reports to various formats.
    
    .DESCRIPTION
    Generates and exports detailed inventory reports including executive summaries,
    detailed item lists, and source code analysis reports.
    
    .PARAMETER OutputDirectory
    Directory where report files will be saved.
    
    .PARAMETER IncludeSourceCode
    Include source code in the exported reports.
    
    .PARAMETER ReportFormats
    Array of formats to export (CSV, JSON, HTML, Excel).
    
    .EXAMPLE
    Export-IntuneInventoryReport -OutputDirectory "C:\Reports"
    
    .EXAMPLE
    Export-IntuneInventoryReport -OutputDirectory "C:\Reports" -IncludeSourceCode -ReportFormats @('CSV', 'HTML')
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,
        
        [Parameter()]
        [switch]$IncludeSourceCode,
        
        [Parameter()]
        [ValidateSet('CSV', 'JSON', 'HTML', 'Excel')]
        [string[]]$ReportFormats = @('CSV', 'JSON')
    )
    
    begin {
        Write-IntuneInventoryLog -Message "Starting comprehensive report export" -Level Info -Source "Export-IntuneInventoryReport"
        
        # Verify database connection
        if (-not $Script:DatabaseConnection -or -not (Test-DatabaseConnection -DatabaseConnection $Script:DatabaseConnection)) {
            throw "Database connection is not available. Please run Connect-IntuneInventory first."
        }
        
        # Ensure output directory exists
        if (-not (Test-Path -Path $OutputDirectory)) {
            New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
        }
    }
    
    process {
        try {
            $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $ReportFiles = @()
            
            # Generate comprehensive report data
            Write-IntuneInventoryLog -Message "Generating comprehensive report data..." -Level Info
            $AllReports = Get-IntuneInventoryReport -ReportType All -IncludeSourceCode:$IncludeSourceCode
            
            foreach ($Format in $ReportFormats) {
                switch ($Format) {
                    'CSV' {
                        Write-IntuneInventoryLog -Message "Exporting CSV reports..." -Level Info
                        foreach ($Section in $AllReports.Keys) {
                            $FilePath = Join-Path -Path $OutputDirectory -ChildPath "IntuneInventory_$Section`_$Timestamp.csv"
                            $AllReports[$Section] | Export-Csv -Path $FilePath -NoTypeInformation
                            $ReportFiles += $FilePath
                        }
                    }
                    'JSON' {
                        Write-IntuneInventoryLog -Message "Exporting JSON report..." -Level Info
                        $FilePath = Join-Path -Path $OutputDirectory -ChildPath "IntuneInventory_Complete_$Timestamp.json"
                        $AllReports | ConvertTo-Json -Depth 10 | Out-File -FilePath $FilePath -Encoding UTF8
                        $ReportFiles += $FilePath
                    }
                    'HTML' {
                        Write-IntuneInventoryLog -Message "Exporting HTML report..." -Level Info
                        $FilePath = Join-Path -Path $OutputDirectory -ChildPath "IntuneInventory_Report_$Timestamp.html"
                        New-HtmlReport -ReportData $AllReports -OutputPath $FilePath
                        $ReportFiles += $FilePath
                    }
                    'Excel' {
                        Write-IntuneInventoryLog -Message "Excel export requires additional modules (ImportExcel)" -Level Warning
                        # Excel export would require ImportExcel module
                        # Implementation would go here if module is available
                    }
                }
            }
            
            # Generate summary report
            Write-IntuneInventoryLog -Message "Generating executive summary..." -Level Info
            $SummaryPath = Join-Path -Path $OutputDirectory -ChildPath "IntuneInventory_ExecutiveSummary_$Timestamp.txt"
            New-ExecutiveSummary -ReportData $AllReports -OutputPath $SummaryPath
            $ReportFiles += $SummaryPath
            
            Write-Host "Report export completed!" -ForegroundColor Green
            Write-Host "Files generated:" -ForegroundColor Cyan
            foreach ($File in $ReportFiles) {
                Write-Host "  $File" -ForegroundColor White
            }
            
            return $ReportFiles
        }
        catch {
            Write-IntuneInventoryLog -Message "Report export failed: $($_.Exception.Message)" -Level Error
            throw
        }
    }
    
    end {
        Write-IntuneInventoryLog -Message "Report export completed" -Level Info -Source "Export-IntuneInventoryReport"
    }
}
