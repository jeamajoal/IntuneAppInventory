# IntuneInventory PowerShell Module
# Main module file for Intune application, script, and remediation inventory

# Import required assemblies for SQLite
Add-Type -Path (Join-Path $PSScriptRoot "System.Data.SQLite.dll") -ErrorAction SilentlyContinue

# Module variables
$Script:IntuneConnection = $null
$Script:DatabasePath = $null
$Script:DatabaseConnection = $null

# Import functions from subdirectories
$FunctionDirectories = @('Private', 'Public')
foreach ($Directory in $FunctionDirectories) {
    $FunctionPath = Join-Path -Path $PSScriptRoot -ChildPath $Directory
    if (Test-Path $FunctionPath) {
        $Functions = Get-ChildItem -Path $FunctionPath -Filter "*.ps1" -Recurse
        foreach ($Function in $Functions) {
            try {
                . $Function.FullName
                Write-Verbose "Imported function: $($Function.BaseName)"
            }
            catch {
                Write-Error "Failed to import function $($Function.BaseName): $($_.Exception.Message)"
            }
        }
    }
}

# Module cleanup
$MyInvocation.MyCommand.ScriptBlock.Module.OnRemove = {
    if ($Script:DatabaseConnection) {
        try {
            $Script:DatabaseConnection.Close()
            $Script:DatabaseConnection.Dispose()
        }
        catch {
            Write-Warning "Error closing database connection: $($_.Exception.Message)"
        }
    }
    
    if ($Script:IntuneConnection) {
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
        }
        catch {
            Write-Warning "Error disconnecting from Microsoft Graph: $($_.Exception.Message)"
        }
    }
}

# Export module members (defined in manifest)
Export-ModuleMember -Function @(
    'Initialize-IntuneInventoryDatabase',
    'Connect-IntuneInventory',
    'Disconnect-IntuneInventory',
    'Invoke-IntuneApplicationInventory',
    'Invoke-IntuneScriptInventory',
    'Invoke-IntuneRemediationInventory',
    'Get-IntuneInventoryReport',
    'Export-IntuneInventoryReport',
    'Add-IntuneInventorySourceCode',
    'Get-IntuneInventorySourceCode',
    'Update-IntuneInventoryItem',
    'Remove-IntuneInventoryItem',
    'Get-IntuneInventoryItem',
    'Get-IntuneInventoryAssignments',
    'Backup-IntuneInventoryDatabase',
    'Restore-IntuneInventoryDatabase'
)
