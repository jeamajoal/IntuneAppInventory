# IntuneInventory PowerShell Module
# Main module file for Intune application, script, and remediation inventory
# Uses JSON-based storage for user-friendly data management

# Module variables for JSON storage system
$Script:IntuneConnection = $null
$Script:StorageRoot = $null
$Script:StoragePaths = @{}
$Script:InventoryData = @{}

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

# Module cleanup - Production pattern with JSON storage
$MyInvocation.MyCommand.ScriptBlock.Module.OnRemove = {
    # Clean up Graph authentication resources
    if ($Script:GraphToken) {
        try {
            $Script:GraphToken = $null
            $Script:TokenExpiry = $null
            $Script:GraphHeaders = @{}
            $Script:ConnectionInfo = $null
        }
        catch {
            Write-Warning "Error clearing Graph authentication: $($_.Exception.Message)"
        }
    }
    
    # Clean up JSON storage resources
    try {
        $Script:StorageRoot = $null
        $Script:StoragePaths = @{}
        $Script:InventoryData = @{}
    }
    catch {
        Write-Warning "Error clearing storage resources: $($_.Exception.Message)"
    }
}

# Export module members (defined in manifest)
Export-ModuleMember -Function @(
    'Initialize-IntuneInventoryStorage',
    'Connect-IntuneInventory',
    'Disconnect-IntuneInventory',
    'Test-IntuneConnection',
    'Get-IntuneConnectionInfo',
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
    'Backup-JsonStorage',
    'Clear-StorageCache',
    'Get-StorageStatistics'
)
