<#
    .SYNOPSIS
    Creates a backup copy of a specified file.

    .DESCRIPTION
    Copies the file at the given path to the specified backup location. Overwrites the backup if it already exists.

    .PARAMETER Path
    The path to the file to back up.

    .PARAMETER BackupPath
    The destination path for the backup copy.

    .EXAMPLE
    Backup-File -Path 'C:\Data\file.xlsx' -BackupPath 'C:\Backup\file.xlsx'

    .NOTES
    Internal/private function for use within the module.
#>
function Backup-File {
    param (
        [Parameter(Mandatory = $true)]
        [string] $Path,
        [Parameter(Mandatory = $true)]
        [string] $BackupPath
    )

    Write-Verbose "Creating backup for file: $Path"
    Write-Verbose "Backup will be saved to: $BackupPath"

    Copy-Item -Path $Path -Destination $BackupPath -Force
    Write-Verbose "Backup created at: $backupPath"
}
