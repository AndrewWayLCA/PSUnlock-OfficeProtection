
<#
.SYNOPSIS
Removes password protection from Excel files (.xlsx, .xlsm) for workbook and/or worksheet.

.DESCRIPTION
Unlock-ExcelPasswordProtection extracts, modifies, and repackages Excel files to remove workbook and/or worksheet password protection. Optionally creates a backup before modifying the file.

.PARAMETER Path
Path to the Excel file (.xlsx or .xlsm) to unlock.

.PARAMETER NoBackup
If specified, skips creation of a backup file before modification.

.PARAMETER Remove
Specifies which protection to remove: 'Workbook', 'Worksheet', or 'All'. Default is 'All'.

.PARAMETER OutFile
Optional path for the output file. If not specified, overwrites the original file.

.EXAMPLE
Unlock-ExcelPasswordProtection -Path 'C:\Files\Protected.xlsx'
Removes all password protection from the specified Excel file and creates a backup.

.EXAMPLE
Unlock-ExcelPasswordProtection -Path 'C:\Files\Protected.xlsx' -Remove Workbook -NoBackup
Removes only workbook protection and skips backup creation.

.NOTES
Requires PowerShell 5.1 or later. Uses Expand-Archive and Compress-Archive cmdlets.
#>
function Unlock-ExcelPasswordProtection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string] $Path,

        [Parameter(Mandatory = $false)]
        [switch] $NoBackup = $false,

        [ValidateSet("Workbook", "Worksheet", "All")]
        [string] $Remove = "All",

        [Parameter(Mandatory = $false)]
        [string] $OutFile
    )

    # Implementation of the function goes here
    Write-Verbose "Unlocking Excel password protection for file: $Path"
    
    # convert the file path into a FileInfo object so that it can be properly handled
    $sourceFile = Get-Item -Path $Path

    # validate the file exists and is an Excel file extension
    
    if (-not $sourceFile.Exists) {
        Write-Error "The specified file does not exist: $Path"
        return
    }

    if ($sourceFile.PSIsContainer -and $sourceFile.Extension -ne ".xlsx" -and $sourceFile.Extension -ne ".xlsm") {
        Write-Error "The specified file is not a valid Excel file: $Path"
        return
    }

    # Create a backup if NoBackup is not set
    $backupFileName = $sourceFile.DirectoryName + "\" + $sourceFile.BaseName + ".bak"
    if (-not $NoBackup) {
        Backup-File -Path $sourceFile.FullName -BackupPath $backupFileName
    } else {
        Write-Verbose "Skipping backup creation for file: $Path"
    }

    $tempFolder = $env:TEMP + "\" + [guid]::NewGuid().ToString()
    New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null

    Write-Verbose "Temporary folder created at: $tempFolder"
    try {
        # Unzip the Excel file to the temporary folder
        Expand-Archive -Path $sourceFile.FullName -DestinationPath $tempFolder -Force -Verbose:$false

        # Remove password protection based on the specified option
        if ($Remove -eq "Workbook" -or $Remove -eq "All") {
            Remove-WorkbookProtection -TempPathRoot $tempFolder
        }

        if ($Remove -eq "Worksheet" -or $Remove -eq "All") {
            Remove-WorksheetProtection -TempPathRoot $tempFolder
        }

        # Repackage the Excel file
        $outputFile = if ($OutFile) { $OutFile } else { $sourceFile.FullName }
        Compress-Archive -Path (Join-Path -Path $tempFolder -ChildPath "*") -DestinationPath $outputFile -Force

        Write-Verbose "Excel file unlocked and saved to: $outputFile"
    } catch {
        Write-Error "An error occurred while unlocking the Excel file: $($sourceFile.Name)"
    } finally {
        # Clean up temporary files
        Remove-Item -Path $tempFolder -Recurse -Force
        Write-Verbose "Temporary folder cleaned up."
    }

}
