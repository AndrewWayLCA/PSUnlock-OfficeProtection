# PSUnlock-OfficeProtection
Commands to remove basic worksheet and workbook modify passwords (DOES NOT UNLOCK OPEN PASSWORDS)

See the most up-to-date help, status and documentation at the [GitHub Repository](https://github.com/AndrewWayLCA/PSUnlock-OfficeProtection).

## Installation
This module is available from the [PowerShell Gallery](https://www.powershellgallery.com/packages/PSUnlock-OfficeProtection).
```powershell
Install-Module -Name PSUnlock-OfficeProtection
```

# CmdLets
## Unlock-ExcelPasswordProtection
### Synopsis
Removes password protection from Excel files (.xlsx, .xlsm) for workbook and/or worksheet.

### Description
Unlock-ExcelPasswordProtection extracts, modifies, and repackages Excel files to remove workbook and/or worksheet password protection. Optionally creates a backup before modifying the file.

### Parameters
- `-Path <String>`: Path to the Excel file (.xlsx or .xlsm) to unlock.
- `-NoBackup <SwitchParameter>`: If specified, skips creation of a backup file before modification.
- `-Remove <String>`: Specifies which protection to remove: 'Workbook', 'Worksheet', or 'All'. Default is 'All'.
- `-OutFile <String>`: Optional path for the output file. If not specified, overwrites the original file.

### Syntax
```powershell
Unlock-ExcelPasswordProtection [-Path] <String> [-NoBackup] [-Remove <String>] [-OutFile <String>] [<CommonParameters>]
```

### Notes
Requires PowerShell 5.1 or later. Uses Expand-Archive and Compress-Archive cmdlets.

### Examples
Use inbuilt `Get-Help` commands to get examples.

# Could Be Better
Notes for coming, planned, possible improvements to the module will be listed here.

These may or may not be implemented.

Find [the file](CBB.md) here.

# How it works
Have a read of [How it works](How-It-Works.md) to see what this code gets up to. It's pretty easy to remove protection from any modern Office document by extracting the file to it's component `.xml` parts, modifying those parts and then reconstructing the file.