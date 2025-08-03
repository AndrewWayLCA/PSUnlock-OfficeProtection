<#
.SYNOPSIS
    PSUnlock-OfficeProtection module loader.

.DESCRIPTION
    Contains functions to unlock password protection from Office files, specifically Excel workbooks and worksheets.
.NOTES
    Requires PowerShell 5.0 or later.
    Place custom scripts in Private or Public as needed.
#>
# Dot-source all Private scripts
$privateScripts = Get-ChildItem -Path "$PSScriptRoot/Private" -Filter *.ps1 -ErrorAction SilentlyContinue
foreach ($script in $privateScripts) {
    . $script.FullName
}

# Dot-source all Public scripts and collect function names
$publicScripts = Get-ChildItem -Path "$PSScriptRoot/Public" -Filter *.ps1 -ErrorAction SilentlyContinue
$publicFunctions = @()
foreach ($script in $publicScripts) {
    . $script.FullName
    $functionName = [System.IO.Path]::GetFileNameWithoutExtension($script.Name)
    $publicFunctions += $functionName
}

Export-ModuleMember -Function $publicFunctions
