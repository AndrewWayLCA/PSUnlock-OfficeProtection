function Remove-WorksheetProtection {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TempPathRoot
    )

    $sheetsXmlPath = Join-Path -Path $TempPathRoot -ChildPath "xl\worksheets"
    if (Test-Path -Path $sheetsXmlPath) {
        Get-ChildItem -Path $sheetsXmlPath -Filter "*.xml" | ForEach-Object {
            $isRemoved = Remove-XmlNodeFromFile -FilePath $_.FullName -NodeName "sheetProtection"
            if ($isRemoved) {
                Write-Verbose "Worksheet protection removed from: $($_.Name)"
            } else {
                Write-Verbose "No protection or failed to remove worksheet protection from: $($_.Name)"
            }
        }
    }
}
