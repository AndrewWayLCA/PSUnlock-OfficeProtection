function Remove-WorkbookProtection {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TempPathRoot
    )

    $workbookXmlPath = Join-Path -Path $TempPathRoot -ChildPath "xl\workbook.xml"
    if (Test-Path -Path $workbookXmlPath) {
        $isRemoved = Remove-XmlNodeFromFile -FilePath $workbookXmlPath -NodeName "workbookProtection"
        if ($isRemoved) {
            Write-Verbose "Workbook protection removed successfully."
        } else {
            Write-Verbose "No protection or failed to remove workbook protection."
        }
        # also remove the workbook protection from the shared strings
        $isRemoved = Remove-XmlNodeFromFile -FilePath $workbookXmlPath -NodeName "fileSharing"
        if ($isRemoved) {
            Write-Verbose "File sharing protection removed successfully."
        } else {
            Write-Verbose "No file sharing protection or failed to remove it."
        }
    }
}
