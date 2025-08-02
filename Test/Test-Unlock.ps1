$fdr = $PSScriptRoot + "\..\"
Import-Module -Name "$fdr\src\PSUnlock-OfficeProtection.psm1" -Force


foreach ($file in Get-ChildItem -Path ($fdr + "\Test\Protected")) {
    $filePath = $file.FullName
    $outFile = $fdr + "\Test\Unprotected\" + $file.Name

    # Execute the script with the parameters
    Unlock-ExcelPasswordProtection -Path $filePath -NoBackup:$true -Remove All -OutFile $outFile -Verbose
}