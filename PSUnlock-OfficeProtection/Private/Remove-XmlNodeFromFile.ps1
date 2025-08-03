function Remove-XmlNodeFromFile {
    param (
        [Parameter(Mandatory = $true)]
        [string] $FilePath,
        [Parameter(Mandatory = $true)]
        [string] $NodeName
    )

    [xml]$xmlContent = Get-Content -Path $FilePath

    # Create a namespace manager
    $nsMgr = New-Object System.Xml.XmlNamespaceManager($xmlContent.NameTable)
    $nsMgr.AddNamespace("ns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

    # boy, what a fuck around to get the SelectSingleNode to work
    $node = $xmlContent.SelectSingleNode("//ns:$NodeName", $nsMgr)
    
    if ($node) {
        $node.ParentNode.RemoveChild($node) | Out-Null
        $xmlContent.Save($FilePath)
        return $true
    } else {
        return $false
    }
}
