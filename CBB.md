# Could Be Better
The Could Be Better page. Here are brain-dump ideas for this code.

## PSGallery Publishing
Integrate the module into [PSGallery](https://www.powershellgallery.com/) so that it can be installed with standard PowerShell commands.

```powershell
Install-Module -Name PSUnlock-OfficeProtection
```

## Object Pipeline
Review usage of the Object Pipeline and options around how to send in one or more files to be unlocked.

```powershell
Get-ChildItem -Path .\*.xlsx | Unlock-ExcelPasswordProtection
```

It may need additional Parameter like `-OutPath` instead of `-OutFile`, or it might need a rule that `-OutFile` can't be used when handling objects from the pipeline.

## Create `Unlock-WordPasswordProtection`
Review what password security can be applied to Word `.docx` and `.docm` files that can be easily removed by extracting the files to `.xml` and removing the document protection sections of the `.xml`.

## Sign the Code
With the code signing certificate coming, use this to integrate code signing into the release pipeline.