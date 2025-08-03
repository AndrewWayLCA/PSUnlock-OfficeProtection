# How it works
Modern office files (.xlsx, etc) are actually just `.zip` files with a [specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offfflp/8aea05e3-8c1e-4a9a-9614-31f71e679456) for their contents.

When users set passwords to protect parts of an Excel Worksheet, the structure of a Workbook and parts of Word documents, they often forget those passwords.

Fortunately with the structure above the designation of protection is stored as `.xml` in the sub-components/nodes of that .zip file.

Workbook protection information is stored in the zip file in the `xl\workbook.xml` file. Sheet protections are stored in the `xl\worksheets\sheetX.xml` files.

So, the solution:
- a PowerShell CmdLet that can open, extract the Office file, remove the XML nodes related to protection
- then re-compose the file by zipping those extracted components to the original file

Effectively (once the module is installed) you can remove password protection from Workbooks and Sheets with a single command:

```powershell
Unlock-ExcelPasswordProtection -Path .\ProtectedFile.xlsx
```