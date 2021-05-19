# 4d-plugin-office-automation
instruct Word and Excel to open and close documents

### Windows

[How to use a type library for Office Automation from Visual C++.NET](https://docs.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/use-type-library-for-office-from-visual-c-net)

* create new MFC Application project
* add > class > MFC class from Typelib

* from Word 16.0
  * Appliction (rename: `WordAppliction`)
  * Document (rename: `WordDocument`)

* from Excel 16.0
  * Appliction (rename: `ExcelAppliction`)
  * Workbook (rename: `ExcelWorkbook`)

* release
  * additional dependencies: `uafxcw.lib;Libcmt.lib;`
  * ignore default libraries: `uafxcw.lib;Libcmt.lib` 

* debug
  * additional dependencies: `uafxcwd.lib;Libcmtd.lib`   
  * ignore default libraries: `uafxcwd.lib;Libcmtd.lib`
