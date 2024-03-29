![version](https://img.shields.io/badge/version-19%2B-5682DF)
![platform](https://img.shields.io/static/v1?label=platform&message=mac-intel%20|%20mac-arm%20|%20win-64&color=blue)
[![license](https://img.shields.io/github/license/miyako/4d-plugin-office-automation)](LICENSE)
![downloads](https://img.shields.io/github/downloads/miyako/4d-plugin-office-automation/total)

# 4d-plugin-office-automation
instruct Word and Excel to open and close documents

### Examples

* Open Word document

```4d
$file:=Folder(fk resources folder).file("テスト.docx")

$params:=New object
$params.app:="word"
$params.path:=$file.platformPath  //if omitted, save in temporary folder (mac)
$params.command:="open"

$status:=office automation ($params)
```

* Save and close Word document

```4d
$file:=Folder(fk resources folder).file("テスト.docx")

$params:=New object
$params.app:="word"
$params.command:="close"
$params.name:=$file.name+$file.extension  //specify the name, not a path

$status:=office automation ($params)
```

* Open Excel document

```4d
$file:=Folder(fk resources folder).file("テスト.xlsx")

$params:=New object
$params.app:="excel"
$params.path:=$file.platformPath  //if omitted, save in temporary folder (mac)
$params.command:="open"

$status:=office automation ($params)
```

* Save and close Excel document

```4d
$file:=Folder(fk resources folder).file("テスト.xlsx")

$params:=New object
$params.app:="excel"
$params.command:="close"
$params.name:=$file.name+$file.extension  //specify the name, not a path

$status:=office automation ($params)
```

### Mac

* create header file from app

```
sdef Word.app > word.sdef
sdp --basename Word -fh Word.sdef Word.h
```

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
  * additional dependencies: `uafxcw.lib;Libcmt.lib`
  * ignore specific default libraries: `uafxcw.lib;Libcmt.lib` 

* debug
  * additional dependencies: `uafxcwd.lib;Libcmtd.lib`   
  * ignore specific default libraries: `uafxcwd.lib;Libcmtd.lib`

* other
  * `uxtheme.lib;windowscodecs.lib` 

remove 

```c
#import "C:\\Program Files (x86)\\Microsoft Office\\Root\\Office16\\EXCEL.EXE" no_namespace
#import "C:\\Program Files (x86)\\Microsoft Office\\Root\\Office16\\MSWORD.OLB" no_namespace
```

remove

```c
#include "windows.h"
```

### Issues 

On windows, `CreateDispatch` creates a new dispatch. need to search for running instances first.

```c
COleException e;
WordApplication wdApp;

bool isAvailable = false;

CLSID clsid;
CLSIDFromProgID(L"Word.Application", &clsid);

IUnknown *pUnknown;
HRESULT hr = GetActiveObject(clsid, NULL, (IUnknown**)&pUnknown);
if (SUCCEEDED(hr)) {
	//there is a running instance
	// Get IDispatch interface for Automation...
	IDispatch *pDispatch;
	hr = pUnknown->QueryInterface(IID_IDispatch, (void **)&pDispatch);
	if (SUCCEEDED(hr)) {
		wdApp.AttachDispatch(pDispatch, TRUE);
		isAvailable = true;
	}
	}
	else {
		if (wdApp.CreateDispatch(L"Word.Application", &e)) {
		isAvailable = true;
	}
}
```

On Mac, the name of a Word document is its file name.

On Windows, it is localised serial, e.g. "文書 1". depending on the options passed to [Open](https://docs.microsoft.com/en-us/office/vba/api/word.documents.open).

in particular, `OpenAndRepair` should  be `False` to keep the original name and path.
