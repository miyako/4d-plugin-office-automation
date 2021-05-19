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

On windows, `CreateDispatch` creates a new dispath. need to search for running instances first.

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
