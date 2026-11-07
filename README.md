![version](https://img.shields.io/badge/version-19%2B-5682DF)
![platform](https://img.shields.io/static/v1?label=platform&message=mac-intel%20|%20mac-arm%20|%20win-64&color=blue)
[![license](https://img.shields.io/github/license/miyako/4d-plugin-office-automation)](LICENSE)
![downloads](https://img.shields.io/github/downloads/miyako/4d-plugin-office-automation/total)

# 4d-plugin-office-automation

Tells Microsoft Word and Microsoft Excel to open a document, or to save and close a document that is already open. On macOS the plugin sends Apple events to the Office app through ScriptingBridge; on Windows it drives the app through COM automation (MFC `COleDispatchDriver` wrappers). Everything goes through one command that takes an `Object` of options and returns an `Object` with a `success` flag and the document's `path`.

| Command | Returns | Purpose |
|---|---|---|
| [office automation](#office-automation) | `Object` | Open a Word/Excel document, or save and close an open one |

**Platforms:** macOS (Apple Silicon and Intel, macOS 11.5 or later for builds from the release workflow), Windows 64-bit. 4D v19 or later.

[Releases](https://github.com/miyako/4d-plugin-office-automation/releases)

---

## Requirements & platform notes

- **Microsoft Word and/or Excel must be installed.** The plugin talks to the installed app (`com.microsoft.Word` / `com.microsoft.Excel` on macOS, the `Word.Application` / `Excel.Application` COM servers on Windows). Outlook is not supported: `app:"outlook"` is accepted but does nothing and returns `success:false`.
- **macOS asks for Automation permission.** The first call makes macOS ask whether 4D may control Word or Excel (System Settings > Privacy & Security > Automation). If the permission is denied, the Apple events fail without a 4D error.
- **Paths are platform paths.** Pass `File.platformPath` (or an equivalent). On macOS the plugin reads the path as an **HFS path** (`Macintosh HD:Users:…`), not a POSIX path; a POSIX path (`/Users/…`) is turned into a wrong file URL and the command has no effect.
- **Documents are identified by name when closing.** `close` looks for open documents whose name equals `name` exactly (case-sensitive), not by path.
- **`success` does not always mean the app did what you asked.** See [Error handling & troubleshooting](#error-handling--troubleshooting) for what `success` actually reports on each platform.
- **The command blocks while Office works.** The call is synchronous; if Word or Excel shows a dialog (a password prompt, a conversion warning, a sign-in window), the 4D process waits until the app answers or the request times out.
- The command is declared thread-safe in the manifest, so it can be called in a preemptive process.

---

## office automation

### Syntax

```
office automation ( options ) -> status
```

| Parameter | Type | Description |
|---|---|---|
| `options` | Object | What to do; see the properties below |
| Result | Object | `success` and, on success, `path`; see below |

#### Options object

| Property | Type | Description |
|---|---|---|
| `app` | Text | `"word"` or `"excel"` (lower case, exact match) |
| `command` | Text | `"open"` or `"close"` (lower case, exact match) |
| `path` | Text | `open`: the file to open. `close`: where to save the document before closing it (Save As). Optional for `close`; see below |
| `name` | Text | `close`: the name of the open document, e.g. `"report.docx"` (the name, not a path). `open`: passed to Word for Mac as the `file name` parameter of its `open` command (effect not verified); ignored by Excel and on Windows |
| `password` | Text | Password used for all four Word passwords (open document, open template, write document, write template); for Excel on Windows, the open password **and** the write-reservation password |
| `templatePassword` | Text | Word: password for opening and for saving the template (overrides `password` for those two). Excel on Windows: write-reservation password (overrides `password`) |
| `writePassword` | Text | Word: password for saving the document and the template (overrides `password` and `templatePassword` for those two). Ignored by Excel |
| `writeTemplatePassword` | Text | Word: password for saving the template (overrides everything above for that one) |

Passwords are applied only by `open`. Excel for Mac ignores every password property.

#### Status object

| Property | Type | Description |
|---|---|---|
| `success` | Boolean | `True` if the command was carried out (see the platform notes below for what is checked) |
| `path` | Text | Set when `success` is `True`. `open`: the `path` you passed. `close`: the path the document was saved to |

### Description

#### open

Starts the app if needed, then opens the file at `path`.

- **On macOS**, the plugin uses the running copy of the app if there is one (the first one found), otherwise the app is launched. Word opens the file with conversion dialogs off, read-write, added to the recent files, not in repair mode, and, if the document is already open, the open window is activated (`revert` false). Excel opens the file without options.
- **On Windows**, the plugin attaches to a running instance (`GetActiveObject`) or starts a new one, and makes the app visible. Word opens the file read-write, visible, without dialogs, added to the recent files, not in repair mode, and with `Revert` set to `True`: **if the document is already open, unsaved changes are discarded and the file is reopened.** Excel opens the workbook with the given passwords and default options.

#### close

Finds the open documents named `name`, saves each one, and closes it.

- **With `path`**, the document is saved to `path` (Save As) before it is closed, and `path` is returned.
- **Without `path`** (or with an empty `path`):
  - **On macOS**, the document is saved into a **new temporary folder** (created by the system for each call; characters such as `:` `/` `\` `?` `%` `*` `|` `"` `<` `>` are removed from the file name) and that file's HFS path is returned. The original file is not updated, and the temporary folder is not removed by the plugin.
  - **On Windows**, the document is saved in place (Save), and its full path is returned.
- **On macOS**, if the app is not running it is launched to look for the document. All matching documents are closed (saved to the same `path` if one is given).
- **On Windows**, only a running instance is used; if the app is not running, `success` is `False`. Excel's alerts are turned off while the workbook is saved, then restored. **When no documents are left open after the close, the plugin quits Word or Excel.**

### Example

From the plugin's own test methods (`TEST_open_word.4dm`, `TEST_close_word.4dm`):

```4d
$file:=Folder:C1567(fk resources folder:K87:11).file("テスト.docx")

$params:=New object:C1471
$params.app:="word"
$params.path:=$file.platformPath  //if omitted, save in temporary folder (mac)
$params.command:="open"

$status:=office automation ($params)
```

```4d
$file:=Folder:C1567(fk resources folder:K87:11).file("テスト.docx")

$params:=New object:C1471
$params.app:="word"
$params.command:="close"
$params.name:=$file.name+$file.extension  //specify the name, not a path

$status:=office automation ($params)
```

The Excel versions (`TEST_open_excel.4dm`, `TEST_close_excel.4dm`) are the same with `"excel"` and `テスト.xlsx`.

Save a workbook to a chosen file, then close it:

```4d
$source:=Folder(fk resources folder).file("テスト.xlsx")
$target:=Folder(fk desktop folder).file("copy.xlsx")

$params:=New object
$params.app:="excel"
$params.command:="close"
$params.name:=$source.name+$source.extension
$params.path:=$target.platformPath

$status:=office automation($params)

If ($status.success)
	ALERT("Saved to "+$status.path)
End if
```

Open a password-protected Word document:

```4d
$params:=New object
$params.app:="word"
$params.command:="open"
$params.path:=$file.platformPath
$params.password:="secret"

$status:=office automation($params)
```

---

## Error handling & troubleshooting

- **No 4D error is raised.** Every failure is reported only as `success:false` in the returned object; check it after each call.
- **On macOS, `open` reports `success:true` as soon as the request is sent.** The plugin does not check whether Word or Excel actually opened the file: a missing file, a wrong path, a denied Automation permission or a wrong password still return `success:true` with `path` echoed back.
- **On macOS, `close` reports `success:true` when a document named `name` was found** and a save request was sent. If no open document has that name (or the Automation permission was denied), `success` is `false`.
- **On Windows, `success:true` means Office accepted the request.** If Office rejects it (for example the file to open does not exist, or the password is wrong), the COM call fails and the command returns `success:false`. That return path was added by the November 2026 review fixes (not in release 1.0.1); before, such a failure left the command without a return value.
- **On Windows, `close` can report `false` for a document it did close** (inferred from the code, not reproduced). After closing a match, the plugin keeps walking the document list with the count taken before the close; when two or more documents were open and the match was not the last one, Office rejects the out-of-range index and the command returns `success:false` although the document was saved and closed.
- **A POSIX path on macOS does nothing.** Use `File.platformPath`, which is an HFS path on macOS.
- **`close` needs the document name exactly as Office shows it**, including the extension and with the same case. On Windows, a new unsaved Word document has a localised name such as `"文書 1"`; Word can also rename a document it opened in repair mode, which the plugin turns off.
- **On Windows, `close` only works with a running app.** It never starts Word or Excel.
- **On Windows, the last `close` quits the app.** If the user had other work in that app with no documents open, the app is still closed.
- **On macOS, `close` without `path` does not update the original file.** It saves a copy into a new temporary folder; pass `path` to save somewhere else, including over the original.
- **On Windows, Word's `open` discards unsaved changes in a document that is already open** (`Revert` is `True`); on macOS the open window is activated instead.
- **A dialog in Office blocks 4D.** Password prompts, conversion warnings or sign-in windows keep the call waiting; run the command in a separate process if that matters.
- **`writePassword` is ignored by Excel.** For Excel on Windows, the write-reservation password comes from `templatePassword`, or from `password` if that is not set.

---

## Quick reference

```4d
// open
$status:=office automation(New object("app"; "word"; "command"; "open"; "path"; $file.platformPath))

// save to a file and close
$status:=office automation(New object("app"; "excel"; "command"; "close"; "name"; "book.xlsx"; "path"; $target.platformPath))

// save and close (Windows: in place; macOS: into a temporary folder)
$status:=office automation(New object("app"; "word"; "command"; "close"; "name"; "report.docx"))

If ($status.success)
	// $status.path
End if
```

---

## Developer notes

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
	pUnknown->Release();
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
