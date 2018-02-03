# Internet Explorer

Despite the efforts to deprecate the pertinent operating systems, IE is still
very popular, required for various government and corporate websites throughout
the world.  The modern download strategies are not available in older versions
of IE, but there are alternative approaches.

## Strategies

#### IE10 and IE11 File API

As part of the File API implementation, IE10 and IE11 provide the `msSaveBlob`
and `msSaveOrOpenBlob` functions to save blobs to the client computer.  This
approach is embedded in `XLSX.writeFile` and no additional shims are necessary.

#### Flash-based Download

It is possible to write to the file system using a SWF.  `Downloadify` library
implements one solution.  Since a genuine click is required, there is no way to
force a download.  The demo generates a button for each desired output format.

#### ActiveX-based Download

Through the `Scripting.FileSystemObject` object model, a script in the VBScript
scripting language can write to an arbitrary path on the filesystem.  The shim
includes a special `IE_SaveFile` function to write binary strings to file.  It
attempts to write to the Downloads folder or Documents folder or Desktop.

This approach can be triggered, but it requires the user to enable ActiveX.  It
is embedded as a strategy in `writeFile` and used only if the shim script is
included in the page and the relevant features are enabled on the target system.

## Demo

The included demo starts from an array of arrays, generating an editable HTML
table with `aoa_to_sheet` and adding it to the page:

```js
var ws = XLSX.utils.aoa_to_sheet(aoa);
var html_string = XLSX.utils.sheet_to_html(ws, { id: "table", editable: true });
document.getElementById("container").innerHTML = html_string;
```

The included download buttons use `table_to_book` to construct a new workbook
based on the table and `writeFile` to force a download:


```js
var elt = document.getElementById('table');
var wb = XLSX.utils.table_to_book(elt, { sheet: "Sheet JS" });
XLSX.writeFile(wb, filename);
```

The shim is included in the HTML page, unlocking the ActiveX pathway if enabled
in browser settings.

The corresponding SWF buttons are displayed in environments where Flash is
available and `Downloadify` is supported.  The easiest solution involves writing
to a Base64 string and passing to the library:

```js
Downloadify.create(element_id, {
	/* the demo includes the other options required by Downloadify */
	filename: "test.xlsx",
	data: function() { return XLSX.write(wb, {bookType:"xlsx", type:'base64'}); },
	dataType: 'base64'
});
```

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
