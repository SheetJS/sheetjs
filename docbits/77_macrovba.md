#### VBA and Macros

VBA Macros are stored in a special data blob that is exposed in the `vbaraw`
property of the workbook object when the `bookVBA` option is `true`.  They are
supported in `XLSM`, `XLSB`, and `BIFF8 XLS` formats.  The supported format
writers automatically insert the data blobs if it is present in the workbook and
associate with the worksheet names.

<details>
	<summary><b>Custom Code Names</b> (click to show)</summary>

The workbook code name is stored in `wb.Workbook.WBProps.CodeName`.  By default,
Excel will write `ThisWorkbook` or a translated phrase like `DieseArbeitsmappe`.
Worksheet and Chartsheet code names are in the worksheet properties object at
`wb.Workbook.Sheets[i].CodeName`.  Macrosheets and Dialogsheets are ignored.

The readers and writers preserve the code names, but they have to be manually
set when adding a VBA blob to a different workbook.

</details>

<details>
	<summary><b>Macrosheets</b> (click to show)</summary>

Older versions of Excel also supported a non-VBA "macrosheet" sheet type that
stored automation commands.  These are exposed in objects with the `!type`
property set to `"macro"`.

</details>

<details>
	<summary><b>Detecting macros in workbooks</b> (click to show)</summary>

The `vbaraw` field will only be set if macros are present, so testing is simple:

```js
function wb_has_macro(wb/*:workbook*/)/*:boolean*/ {
	if(!!wb.vbaraw) return true;
	const sheets = wb.SheetNames.map((n) => wb.Sheets[n]);
	return sheets.some((ws) => !!ws && ws['!type']=='macro');
}
```

</details>

