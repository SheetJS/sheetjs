#### VBA and Macros

VBA Macros are stored in a special data blob that is exposed in the `vbaraw`
property of the workbook object when the `bookVBA` option is `true`.  They are
supported in `XLSM`, `XLSB`, and `BIFF8 XLS` formats.  The `XLSM` and `XLSB`
writers automatically insert the data blobs if it is present in the workbook.

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

