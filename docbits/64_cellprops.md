#### Hyperlinks

<details>
  <summary><b>Format Support</b> (click to show)</summary>

**Cell Hyperlinks**: XLSX/M, XLSB, BIFF8 XLS, XLML, ODS

**Tooltips**: XLSX/M, XLSB, BIFF8 XLS, XLML

</details>

Hyperlinks are stored in the `l` key of cell objects.  The `Target` field of the
hyperlink object is the target of the link, including the URI fragment. Tooltips
are stored in the `Tooltip` field and are displayed when you move your mouse
over the text.

For example, the following snippet creates a link from cell `A3` to
<https://sheetjs.com> with the tip `"Find us @ SheetJS.com!"`:

```js
ws['A1'].l = { Target:"https://sheetjs.com", Tooltip:"Find us @ SheetJS.com!" };
```

Note that Excel does not automatically style hyperlinks -- they will generally
be displayed as normal text.

_Remote Links_

HTTP / HTTPS links can be used directly:

```js
ws['A2'].l = { Target:"https://docs.sheetjs.com/#hyperlinks" };
ws['A3'].l = { Target:"http://localhost:7262/yes_localhost_works" };
```

Excel also supports `mailto` email links with subject line:

```js
ws['A4'].l = { Target:"mailto:ignored@dev.null" };
ws['A5'].l = { Target:"mailto:ignored@dev.null?subject=Test Subject" };
```

_Local Links_

Links to absolute paths should use the `file://` URI scheme:

```js
ws['B1'].l = { Target:"file:///SheetJS/t.xlsx" }; /* Link to /SheetJS/t.xlsx */
ws['B2'].l = { Target:"file:///c:/SheetJS.xlsx" }; /* Link to c:\SheetJS.xlsx */
```

Links to relative paths can be specified without a scheme:

```js
ws['B3'].l = { Target:"SheetJS.xlsb" }; /* Link to SheetJS.xlsb */
ws['B4'].l = { Target:"../SheetJS.xlsm" }; /* Link to ../SheetJS.xlsm */
```

Relative Paths have undefined behavior in the SpreadsheetML 2003 format.  Excel
2019 will treat a `..\` parent mark as two levels up.

_Internal Links_

Links where the target is a cell or range or defined name in the same workbook
("Internal Links") are marked with a leading hash character:

```js
ws['C1'].l = { Target:"#E2" }; /* Link to cell E2 */
ws['C2'].l = { Target:"#Sheet2!E2" }; /* Link to cell E2 in sheet Sheet2 */
ws['C3'].l = { Target:"#SomeDefinedName" }; /* Link to Defined Name */
```

