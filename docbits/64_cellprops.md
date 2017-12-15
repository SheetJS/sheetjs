#### Hyperlinks

Hyperlinks are stored in the `l` key of cell objects.  The `Target` field of the
hyperlink object is the target of the link, including the URI fragment. Tooltips
are stored in the `Tooltip` field and are displayed when you move your mouse
over the text.

For example, the following snippet creates a link from cell `A3` to
<http://sheetjs.com> with the tip `"Find us @ SheetJS.com!"`:

```js
ws['A3'].l = { Target:"http://sheetjs.com", Tooltip:"Find us @ SheetJS.com!" };
```

Note that Excel does not automatically style hyperlinks -- they will generally
be displayed as normal text.

Links where the target is a cell or range or defined name in the same workbook
("Internal Links") are marked with a leading hash character:

```js
ws['A2'].l = { Target:"#E2" }; /* link to cell E2 */
```

