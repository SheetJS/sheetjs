# x-spreadsheet

The `sheet_to_json` utility function generates output arrays suitable for use
with other JS libraries such as data grids for previewing data.  With a familiar
UI, [`x-spreadsheet`](https://myliang.github.io/x-spreadsheet/) is an excellent
choice for developers looking for a modern editor.

This demo is available at <https://oss.sheetjs.com/sheetjs/x-spreadsheet.html>

## Obtaining the Library

The [`x-data-spreadsheet` module](http://npm.im/x-data-spreadsheet) includes a
minified script `dist/xspreadsheet.js` that can be directly inserted as a script
tag.  The unpkg CDN also exposes the latest version:

```html
<script src="https://unpkg.com/x-data-spreadsheet/dist/xspreadsheet.js"></script>
```

## Previewing Data

The HTML document needs a container element:

```html
<div id="gridctr"></div>
```

Grid initialization is a one-liner:

```js
/* note that the browser build exposes the variable `x` */
var grid = x_spreadsheet(document.getElementById("gridctr"));
```

The following function converts data from SheetJS to x-spreadsheet:

```js
/* load data */
grid.loadData(stox(workbook_object));
```

`stox` is defined in [`xlsxspread.js`](./xlsxspread.js)

## Editing

`x-spreadsheet` handles the entire edit cycle. No intervention is necessary.

## Saving Data

`grid.getData()` returns an object that can be converted back to a worksheet:

```js
/* build workbook from the grid data */
var new_wb = xtos(xspr.getData());

/* generate download */
XLSX.writeFile(new_wb, "SheetJS.xlsx");
```

`stox` is defined in [`xlsxspread.js`](./xlsxspread.js)

## Additional Features

This demo barely scratches the surface.  The underlying grid component includes
many additional features that work with [SheetJS Pro](https://sheetjs.com/pro).

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
