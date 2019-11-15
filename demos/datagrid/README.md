# canvas-datagrid

The `sheet_to_json` utility function generates output arrays suitable for use
with other JS libraries such as data grids for previewing data.  After extensive
testing, [`canvas-datagrid`](https://tonygermaneri.github.io/canvas-datagrid/)
stood out as a very high-performance grid with an incredibly simple API.

This demo is available at <http://oss.sheetjs.com/js-xlsx/datagrid.html>

## Obtaining the Library

The [`canvas-datagrid` npm nodule](http://npm.im/canvas-datagrid) includes a
minified script `dist/canvas-datagrid.js` that can be directly inserted as a
script tag.  The unpkg CDN also exposes the latest version:

```html
<script src="https://unpkg.com/canvas-datagrid/dist/canvas-datagrid.js"></script>
```

## Previewing Data

The HTML document needs a container element:

```html
<div id="gridctr"></div>
```

Grid initialization is a one-liner:

```js
var grid = canvasDatagrid({
  parentNode: document.getElementById('gridctr'),
  data: []
});
```

For large data sets, it's necessary to constrain the size of the grid.

```js
grid.style.height = '100%';
grid.style.width = '100%';
```

Once the workbook is read and the worksheet is selected, assigning the data
variable automatically updates the view:

```js
grid.data = XLSX.utils.sheet_to_json(ws, {header:1});
```

This demo previews the first worksheet.

## Editing

`canvas-datagrid` handles the entire edit cycle.  No intervention is necessary.

## Saving Data

`grid.data` is immediately readable and can be converted back to a worksheet.
Some versions return an array-like object without the length, so a little bit of
preparation may be needed:

```js
/* converts an array of array-like objects into an array of arrays */
function prep(arr) {
  var out = [];
  for(var i = 0; i < arr.length; ++i) {
    if(!arr[i]) continue;
    if(Array.isArray(arr[i])) { out[i] = arr[i]; continue };
    var o = new Array();
    Object.keys(arr[i]).forEach(function(k) { o[+k] = arr[i][k] });
    out[i] = o;
  }
  return out;
}

/* build worksheet from the grid data */
var ws = XLSX.utils.aoa_to_sheet(prep(grid.data));

/* build up workbook */
var wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'SheetJS');

/* generate download */
XLSX.writeFile(wb, "SheetJS.xlsx");
```

## Additional Features

This demo barely scratches the surface.  The underlying grid component includes
many additional features including massive data streaming, sorting and styling.

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
