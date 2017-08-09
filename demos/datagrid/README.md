# canvas-datagrid

The `sheet_to_json` utility function generates output arrays suitable for use
with other JS libraries such as data grids for previewing data.  After extensive
testing, [`canvas-datagrid`](https://tonygermaneri.github.io/canvas-datagrid/)
stood out as a very high-performance grid with an incredibly simple API.

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

Once the workbook is read and the worksheet is selected, assigning the data
variable automatically updates the view:

```js
grid.data = XLSX.utils.sheet_to_json(ws, {header:1});
```

This demo previews the first worksheet, but it is easy to add buttons and other
features to support multiple worksheets.

## Editing

The library handles the whole edit cycle.  No intervention is necessary.

## Saving Data

`grid.data` is immediately readable and can be converted back to a worksheet:

```js
/* build worksheet from the grid data */
var ws = XLSX.utils.aoa_to_sheet(grid.data);

/* build up workbook */
var wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'SheetJS');

/* .. generate download (see documentation for examples) .. */
```

## Additional Features

This demo barely scratches the surface.  The underlying grid component includes
many additional features including massive data streaming, sorting and styling.
