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
var grid = x.spreadsheet(document.getElementById("gridctr"));
```

The following function converts data from SheetJS to x-spreadsheet:

```js
function stox(wb) {
  var out = [];
  wb.SheetNames.forEach(function(name) {
    var o = {name:name, rows:{}};
    var ws = wb.Sheets[name];
    var aoa = XLSX.utils.sheet_to_json(ws, {raw: false, header:1});
    aoa.forEach(function(r, i) {
      var cells = {};
      r.forEach(function(c, j) { cells[j] = ({ text: c }); });
      o.rows[i] = { cells: cells };
    })
    out.push(o);
  });
  return out;
}

/* load data */
grid.loadData(stox(workbook_object));
```

## Editing

`x-spreadsheet` handles the entire edit cycle.  No intervention is necessary.

## Saving Data

`grid.getData()` returns an object that can be converted back to a worksheet:

```js
function xtos(sdata) {
  var out = XLSX.utils.book_new();
  sdata.forEach(function(xws) {
    var aoa = [[]];
    var rowobj = xws.rows;
    for(var ri = 0; ri < rowobj.len; ++ri) {
      var row = rowobj[ri];
      if(!row) continue;
      aoa[ri] = [];
      Object.keys(row.cells).forEach(function(k) {
        var idx = +k;
        if(isNaN(idx)) return;
        aoa[ri][idx] = row.cells[k].text;
      });
    }
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(out, ws, xws.name);
  });
  return out;
}

/* build workbook from the grid data */
var new_wb = xtos(xspr.getData());

/* generate download */
XLSX.writeFile(new_wb, "SheetJS.xlsx");
```

## Additional Features

This demo barely scratches the surface.  The underlying grid component includes
many additional features that work with [SheetJS Pro](https://sheetjs.com/pro).

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
