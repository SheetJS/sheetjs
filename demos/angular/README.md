# Angular 1

The `xlsx.core.min.js` and `xlsx.full.min.js` scripts are designed to be dropped
into web pages with script tags:

```html
<script src="xlsx.full.min.js"></script>
```

Strictly speaking, there should be no need for an angular demo!  You can proceed
as you would with any other browser-friendly library.  To make this meaningful,
we chose to show an integration with a common angular table component.

This demo uses angular-ui-grid to display a data table.  The ui-grid does not
provide any way to modify the import button, so the demo includes a simple
directive for a HTML File Input control.  It also includes a sample service for
export which adds an item to the export menu.

## Import Directive

A general import directive is fairly straightforward:

- Define the `importSheetJs` directive in the app:

```js
app.directive("importSheetJs", [SheetJSImportDirective]);
```

- Add the attribute `import-sheet-js=""` to the file input element:

```html
<input type="file" import-sheet-js="" multiple="false"  />
```

- Define the directive:

```js
var SheetJSImportDirective = function() {
  return {
    scope: { },
    link: function ($scope, $elm, $attrs) {
      $elm.on('change', function (changeEvent) {
        var reader = new FileReader();

        reader.onload = function (e) {
          /* read workbook */
          var bstr = e.target.result;
          var workbook = XLSX.read(bstr, {type:'binary'});

          /* DO SOMETHING WITH workbook HERE */
        };

        reader.readAsBinaryString(changeEvent.target.files[0]);
      });
    }
  };
};
```

The demo `SheetJSImportDirective` follows the prescription from the README for
File input controls using `readAsBinaryString`, converting to a suitable
representation and updating the scope.

## Export Service

An export can be triggered at any point!  Depending on how data is represented,
a workbook object can be built using the utility functions.  For example, using
an array of objects:

```js
/* starting from this data */
var data = [
  { name: "Barack Obama", pres: 44 },
  { name: "Donald Trump", pres: 45 }
];

/* generate a worksheet */
var ws = XLSX.utils.json_to_sheet(data);

/* add to workbook */
var wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "Presidents");

/* write workbook (use type 'array' for ArrayBuffer) */
var wbout = XLSX.write(wb, {bookType:'xlsx', type:'array'});

/* generate a download */
saveAs(new Blob([wbout],{type:"application/octet-stream"}), "sheetjs.xlsx");
```


`SheetJSExportService` exposes export functions for `XLSB` and `XLSX`.  Other
formats are easily supported by changing the `bookType` variable.  It grabs
values from the grid, builds an array of arrays, generates a workbook and uses
FileSaver to generate a download.  By setting the `filename` and `sheetname`
options in the ui-grid options, the output can be controlled.

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
