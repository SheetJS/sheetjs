# AngularJS

The `xlsx.core.min.js` and `xlsx.full.min.js` scripts are designed to be dropped
into web pages with script tags:

```html
<script src="xlsx.full.min.js"></script>
```

Strictly speaking, there should be no need for an Angular demo!  You can proceed
as you would with any other browser-friendly library.


## Array of Objects

A common data table is often stored as an array of objects:

```js
$scope.data = [
  { Name: "Bill Clinton", Index: 42 },
  { Name: "GeorgeW Bush", Index: 43 },
  { Name: "Barack Obama", Index: 44 },
  { Name: "Donald Trump", Index: 45 }
];
```

This neatly maps to a table with `ng-repeat`:

```html
<table id="sjs-table">
  <tr><th>Name</th><th>Index</th></tr>
  <tr ng-repeat="row in data">
    <td>{{row.Name}}</td>
    <td>{{row.Index}}</td>
  </tr>
</table>
```

The `$http` service can request binary data using the `"arraybuffer"` response
type coupled with `XLSX.read` with type `"array"`:

```js
  $http({
    method:'GET',
    url:'https://sheetjs.com/pres.xlsx',
    responseType:'arraybuffer'
  }).then(function(data) {
    var wb = XLSX.read(data.data, {type:"array"});
    var d = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
    $scope.data = d;
  }, function(err) { console.log(err); });
```

The HTML table can be directly exported with `XLSX.utils.table_to_book`:

```js
var wb = XLSX.utils.table_to_book(document.getElementById('sjs-table'));
XLSX.writeFile(wb, "export.xlsx");
```


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
function SheetJSImportDirective() {
  return {
    scope: { opts: '=' },
    link: function ($scope, $elm) {
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
}
```


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

/* write workbook and force a download */
XLSX.writeFile(wb, "sheetjs.xlsx");
```

## Demo

`grid.html` uses `angular-ui-grid` to display a table.  The library does not
provide any way to modify the import button, so the demo includes a simple
directive for a HTML File Input control.  It also includes a sample service for
export which adds an item to the export menu.

The demo `SheetJSImportDirective` follows the prescription from the README for
File input controls using `readAsBinaryString`, converting to a suitable
representation and updating the scope.

`SheetJSExportService` exposes export functions for `XLSB` and `XLSX`.  Other
formats are easily supported by changing the `bookType` variable.  It grabs
values from the grid, builds an array of arrays, generates a workbook and forces
a download.  By setting the `filename` and `sheetname` options in the `ui-grid`
options, the output can be controlled.


[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
