# Angular 1

The `xlsx.core.min.js` and `xlsx.full.min.js` scripts are designed to be dropped
into web pages with script tags e.g.

```html
<script src="xlsx.full.min.js"></script>
```

Strictly speaking, there should be no need for an angular demo!  You can proceed
as you would with any other browser-friendly library.  To make this meaningful,
we chose to show an integration with a common angular table component.

This demo uses angular-ui-grid to display a data table.  The ui-grid does not
provide any way to hook into the import button, so the demo includes a simple
directive for a HTML File Input control.  It also includes a sample service for
export which adds an item to the export menu.

## Import Directive

`SheetJSImportDirective` follows the prescription from the README for File input
controls using `readAsBinaryString`, converting to a suitable representation
and updating the scope.

## Export Service

`SheetJSExportService` exposes export functions for `XLSB` and `XLSX`.  Other
formats are easily supported by changing the `bookType` variable.  It grabs
values from the grid, builds an array of arrays, generates a workbook and uses
FileSaver to generate a download.  By setting the `filename` and `sheetname`
options in the ui-grid options, the output can be controlled.

