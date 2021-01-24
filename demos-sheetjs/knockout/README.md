# Knockout 

The `xlsx.core.min.js` and `xlsx.full.min.js` scripts are designed to be dropped
into web pages with script tags:

```html
<script src="xlsx.full.min.js"></script>
```

Strictly speaking, there should be no need for a Knockout demo!  You can proceed
as you would with any other browser-friendly library.


## Array of Arrays

A common data table is often stored as an array of arrays:

```js
var aoa = [ [1,2], [3,4] ];
```

This neatly maps to a table with `data-bind="foreach: ..."`:

```html
<table data-bind="foreach: aoa">
  <tr data-bind="foreach: $data">
    <td><span data-bind="text: $data"></span></td>
  </tr>
</table>
```

The `sheet_to_json` utility function can generate array of arrays for model use:

```js
/* starting from a `wb` workbook object, pull first worksheet */
var ws = wb.Sheets[wb.SheetNames[0]];
/* convert the worksheet to an array of arrays */
var aoa = XLSX.utils.sheet_to_json(ws, {header:1});
/* update model */
model.aoa(aoa);
```


## Demo

The easiest observable representation is an `observableArray`:

```js
var ViewModel = function() { this.aoa = ko.observableArray([[1,2],[3,4]]); };
var model = new ViewModel();
ko.applyBindings(model);
```

Unfortunately the nested `"foreach: $data"` binding is read-only.  A two-way
binding is possible using the `$parent` and `$index` binding context properties:

```html
<table data-bind="foreach: aoa">
  <tr data-bind="foreach: $data">
    <td><input data-bind="value: $parent[$index()]" /></td>
  </tr>
</table>
```

The demo shows reading worksheets into a view model and writing models to XLSX.


[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
