# NW.js

This library is compatible with NW.js and should just work out of the box.
The demonstration uses NW.js 0.24 with the dist script.

## Reading data

The standard HTML5 `FileReader` techniques from the browser apply to NW.js!
This demo includes a drag-and-drop box as well as a file input box, mirroring
the [SheetJS Data Preview Live Demo](http://oss.sheetjs.com/js-xlsx/).

## Writing data

File input elements with the attribute `nwsaveas` show UI for saving a file. The
standard trick is to generate a hidden file input DOM element and "click" it.
Since NW.js does not present a `writeFileSync` in the `fs` package, a manual
step is required:

```js
/* from within the input change callback, `this.value` is the file name */
var filename = this.value, bookType = (filename.match(/[^\.]*$/)||["xlsx"])[0];

/* convert the TABLE element back to a workbook */
var wb = XLSX.utils.table_to_book(HTMLOUT);

/* write to buffer */
var wbout = XLSX.write(wb, {type:'buffer', bookType:bookType});

/* use the async fs.writeFile to save the data */
fs.writeFile(filename, wbout, function(err) {
	if(!err) return alert("Saved to " + filename);
	alert("Error: " + (err.message || err));
});
```

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
