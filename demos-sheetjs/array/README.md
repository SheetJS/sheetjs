# Typed Arrays and Math

ECMAScript version 6 introduced Typed Arrays, array-like objects designed for
low-level optimizations and predictable operations.  They are supported in most
modern browsers and form the basis of various APIs, including NodeJS Buffers,
WebGL buffers, WebAssembly, and tensors in linear algebra and math libraries.

This demo covers conversions between worksheets and Typed Arrays.  It also tries
to cover common numerical libraries that work with data arrays.

Excel supports a subset of the IEEE754 Double precision floating point numbers,
but many libraries only support `Float32` Single precision values. `Math.fround`
rounds `Number` values to the nearest single-precision floating point value.

## Working with Data in Typed Arrays

Typed arrays are not true Array objects.  The array of array utility functions
like `aoa_to_sheet` will not handle arrays of Typed Arrays.

#### Exporting Typed Arrays to a Worksheet

A single typed array can be converted to a pure JS array with `Array.from`:

```js
var column = Array.from(dataset_typedarray);
```

`aoa_to_sheet` expects a row-major array of arrays.  To export multiple data
sets, "transpose" the data:

```js
/* assuming data is an array of typed arrays */
var aoa = [];
for(var i = 0; i < data.length; ++i) {
  for(var j = 0; j < data[i].length; ++j) {
    if(!aoa[j]) aoa[j] = [];
    aoa[j][i] = data[i][j];
  }
}
/* aoa can be directly converted to a worksheet object */
var ws = XLSX.utils.aoa_to_sheet(aoa);
```

#### Importing Data from a Spreadsheet

`sheet_to_json` with the option `header:1` will generate a row-major array of
arrays that can be transposed.  However, it is more efficient to walk the sheet
manually:

```js
/* find worksheet range */
var range = XLSX.utils.decode_range(ws['!ref']);
var out = []
/* walk the columns */
for(var C = range.s.c; C <= range.e.c; ++C) {
  /* create the typed array */
  var ta = new Float32Array(range.e.r - range.s.r + 1);
  /* walk the rows */
  for(var R = range.s.r; R <= range.e.r; ++R) {
    /* find the cell, skip it if the cell isn't numeric or boolean */
    var cell = ws[XLSX.utils.encode_cell({r:R, c:C})];
    if(!cell || cell.t != 'n' && cell.t != 'b') continue;
    /* assign to the typed array */
    ta[R - range.s.r] = cell.v;
  }
  out.push(ta);
}
```

If the data set has a header row, the loop can be adjusted to skip those rows.


## Demos

Each example focuses on single-variable linear regression.  Sample worksheets
will start with a label row.  The first column is the x-value and the second
column is the y-value.  A sample spreadsheet can be generated randomly:

```js
var aoo = [];
for(var i = 0; i < 100; ++i) aoo.push({x:i, y:2 * i + Math.random()});
var ws = XLSX.utils.json_to_sheet(aoo);
var wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
XLSX.writeFile(wb, "linreg.xlsx");
```

Some libraries provide utility functions that work with plain arrays of numbers.
When possible, they should be preferred over manual conversion.

Reshaping raw float arrays and exporting to a worksheet is straightforward:

```js
function array_to_sheet(farray, shape, headers) {
  /* generate new AOA from the float array */
  var aoa = [];
  for(var j = 0; j < shape[0]; ++j) {
    aoa[j] = [];
    for(var i = 0; i < shape[1]; ++i) aoa[j][i] = farray[j * shape[1] + i];
  }

  /* add headers and generate worksheet */
  if(headers) aoa.unshift(headers);
  return XLSX.utils.aoa_to_sheet(aoa);
}
```

#### Tensor Operations with Propel ML

[Propel ML](http://propelml.org/) `tensor` objects can be transposed:

```js
var tensor = pr.tensor(aoa).transpose();
var col1 = tensor.slice(0, 1);
var col2 = tensor.slice(1, 1);
```

To export to a worksheet, `dataSync` generates a `Float32Array` that can be
re-shaped in JS:

```js
/* extract shape and float array */
var tensor = pr.concat([col1, col2]).transpose();
var shape = tensor.shape;
var farray = tensor.dataSync();
var ws = array_to_sheet(farray, shape, ["header1", "header2"]);
```

The demo generates a sample dataset and uses Propel to calculate the OLS linear
regression coefficients.  Afterwards, the tensors are exported to a new file.

#### TensorFlow

[TensorFlow](https://js.tensorflow.org/) `tensor` objects can be created from
arrays of arrays:

```js
var tensor = tf.tensor2d(aoa).transpose();
var col1 = tensor.slice([0,0], [1,tensor.shape[1]]).flatten();
var col2 = tensor.slice([1,0], [1,tensor.shape[1]]).flatten();
```

`stack` should be used to create the 2-d tensor for export:

```js
var tensor = tf.stack([col1, col2]).transpose();
var shape = tensor.shape;
var farray = tensor.dataSync();
var ws = array_to_sheet(farray, shape, ["header1", "header2"]);
```

The demo generates a sample dataset and uses a simple linear predictor with
least-squares scoring to calculate regression coefficients.  The tensors are
exported to a new file.

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
