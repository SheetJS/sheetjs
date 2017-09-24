# Angular 2+

The library can be imported directly from TS code with:

```typescript
import * as XLSX from 'xlsx';
```

This demo uses an array of arrays (type `Array<Array<any>>`) as the core state.
The component template includes a file input element, a table that updates with
the data, and a button to export the data.

## Array of Arrays

`Array<Array<any>>` neatly maps to a table with `ngFor`:

```html
<table class="sjs-table">
  <tr *ngFor="let row of data">
    <td *ngFor="let val of row">
      {{val}}
    </td>
  </tr>
</table>
```

The `aoa_to_sheet` utility function returns a worksheet.  Exporting is simple:

```typescript
/* generate worksheet */
const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);

/* generate workbook and add the worksheet */
const wb: XLSX.WorkBook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

/* save to file */
const wbout: string = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
saveAs(new Blob([s2ab(wbout)]), 'SheetJS.xlsx');
```

`sheet_to_json` with the option `header:1` makes importing simple:

```typescript
/* <input type="file" (change)="onFileChange($event)" multiple="false" /> */
/* ... (within the component class definition) ... */
  onFileChange(evt: any) {
    /* wire up file reader */
    const target: DataTransfer = <DataTransfer>(evt.target);
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      /* save data */
      this.data = <AOA>(XLSX.utils.sheet_to_json(ws, {header: 1}));
    };
    reader.readAsBinaryString(target.files[0]);
  }
```

## Switching between Angular versions

Modules that work with Angular 2 largely work as-is with Angular 4.  Switching
between versions is mostly a matter of installing the correct version of the
core and associated modules.  This demo includes a `package.json` for Angular 2
and another `package.json` for Angular 4.

Switching to Angular 2 is as simple as:

```bash
$ cp package.json-angular2 package.json
$ npm install
$ ng serve
```

Switching to Angular 4 is as simple as:

```bash
$ cp package.json-angular4 package.json
$ npm install
$ ng serve
```

## XLSX Symbolic Link

In this tree, `node_modules/xlsx` is a link pointing back to the root.  This
enables testing the development version of the library.  In order to use this
demo in other applications, add the `xlsx` dependency:

```bash
$ npm install --save xlsx
```

## SystemJS Configuration

The default angular-cli configuration requires no additional configuration.

Some deployments use the SystemJS loader, which does require configuration.  The
SystemJS example shows the required meta and map settings:

```js
SystemJS.config({
  meta: {
    'xlsx': {
      exports: 'XLSX' // <-- tell SystemJS to expose the XLSX variable
    }
  },
  map: {
    'xlsx': 'xlsx.full.min.js', // <-- make sure xlsx.full.min.js is in same dir
    'fs': '',     // <--|
    'crypto': '', // <--| suppress native node modules
    'stream': ''  // <--|
  }
});
```

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
