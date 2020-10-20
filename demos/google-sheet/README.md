# Google Sheet Demo

This demo is using [`drive-db`](https://github.com/franciscop/drive-db) to fetch a public Google Sheet and then `xlsx` to save the data locally as `test.xlsx`.

It uses modern Javascript; `import/export`, `async/away`, etc. To run this you need Node.js 12 or newer, and you will notice we added `"type": "module"` to the `package.json`.

Here is the full code:

```js
import xlsx from "xlsx";
import drive from "drive-db";

(async () => {
  const data = await drive("1fvz34wY6phWDJsuIneqvOoZRPfo6CfJyPg1BYgHt59k");

  /* Create a new workbook */
  const workbook = xlsx.utils.book_new();

  /* make worksheet */
  const worksheet = xlsx.utils.json_to_sheet(data);

  /* Add the worksheet to the workbook */
  xlsx.utils.book_append_sheet(workbook, worksheet);

  xlsx.writeFile(workbook, "test.xlsx");
})();
```

Let's go over the different parts:

```js
import xlsx from "xlsx";
import drive from "drive-db";
```

This imports both `xlsx` and `drive-db` libraries. While these are written in commonjs, Javascript Modules can usually import the commonjs modules with no problem.

```js
(async () => {
  // ...
})();
```

This is what is called an [Immediately Invoked Function Expression](https://flaviocopes.com/javascript-iife/). These are normally used to either create a new execution context, or in this case to allow to run async code easier.

```js
const data = await drive("1fvz34wY6phWDJsuIneqvOoZRPfo6CfJyPg1BYgHt59k");
```

Using `drive-db`, fetch the data for the given spreadsheet id. In this case it's [this Google Sheet document](https://docs.google.com/spreadsheets/d/1fvz34wY6phWDJsuIneqvOoZRPfo6CfJyPg1BYgHt59k/edit), and since we don't specify the sheet it's the default one.

```js
const workbook = xlsx.utils.book_new();
const worksheet = xlsx.utils.json_to_sheet(data);
```

We need to create a workbook with a worksheet inside. The worksheet is created from the previously fetched data. `drive-db` exports the data in the same format that `xlsx`'s `.json_to_sheet()` method expects, so it's a straightforward operation.

```js
xlsx.utils.book_append_sheet(workbook, worksheet);
```

The worksheet needs to be inside the workbook, so we use the operation `.book_append_sheet()` to make it so.

```js
xlsx.writeFile(workbook, "test.xlsx");
```

Finally we save the workbook into a XLSX file in out filesystem. With this, now it can be opened by any spreadsheet program that we have installed.

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
