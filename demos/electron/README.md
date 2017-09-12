# Electron

This library is compatible with Electron and should just work out of the box.
The demonstration uses Electron v1.7.5.  The library is added via `require` from
the render process.  It can also be required from the main process, as shown in
this demo to render a version string in the About dialog on OSX.

The standard HTML5 `FileReader` techniques from the browser apply to Electron.
This demo includes a drag-and-drop box as well as a file input box, mirroring
the [SheetJS Data Preview Live Demo](http://oss.sheetjs.com/js-xlsx/)

Since electron provides an `fs` implementation, `readFile` and `writeFile` can
be used in conjunction with the standard dialogs.  For example:

```js
var dialog = require('electron').remote.dialog;
var o = (dialog.showOpenDialog({ properties: ['openFile'] })||[''])[0];
var workbook = X.readFile(o);
```
