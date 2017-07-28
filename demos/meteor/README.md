# Meteor

This library is universal: outside of environment-specific features (parsing DOM
tables in the browser, streaming write in nodejs), the core is ES3/ES5 and can
be used in any reasonably compliant JS implementation.  It should play nice with
meteor out of the box.


## This demonstration

You can split the work between the client and server side as you see fit.  The
obvious extremes of pure-client code and pure-server code are straightforward.
This demo tries to split the work to demonstrate that the workbook object can be
passed on the wire.

The read demo:
- accepts files from the client side
- sends binary string to server
- processes data on server side
- sends workbook object to client
- renders HTML and adds to a DOM element

The write demo:
- generates workbook on server side
- sends workbook object to client
- generates file on client side
- triggers a download.

This demo uses the FileSaver.js library for writing files, installed through the
[`pfafman:filesaver` wrapper](https://atmospherejs.com/pfafman/filesaver):

```bash
meteor add pfafman:filesaver
```

## Setup

This tree does not include the `.meteor` structure.  Rebuild the project with:

```bash
meteor create .
npm install babel-runtime meteor-node-stubs xlsx
meteor add pfafman:filesaver
meteor
```


## Environment-specific features

File-related operations (e.g. `XLSX.readFile` and `XLSX.writeFile`) will not be
available in client-side code. If you need to read a local file from the client,
use a file input or drag-and-drop.

Browser-specific operations (e.g. `XLSX.utils.table_to_book`) are limited to
client side code. You should never have to read from DOM elements on the server
side, but you can use a third-party virtual DOM to provide the required API.
