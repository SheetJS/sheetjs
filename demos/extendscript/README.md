# ExtendScript demos

The main file is `test.jsx`.  Target-specific files prepend target directives.

Copy the `test.jsx` file as well as the `shim.js` and `xlsx.core.min.js` files
to wherever you want the scripts to reside.  The demo shows opening a file and
converting to an array of arrays.

NOTE: [We forked the minifier](https://www.npmjs.com/package/@sheetjs/uglify-js)
and included a bugfix for ExtendScript's misparsing of switch statements.

