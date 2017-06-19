# VueJS 2

The `xlsx.core.min.js` and `xlsx.full.min.js` scripts are designed to be dropped
into web pages with script tags e.g.

```html
<script src="xlsx.full.min.js"></script>
```

Strictly speaking, there should be no need for a Vue.JS demo!  You can proceed
as you would with any other browser-friendly library.

This demo directly generates HTML using `sheet_to_html` and adds an element to
a pregenerated template.  It also has a button for exporting as XLSX.


## Single File Components

For Single File Components, a simple `import XLSX from 'xlsx'` should suffice.
The webpack demo includes a sample `webpack.config.js`.
