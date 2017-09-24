# RequireJS

The minified dist files trip up the RequireJS mechanism.  To bypass, the scripts
automatically expose an `XLSX` variable that can be used if the require callback
argument is `_XLSX` rather than `XLSX`.  This trick is employed in the included
`xlsx-shim.js` script:

```js
/* xlsx-shim.js */
define(['xlsx'], function (_XLSX) {
	return XLSX;
});
```

The require config should set `xlsx` path to the appropriate dist file:

```js
	paths: {
		xlsx: "xlsx.full.min"
	},
```

Once that is set, app code can freely require `"xlsx-shim"`:

```js
require(["xlsx-shim"], function(XLSX) {
	/* use XLSX here */
});
```

## Deployments

`browser.html` demonstrates a dynamic deployment, using the in-browser config:

```html
<script src="require.js"></script>
<script>
require.config({
	baseUrl: ".",
	name: "app",
	paths: {
		xlsx: "xlsx.full.min"
	}
});
</script>
<script src="app.js"></script>
```

`optimizer.html` demonstrates an optimized deployment using `build.js` config:

```js
/* build config */
({
	baseUrl: ".",
	name: "app",
	paths: {
		xlsx: "xlsx.full.min"
	},
	out: "app-built.js"
})
```

The optimizer is invoked with:

```bash
node r.js -o build.js paths.requireLib=./require include=requireLib
```

That step creates a file `app-built.js` that can be included in a page:

```html
<!-- final bundle includes require.js, xlsx-shim, library and app code -->
<script src="app-built.js"></script>
```

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
