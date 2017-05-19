# Webpack

This library is built with some dynamic logic to determine if it is invoked in a
script tag or in nodejs.  Webpack does not understand those feature tests, so by
default it will do some strange things.

## Suppressing the Node shims

The library properly guards against accidental leakage of node features in the
browser but webpack disregards those.  The config should explicitly suppress:

```js
	node: {
		fs: false,
		process: false,
		Buffer: false
	}
```

## Exporting the XLSX variable

This library will not assign to module.exports if it is run in the browser.  To
convince webpack, set `output` in the webpack config:

```js
	output: {
		libraryTarget: 'var',
		library: 'XLSX'
	}
```
