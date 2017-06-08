# Rollup

This library presents itself as a CommonJS library, so some configuration is
required.  The examples at <https://rollupjs.org> can be followed pretty much in
verbatim.  This sample demonstrates a rollup for browser as well as for node.

## Required Plugins

The `rollup-plugin-node-resolve` and `rollup-plugin-commonjs` plugins are used:

```js
import resolve from 'rollup-plugin-node-resolve';
import commonjs from 'rollup-plugin-commonjs';
export default {
	/* ... */
	plugins: [
		resolve({
			module: false, // <-- this library is not an ES6 module
			browser: true, // <-- suppress node-specific features
		}),
		commonjs()
	],
	/* ... */
};
```

