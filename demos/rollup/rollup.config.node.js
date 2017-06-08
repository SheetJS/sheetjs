/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
import resolve from 'rollup-plugin-node-resolve';
import commonjs from 'rollup-plugin-commonjs';
export default {
	entry: 'main.js',
	dest: 'rollup.node.js',
	plugins: [
		resolve({
			module: false,
			browser: true,
		}),
		commonjs()
	],
	format: 'cjs'
};
