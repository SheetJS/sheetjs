/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
import resolve from '@rollup/plugin-node-resolve';
export default {
	input: 'app.js',
	output: {
		file: 'rollup.js',
		format: 'iife'
	},
	//dest: 'rollup.js',
	plugins: [
		resolve({
			module: false,
			browser: true,
		}),
	],
};
