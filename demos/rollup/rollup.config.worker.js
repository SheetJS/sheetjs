/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
import resolve from '@rollup/plugin-node-resolve';
export default {
	input: 'xlsxworker.js',
	output: {
		file: 'worker.js',
		format: 'iife'
	},
	//dest: 'worker.js',
	plugins: [
		resolve({
			module: false,
			browser: true,
		}),
	],
};
