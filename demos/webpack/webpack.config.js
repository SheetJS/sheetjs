module.exports = {
	output: {
		libraryTarget: 'var',
		library: 'XLSX'
	},
	/* Uncomment the next block to suppress codepage */
	/*
	resolve: {
		alias: { "./dist/cpexcel.js": "" }
	},
	*/
	node: {
		fs: false,
		process: false,
		Buffer: false
	}
}
