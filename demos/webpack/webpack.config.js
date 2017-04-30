module.exports = {
	output: {
		libraryTarget: 'var',
		library: 'XLSX'
	},
	module: {
		noParse: [/jszip.js$/]
	},
	node: {
		fs: false,
		process: false,
		Buffer: false
	}
}
