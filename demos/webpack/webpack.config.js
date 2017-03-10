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
		Buffer: false
	},
	externals: [
		{
			'./cptable': 'var cptable'
		}
	]
}
