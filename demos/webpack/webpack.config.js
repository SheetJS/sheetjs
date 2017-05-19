module.exports = {
	output: {
		libraryTarget: 'var',
		library: 'XLSX'
	},
	node: {
		fs: false,
		process: false,
		Buffer: false
	}
}
