var SystemJS = require('systemjs');
SystemJS.config({
	meta: {
		'../../xlsx.js': { format: 'global' },
		'../../dist/xlsx.core.min.js': { format: 'global' },
		'../../dist/xlsx.full.min.js': { format: 'global' },
	},
	paths: {
		'npm:': '/usr/local/lib/node_modules/'
	},
	map: {
		'xlsx': 'npm:xlsx/xlsx.js',
		'fs': '@node/fs',
		'crypto': '@node/fs'
	}
});
SystemJS.import('./app.js');
