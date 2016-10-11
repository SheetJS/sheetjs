var X = require('../xlsx.js');
var file = 'test_files/2013/apachepoi_44861.xls.xlsb';
var file = 'test_files/apachepoi_44861.xls';
var opts = {cellNF: true};
describe('from 2013', function() {
		it('should parse', function() {
			var wb = X.readFile(file, opts);
		});
	});

