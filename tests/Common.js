var XLSX = require('../');

var tests = {
	'should be able to open workbook': function (file) {
		var xlsx = XLSX.readFile('tests/files/' + file);
		expect(xlsx).toBeTruthy();
		expect(xlsx).toEqual(jasmine.any(Object));
	},
	'should define all api properties correctly': function (file) {
		var xlsx = XLSX.readFile('tests/files/' + file);
		expect(xlsx.Workbook).toEqual(jasmine.any(Object));
		expect(xlsx.Props).toBeDefined();
		expect(xlsx.Deps).toBeDefined();
		expect(xlsx.Sheets).toEqual(jasmine.any(Object));
		expect(xlsx.SheetNames).toEqual(jasmine.any(Array));
		expect(xlsx.Strings).toBeDefined();
		expect(xlsx.Styles).toBeDefined();
	}
};

module.exports = function (file) {
	for (var key in tests) {
		it(key, tests[key].bind(undefined, file));
	}
};