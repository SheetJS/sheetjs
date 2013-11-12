
if(typeof require !== 'undefined' && typeof exports !== 'undefined') {
	exports.read = XLSX.read;
	exports.readFile = XLSX.readFile;
	exports.utils = XLSX.utils;
	exports.main = function(args) {
		var zip = XLSX.read(args[0], {type:'file'});
		console.log(zip.Sheets);
	};
if(typeof module !== 'undefined' && require.main === module)
	exports.main(process.argv.slice(2));
}
