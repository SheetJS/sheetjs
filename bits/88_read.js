function readSync(data, options) {
	var zip, d = data;
	var o = options||{};
	switch((o.type||"base64")){
		case "file":
			if(typeof Buffer !== 'undefined') { zip=new jszip(d=_fs.readFileSync(data)); break; }
			d = _fs.readFileSync(data).toString('base64');
			/* falls through */
		case "base64": zip = new jszip(d, { base64:true }); break;
		case "binary": zip = new jszip(d, { base64:false }); break;
	}
	return parseZip(zip);
}

function readFileSync(data, options) {
	var o = options||{}; o.type = 'file';
	return readSync(data, o);
}

XLSX.read = readSync;
XLSX.readFile = readFileSync;
XLSX.parseZip = parseZip;
