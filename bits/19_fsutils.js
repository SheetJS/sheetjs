var _fs;
if(typeof require !== 'undefined') try { _fs = require('fs'); } catch(e) {}

/* normalize data for blob ctor */
function blobify(data) {
	if(typeof data === "string") return s2ab(data);
	if(Array.isArray(data)) return a2u(data);
	return data;
}
/* write or download file */
function write_dl(fname/*:string*/, payload/*:any*/, enc/*:?string*/) {
	/*global IE_SaveFile, Blob, navigator, saveAs, URL, document */
	if(typeof _fs !== 'undefined' && _fs.writeFileSync) return enc ? _fs.writeFileSync(fname, payload, enc) : _fs.writeFileSync(fname, payload);
	var data = (enc == "utf8") ? utf8write(payload) : payload;
	/*:: declare var IE_SaveFile: any; */
	if(typeof IE_SaveFile !== 'undefined') return IE_SaveFile(data, fname);
	if(typeof Blob !== 'undefined') {
		var blob = new Blob([blobify(data)], {type:"application/octet-stream"});
		/*:: declare var navigator: any; */
		if(typeof navigator !== 'undefined' && navigator.msSaveBlob) return navigator.msSaveBlob(blob, fname);
		/*:: declare var saveAs: any; */
		if(typeof saveAs !== 'undefined') return saveAs(blob, fname);
		if(typeof URL !== 'undefined' && typeof document !== 'undefined' && document.createElement && URL.createObjectURL) {
			var a = document.createElement("a");
			if(a.download != null) {
				var url = URL.createObjectURL(blob);
				/*:: if(document.body == null) throw new Error("unreachable"); */
				a.download = fname; a.href = url; document.body.appendChild(a); a.click();
				/*:: if(document.body == null) throw new Error("unreachable"); */ document.body.removeChild(a);
				if(URL.revokeObjectURL && typeof setTimeout !== 'undefined') setTimeout(function() { URL.revokeObjectURL(url); }, 60000);
				return url;
			}
		}
	}
	throw new Error("cannot initiate download");
}

