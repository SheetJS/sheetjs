function getdatastr(data)/*:?string*/ {
	if(!data) return null;
	if(data.content && data.type) return cc2str(data.content, true);
	if(data.data) return debom(data.data);
	if(data.asNodeBuffer && has_buf) return debom(data.asNodeBuffer().toString('binary'));
	if(data.asBinary) return debom(data.asBinary());
	if(data._data && data._data.getContent) return debom(cc2str(Array.prototype.slice.call(data._data.getContent(),0)));
	return null;
}

function getdatabin(data) {
	if(!data) return null;
	if(data.data) return char_codes(data.data);
	if(data.asNodeBuffer && has_buf) return data.asNodeBuffer();
	if(data._data && data._data.getContent) {
		var o = data._data.getContent();
		if(typeof o == "string") return char_codes(o);
		return Array.prototype.slice.call(o);
	}
	if(data.content && data.type) return data.content;
	return null;
}

function getdata(data) { return (data && data.name.slice(-4) === ".bin") ? getdatabin(data) : getdatastr(data); }

/* Part 2 Section 10.1.2 "Mapping Content Types" Names are case-insensitive */
/* OASIS does not comment on filename case sensitivity */
function safegetzipfile(zip, file/*:string*/) {
	var k = zip.FullPaths || keys(zip.files);
	var f = file.toLowerCase().replace(/[\/]/g, '\\'), g = f.replace(/\\/g,'\/');
	for(var i=0; i<k.length; ++i) {
		var n = k[i].replace(/^Root Entry[\/]/,"").toLowerCase();
		if(f == n || g == n) return zip.files ? zip.files[k[i]] : zip.FileIndex[i];
	}
	return null;
}

function getzipfile(zip, file/*:string*/) {
	var o = safegetzipfile(zip, file);
	if(o == null) throw new Error("Cannot find file " + file + " in zip");
	return o;
}

function getzipdata(zip, file/*:string*/, safe/*:?boolean*/)/*:any*/ {
	if(!safe) return getdata(getzipfile(zip, file));
	if(!file) return null;
	try { return getzipdata(zip, file); } catch(e) { return null; }
}

function getzipstr(zip, file/*:string*/, safe/*:?boolean*/)/*:?string*/ {
	if(!safe) return getdatastr(getzipfile(zip, file));
	if(!file) return null;
	try { return getzipstr(zip, file); } catch(e) { return null; }
}

function getzipbin(zip, file/*:string*/, safe/*:?boolean*/)/*:any*/ {
	if(!safe) return getdatabin(getzipfile(zip, file));
	if(!file) return null;
	try { return getzipbin(zip, file); } catch(e) { return null; }
}

function zipentries(zip) {
	var k = zip.FullPaths || keys(zip.files), o = [];
	for(var i = 0; i < k.length; ++i) if(k[i].slice(-1) != '/') o.push(k[i].replace(/^Root Entry[\/]/, ""));
	return o.sort();
}

function zip_add_file(zip, path, content) {
	if(zip.FullPaths) {
		if(typeof content == "string") {
			var res;
			if(has_buf) res = Buffer_from(content);
			/* TODO: investigate performance in Edge 13 */
			//else if(typeof TextEncoder !== "undefined") res = new TextEncoder().encode(content);
			else res = utf8decode(content);
			return CFB.utils.cfb_add(zip, path, res);
		}
		CFB.utils.cfb_add(zip, path, content);
	}
	else zip.file(path, content);
}

function zip_new() { return CFB.utils.cfb_new(); }

function zip_read(d, o) {
	switch(o.type) {
		case "base64": return CFB.read(d, { type: "base64" });
		case "binary": return CFB.read(d, { type: "binary" });
		case "buffer": case "array": return CFB.read(d, { type: "buffer" });
	}
	throw new Error("Unrecognized type " + o.type);
}

function resolve_path(path/*:string*/, base/*:string*/)/*:string*/ {
	if(path.charAt(0) == "/") return path.slice(1);
	var result = base.split('/');
	if(base.slice(-1) != "/") result.pop(); // folder path
	var target = path.split('/');
	while (target.length !== 0) {
		var step = target.shift();
		if (step === '..') result.pop();
		else if (step !== '.') result.push(step);
	}
	return result.join('/');
}
