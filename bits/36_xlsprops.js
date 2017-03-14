function xlml_set_prop(Props, tag/*:string*/, val) {
	/* TODO: Normalize the properties */
	switch(tag) {
		case 'Description': tag = 'Comments'; break;
		case 'Created': tag = 'CreatedDate'; break;
		case 'LastSaved': tag = 'ModifiedDate'; break;
	}
	Props[tag] = val;
}

var XLMLDocumentProperties = [
	['Title', 'Title'],
	['Subject', 'Subject'],
	['Author', 'Author'],
	['Keywords', 'Keywords'],
	['Comments', 'Description'],
	['LastAuthor', 'LastAuthor'],
	['CreatedDate', 'Created', 'date'],
	['ModifiedDate', 'LastSaved', 'date'],
	['Category', 'Category'],
	['Manager', 'Manager'],
	['Company', 'Company'],
	['AppVersion', 'Version']
];

/* TODO: verify */
function xlml_write_docprops(Props) {
	var T = 'DocumentProperties';
	var o = [];
	XLMLDocumentProperties.forEach(function(p) {
		if(!Props[p[0]]) return;
		var m = Props[p[0]];
		switch(p[2]) {
			case 'date': m = new Date(m).toISOString(); break;
		}
		o.push(writetag(p[1], m));
	});
	return '<' + T + ' xmlns="' + XLMLNS.o + '">' + o.join("") + '</' + T + '>';
}
function xlml_write_custprops(Props, Custprops) {
	var T = 'CustomDocumentProperties';
	var o = [];
	if(Props) keys(Props).forEach(function(k) {
		/*:: if(!Props) return; */
		if(!Props.hasOwnProperty(k)) return;
		for(var i = 0; i < XLMLDocumentProperties.length; ++i)
			if(k == XLMLDocumentProperties[i][0]) return;
		var m = Props[k];
		var t = "string";
		if(typeof m == 'number') { t = "float"; m = String(m); }
		else if(m === true || m === false) { t = "boolean"; m = m ? "1" : "0"; }
		else m = String(m);
		o.push(writextag(escapexmltag(k), m, {"dt:dt":t}));
	});
	if(Custprops) keys(Custprops).forEach(function(k) {
		/*:: if(!Custprops) return; */
		if(!Custprops.hasOwnProperty(k)) return;
		var m = Custprops[k];
		var t = "string";
		if(typeof m == 'number') { t = "float"; m = String(m); }
		else if(m === true || m === false) { t = "boolean"; m = m ? "1" : "0"; }
		else if(m instanceof Date) { t = "dateTime.tz"; m = m.toISOString(); }
		else m = String(m);
		o.push(writextag(escapexmltag(k), m, {"dt:dt":t}));
	});
	return '<' + T + ' xmlns="' + XLMLNS.o + '">' + o.join("") + '</' + T + '>';
}
