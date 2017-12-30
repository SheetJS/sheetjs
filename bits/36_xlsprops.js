/* Common Name -> XLML Name */
var XLMLDocPropsMap = {
	Title: 'Title',
	Subject: 'Subject',
	Author: 'Author',
	Keywords: 'Keywords',
	Comments: 'Description',
	LastAuthor: 'LastAuthor',
	RevNumber: 'Revision',
	Application: 'AppName',
	/* TotalTime: 'TotalTime', */
	LastPrinted: 'LastPrinted',
	CreatedDate: 'Created',
	ModifiedDate: 'LastSaved',
	/* Pages */
	/* Words */
	/* Characters */
	Category: 'Category',
	/* PresentationFormat */
	Manager: 'Manager',
	Company: 'Company',
	/* Guid */
	/* HyperlinkBase */
	/* Bytes */
	/* Lines */
	/* Paragraphs */
	/* CharactersWithSpaces */
	AppVersion: 'Version',

	ContentStatus: 'ContentStatus', /* NOTE: missing from schema */
	Identifier: 'Identifier', /* NOTE: missing from schema */
	Language: 'Language' /* NOTE: missing from schema */
};
var evert_XLMLDPM = evert(XLMLDocPropsMap);

function xlml_set_prop(Props, tag/*:string*/, val) {
	tag = evert_XLMLDPM[tag] || tag;
	Props[tag] = val;
}

function xlml_write_docprops(Props, opts) {
	var o/*:Array<string>*/ = [];
	keys(XLMLDocPropsMap).map(function(m) {
		for(var i = 0; i < CORE_PROPS.length; ++i) if(CORE_PROPS[i][1] == m) return CORE_PROPS[i];
		for(i = 0; i < EXT_PROPS.length; ++i) if(EXT_PROPS[i][1] == m) return EXT_PROPS[i];
		throw m;
	}).forEach(function(p) {
		if(Props[p[1]] == null) return;
		var m = opts && opts.Props && opts.Props[p[1]] != null ? opts.Props[p[1]] : Props[p[1]];
		switch(p[2]) {
			case 'date': m = new Date(m).toISOString().replace(/\.\d*Z/,"Z"); break;
		}
		if(typeof m == 'number') m = String(m);
		else if(m === true || m === false) { m = m ? "1" : "0"; }
		else if(m instanceof Date) m = new Date(m).toISOString().replace(/\.\d*Z/,"");
		o.push(writetag(XLMLDocPropsMap[p[1]] || p[1], m));
	});
	return writextag('DocumentProperties', o.join(""), {xmlns:XLMLNS.o });
}
function xlml_write_custprops(Props, Custprops, opts) {
	var BLACKLIST = ["Worksheets","SheetNames"];
	var T = 'CustomDocumentProperties';
	var o/*:Array<string>*/ = [];
	if(Props) keys(Props).forEach(function(k) {
		/*:: if(!Props) return; */
		if(!Props.hasOwnProperty(k)) return;
		for(var i = 0; i < CORE_PROPS.length; ++i) if(k == CORE_PROPS[i][1]) return;
		for(i = 0; i < EXT_PROPS.length; ++i) if(k == EXT_PROPS[i][1]) return;
		for(i = 0; i < BLACKLIST.length; ++i) if(k == BLACKLIST[i]) return;

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
