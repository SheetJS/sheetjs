/* 9.3 Relationships */
var RELS = ({
	WB: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
	SHEET: "http://sheetjs.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
	HLINK: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
	VML: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
	VBA: "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
}/*:any*/);

/* 9.3.3 Representing Relationships */
function get_rels_path(file/*:string*/)/*:string*/ {
	var n = file.lastIndexOf("/");
	return file.substr(0,n+1) + '_rels/' + file.substr(n+1) + ".rels";
}

function parse_rels(data/*:?string*/, currentFilePath/*:string*/) {
	if (!data) return data;
	if (currentFilePath.charAt(0) !== '/') {
		currentFilePath = '/'+currentFilePath;
	}
	var rels = {};
	var hash = {};

	(data.match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		/* 9.3.2.2 OPC_Relationships */
		if (y[0] === '<Relationship') {
			var rel = {}; rel.Type = y.Type; rel.Target = y.Target; rel.Id = y.Id; rel.TargetMode = y.TargetMode;
			var canonictarget = y.TargetMode === 'External' ? y.Target : resolve_path(y.Target, currentFilePath);
			rels[canonictarget] = rel;
			hash[y.Id] = rel;
		}
	});
	rels["!id"] = hash;
	return rels;
}

XMLNS.RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';

var RELS_ROOT = writextag('Relationships', null, {
	//'xmlns:ns0': XMLNS.RELS,
	'xmlns': XMLNS.RELS
});

/* TODO */
function write_rels(rels)/*:string*/ {
	var o = [XML_HEADER, RELS_ROOT];
	keys(rels['!id']).forEach(function(rid) {
		o[o.length] = (writextag('Relationship', null, rels['!id'][rid]));
	});
	if(o.length>2){ o[o.length] = ('</Relationships>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}

function add_rels(rels, rId, f, type, relobj)/*:number*/ {
	if(!relobj) relobj = {};
	if(!rels['!id']) rels['!id'] = {};
	if(rId < 0) for(rId = 1; rels['!id']['rId' + rId]; ++rId){/* empty */}
	relobj.Id = 'rId' + rId;
	relobj.Type = type;
	relobj.Target = f;
	if(relobj.Type == RELS.HLINK) relobj.TargetMode = "External";
	if(rels['!id'][relobj.Id]) throw new Error("Cannot rewrite rId " + rId);
	rels['!id'][relobj.Id] = relobj;
	rels[('/' + relobj.Target).replace("//","/")] = relobj;
	return rId;
}
