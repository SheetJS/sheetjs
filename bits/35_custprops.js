/* 15.2.12.2 Custom File Properties Part */
XMLNS.CUST_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
RELS.CUST_PROPS  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';

function parse_cust_props(data, opts) {
	var p = {}, name;
	data.match(/<[^>]+>([^<]*)/g).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<?xml': break;
			case '<Properties':
				if(y.xmlns !== XMLNS.CUST_PROPS) throw "unrecognized xmlns " + y.xmlns;
				if(y.xmlnsvt && y.xmlnsvt !== XMLNS.vt) throw "unrecognized vt " + y.xmlnsvt;
				break;
			case '<property': name = y.name; break;
			case '</property>': name = null; break;
			default: if (x.indexOf('<vt:') === 0) {
				var toks = x.split('>');
				var type = toks[0].substring(4), text = toks[1];
				/* 22.4.2.32 (CT_Variant). Omit the binary types from 22.4 (Variant Types) */
				switch(type) {
					case 'lpstr': case 'lpwstr': case 'bstr': case 'lpwstr':
						p[name] = unescapexml(text);
						break;
					case 'bool':
						p[name] = parsexmlbool(text, '<vt:bool>');
						break;
					case 'i1': case 'i2': case 'i4': case 'i8': case 'int': case 'uint':
						p[name] = parseInt(text, 10);
						break;
					case 'r4': case 'r8': case 'decimal':
						p[name] = parseFloat(text);
						break;
					case 'filetime': case 'date':
						p[name] = new Date(text);
						break;
					case 'cy': case 'error':
						p[name] = unescapexml(text);
						break;
					default:
						console.warn('Unexpected', x, type, toks);
				}
			} else if(x.substr(0,2) === "</") {
			} else if(opts.WTF) throw new Error(x);
		}
	});
	return p;
}

var CUST_PROPS_XML_ROOT = writextag('Properties', null, {
	'xmlns': XMLNS.CUST_PROPS,
	'xmlns:vt': XMLNS.vt
});

function write_cust_props(cp, opts) {
	var o = [], p = {};
	o.push(XML_HEADER);
	o.push(CUST_PROPS_XML_ROOT);
	if(!cp) return o.join("");
	var pid = 1;
	keys(cp).forEach(function(k) { ++pid;
		o.push(writextag('property', write_vt(cp[k]), {
			'fmtid': '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
			'pid': pid,
			'name': k
		}));
	});
	if(o.length>2){ o.push('</Properties>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
