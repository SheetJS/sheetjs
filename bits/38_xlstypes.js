/* [MS-DTYP] 2.3.3 FILETIME */
/* [MS-OLEDS] 2.1.3 FILETIME (Packet Version) */
/* [MS-OLEPS] 2.8 FILETIME (Packet Version) */
function parse_FILETIME(blob) {
	var dwLowDateTime = blob.read_shift(4), dwHighDateTime = blob.read_shift(4);
	return new Date(((dwHighDateTime/1e7*Math.pow(2,32) + dwLowDateTime/1e7) - 11644473600)*1000).toISOString().replace(/\.000/,"");
}

/* [MS-OSHARED] 2.3.3.1.4 Lpstr */
function parse_lpstr(blob, type, pad) {
	var str = blob.read_shift(0, 'lpstr');
	if(pad) blob.l += (4 - ((str.length+1) & 3)) & 3;
	return str;
}

/* [MS-OSHARED] 2.3.3.1.6 Lpwstr */
function parse_lpwstr(blob, type, pad) {
	var str = blob.read_shift(0, 'lpwstr');
	if(pad) blob.l += (4 - ((str.length+1) & 3)) & 3;
	return str;
}


/* [MS-OSHARED] 2.3.3.1.11 VtString */
/* [MS-OSHARED] 2.3.3.1.12 VtUnalignedString */
function parse_VtStringBase(blob, stringType, pad) {
	if(stringType === 0x1F /*VT_LPWSTR*/) return parse_lpwstr(blob);
	return parse_lpstr(blob, stringType, pad);
}

function parse_VtString(blob, t, pad) { return parse_VtStringBase(blob, t, pad === false ? 0: 4); }
function parse_VtUnalignedString(blob, t) { if(!t) throw new Error("dafuq?"); return parse_VtStringBase(blob, t, 0); }

/* [MS-OSHARED] 2.3.3.1.9 VtVecUnalignedLpstrValue */
function parse_VtVecUnalignedLpstrValue(blob) {
	var length = blob.read_shift(4);
	var ret = [];
	for(var i = 0; i != length; ++i) ret[i] = blob.read_shift(0, 'lpstr');
	return ret;
}

/* [MS-OSHARED] 2.3.3.1.10 VtVecUnalignedLpstr */
function parse_VtVecUnalignedLpstr(blob) {
	return parse_VtVecUnalignedLpstrValue(blob);
}

/* [MS-OSHARED] 2.3.3.1.13 VtHeadingPair */
function parse_VtHeadingPair(blob) {
	var headingString = parse_TypedPropertyValue(blob, VT_USTR);
	var headerParts = parse_TypedPropertyValue(blob, VT_I4);
	return [headingString, headerParts];
}

/* [MS-OSHARED] 2.3.3.1.14 VtVecHeadingPairValue */
function parse_VtVecHeadingPairValue(blob) {
	var cElements = blob.read_shift(4);
	var out = [];
	for(var i = 0; i != cElements / 2; ++i) out.push(parse_VtHeadingPair(blob));
	return out;
}

/* [MS-OSHARED] 2.3.3.1.15 VtVecHeadingPair */
function parse_VtVecHeadingPair(blob) {
	// NOTE: When invoked, wType & padding were already consumed
	return parse_VtVecHeadingPairValue(blob);
}

/* [MS-OLEPS] 2.18.1 Dictionary (uses 2.17, 2.16) */
function parse_dictionary(blob,CodePage) {
	var cnt = blob.read_shift(4);
	var dict/*:{[number]:string}*/ = ({}/*:any*/);
	for(var j = 0; j != cnt; ++j) {
		var pid = blob.read_shift(4);
		var len = blob.read_shift(4);
		dict[pid] = blob.read_shift(len, (CodePage === 0x4B0 ?'utf16le':'utf8')).replace(chr0,'').replace(chr1,'!');
	}
	if(blob.l & 3) blob.l = (blob.l>>2+1)<<2;
	return dict;
}

/* [MS-OLEPS] 2.9 BLOB */
function parse_BLOB(blob) {
	var size = blob.read_shift(4);
	var bytes = blob.slice(blob.l,blob.l+size);
	if((size & 3) > 0) blob.l += (4 - (size & 3)) & 3;
	return bytes;
}

/* [MS-OLEPS] 2.11 ClipboardData */
function parse_ClipboardData(blob) {
	// TODO
	var o = {};
	o.Size = blob.read_shift(4);
	//o.Format = blob.read_shift(4);
	blob.l += o.Size;
	return o;
}

/* [MS-OLEPS] 2.14 Vector and Array Property Types */
function parse_VtVector(blob, cb) {
	/* [MS-OLEPS] 2.14.2 VectorHeader */
/*	var Length = blob.read_shift(4);
	var o = [];
	for(var i = 0; i != Length; ++i) {
		o.push(cb(blob));
	}
	return o;*/
}

/* [MS-OLEPS] 2.15 TypedPropertyValue */
function parse_TypedPropertyValue(blob, type, _opts) {
	var t = blob.read_shift(2), ret, opts = _opts||{};
	blob.l += 2;
	if(type !== VT_VARIANT)
	if(t !== type && VT_CUSTOM.indexOf(type)===-1) throw new Error('Expected type ' + type + ' saw ' + t);
	switch(type === VT_VARIANT ? t : type) {
		case 0x02 /*VT_I2*/: ret = blob.read_shift(2, 'i'); if(!opts.raw) blob.l += 2; return ret;
		case 0x03 /*VT_I4*/: ret = blob.read_shift(4, 'i'); return ret;
		case 0x0B /*VT_BOOL*/: return blob.read_shift(4) !== 0x0;
		case 0x13 /*VT_UI4*/: ret = blob.read_shift(4); return ret;
		case 0x1E /*VT_LPSTR*/: return parse_lpstr(blob, t, 4).replace(chr0,'');
		case 0x1F /*VT_LPWSTR*/: return parse_lpwstr(blob);
		case 0x40 /*VT_FILETIME*/: return parse_FILETIME(blob);
		case 0x41 /*VT_BLOB*/: return parse_BLOB(blob);
		case 0x47 /*VT_CF*/: return parse_ClipboardData(blob);
		case 0x50 /*VT_STRING*/: return parse_VtString(blob, t, !opts.raw && 4).replace(chr0,'');
		case 0x51 /*VT_USTR*/: return parse_VtUnalignedString(blob, t, 4).replace(chr0,'');
		case 0x100C /*VT_VECTOR|VT_VARIANT*/: return parse_VtVecHeadingPair(blob);
		case 0x101E /*VT_LPSTR*/: return parse_VtVecUnalignedLpstr(blob);
		default: throw new Error("TypedPropertyValue unrecognized type " + type + " " + t);
	}
}
/* [MS-OLEPS] 2.14.2 VectorHeader */
/*function parse_VTVectorVariant(blob) {
	var Length = blob.read_shift(4);

	if(Length & 1 !== 0) throw new Error("VectorHeader Length=" + Length + " must be even");
	var o = [];
	for(var i = 0; i != Length; ++i) {
		o.push(parse_TypedPropertyValue(blob, VT_VARIANT));
	}
	return o;
}*/

/* [MS-OLEPS] 2.20 PropertySet */
function parse_PropertySet(blob, PIDSI) {
	var start_addr = blob.l;
	var size = blob.read_shift(4);
	var NumProps = blob.read_shift(4);
	var Props = [], i = 0;
	var CodePage = 0;
	var Dictionary = -1, DictObj/*:{[number]:string}*/ = ({}/*:any*/);
	for(i = 0; i != NumProps; ++i) {
		var PropID = blob.read_shift(4);
		var Offset = blob.read_shift(4);
		Props[i] = [PropID, Offset + start_addr];
	}
	var PropH = {};
	for(i = 0; i != NumProps; ++i) {
		if(blob.l !== Props[i][1]) {
			var fail = true;
			if(i>0 && PIDSI) switch(PIDSI[Props[i-1][0]].t) {
				case 0x02 /*VT_I2*/: if(blob.l +2 === Props[i][1]) { blob.l+=2; fail = false; } break;
				case 0x50 /*VT_STRING*/: if(blob.l <= Props[i][1]) { blob.l=Props[i][1]; fail = false; } break;
				case 0x100C /*VT_VECTOR|VT_VARIANT*/: if(blob.l <= Props[i][1]) { blob.l=Props[i][1]; fail = false; } break;
			}
			if(!PIDSI && blob.l <= Props[i][1]) { fail=false; blob.l = Props[i][1]; }
			if(fail) throw new Error("Read Error: Expected address " + Props[i][1] + ' at ' + blob.l + ' :' + i);
		}
		if(PIDSI) {
			var piddsi = PIDSI[Props[i][0]];
			PropH[piddsi.n] = parse_TypedPropertyValue(blob, piddsi.t, {raw:true});
			if(piddsi.p === 'version') PropH[piddsi.n] = String(PropH[piddsi.n] >> 16) + "." + String(PropH[piddsi.n] & 0xFFFF);
			if(piddsi.n == "CodePage") switch(PropH[piddsi.n]) {
				case 0: PropH[piddsi.n] = 1252;
					/* falls through */
				case 10000: // OSX Roman
				case 1252: // Windows Latin

				case 874: // SB Windows Thai
				case 1250: // SB Windows Central Europe
				case 1251: // SB Windows Cyrillic
				case 1253: // SB Windows Greek
				case 1254: // SB Windows Turkish
				case 1255: // SB Windows Hebrew
				case 1256: // SB Windows Arabic
				case 1257: // SB Windows Baltic
				case 1258: // SB Windows Vietnam

				case 932: // DB Windows Japanese Shift-JIS
				case 936: // DB Windows Simplified Chinese GBK
				case 949: // DB Windows Korean
				case 950: // DB Windows Traditional Chinese Big5

				case 1200: // UTF16LE
				case 1201: // UTF16BE
				case 65000: case -536: // UTF-7
				case 65001: case -535: // UTF-8
					set_cp(CodePage = PropH[piddsi.n]); break;
				default: throw new Error("Unsupported CodePage: " + PropH[piddsi.n]);
			}
		} else {
			if(Props[i][0] === 0x1) {
				CodePage = PropH.CodePage = parse_TypedPropertyValue(blob, VT_I2);
				set_cp(CodePage);
				if(Dictionary !== -1) {
					var oldpos = blob.l;
					blob.l = Props[Dictionary][1];
					DictObj = parse_dictionary(blob,CodePage);
					blob.l = oldpos;
				}
			} else if(Props[i][0] === 0) {
				if(CodePage === 0) { Dictionary = i; blob.l = Props[i+1][1]; continue; }
				DictObj = parse_dictionary(blob,CodePage);
			} else {
				var name = DictObj[Props[i][0]];
				var val;
				/* [MS-OSHARED] 2.3.3.2.3.1.2 + PROPVARIANT */
				switch(blob[blob.l]) {
					case 0x41 /*VT_BLOB*/: blob.l += 4; val = parse_BLOB(blob); break;
					case 0x1E /*VT_LPSTR*/: blob.l += 4; val = parse_VtString(blob, blob[blob.l-4]); break;
					case 0x1F /*VT_LPWSTR*/: blob.l += 4; val = parse_VtString(blob, blob[blob.l-4]); break;
					case 0x03 /*VT_I4*/: blob.l += 4; val = blob.read_shift(4, 'i'); break;
					case 0x13 /*VT_UI4*/: blob.l += 4; val = blob.read_shift(4); break;
					case 0x05 /*VT_R8*/: blob.l += 4; val = blob.read_shift(8, 'f'); break;
					case 0x0B /*VT_BOOL*/: blob.l += 4; val = parsebool(blob, 4); break;
					case 0x40 /*VT_FILETIME*/: blob.l += 4; val = new Date(parse_FILETIME(blob)); break;
					default: throw new Error("unparsed value: " + blob[blob.l]);
				}
				PropH[name] = val;
			}
		}
	}
	blob.l = start_addr + size; /* step ahead to skip padding */
	return PropH;
}

/* [MS-OLEPS] 2.21 PropertySetStream */
function parse_PropertySetStream(file, PIDSI) {
	var blob = file.content;
	prep_blob(blob, 0);

	var NumSets, FMTID0, FMTID1, Offset0, Offset1 = 0;
	blob.chk('feff', 'Byte Order: ');

	var vers = blob.read_shift(2); // TODO: check version
	var SystemIdentifier = blob.read_shift(4);
	blob.chk(CFB.utils.consts.HEADER_CLSID, 'CLSID: ');
	NumSets = blob.read_shift(4);
	if(NumSets !== 1 && NumSets !== 2) throw "Unrecognized #Sets: " + NumSets;
	FMTID0 = blob.read_shift(16); Offset0 = blob.read_shift(4);

	if(NumSets === 1 && Offset0 !== blob.l) throw "Length mismatch";
	else if(NumSets === 2) { FMTID1 = blob.read_shift(16); Offset1 = blob.read_shift(4); }
	var PSet0 = parse_PropertySet(blob, PIDSI);

	var rval = ({ SystemIdentifier: SystemIdentifier }/*:any*/);
	for(var y in PSet0) rval[y] = PSet0[y];
	//rval.blob = blob;
	rval.FMTID = FMTID0;
	//rval.PSet0 = PSet0;
	if(NumSets === 1) return rval;
	if(blob.l !== Offset1) throw new Error("Length mismatch 2: " + blob.l + " !== " + Offset1);
	var PSet1;
	try { PSet1 = parse_PropertySet(blob, null); } catch(e) { }
	for(y in PSet1) rval[y] = PSet1[y];
	rval.FMTID = [FMTID0, FMTID1]; // TODO: verify FMTID0/1
	return rval;
}


function parsenoop2(blob, length) { blob.read_shift(length); return null; }

function parslurp(blob, length, cb) {
	var arr = [], target = blob.l + length;
	while(blob.l < target) arr.push(cb(blob, target - blob.l));
	if(target !== blob.l) throw new Error("Slurp error");
	return arr;
}

function parslurp2(blob, length, cb) {
	var arr = [], target = blob.l + length, len = blob.read_shift(2);
	while(len-- !== 0) arr.push(cb(blob, target - blob.l));
	if(target !== blob.l) throw new Error("Slurp error");
	return arr;
}

function parsebool(blob, length) { return blob.read_shift(length) === 0x1; }

function parseuint16(blob) { return blob.read_shift(2, 'u'); }
function parseuint16a(blob, length) { return parslurp(blob,length,parseuint16);}

/* --- 2.5 Structures --- */

/* [MS-XLS] 2.5.14 Boolean */
var parse_Boolean = parsebool;

/* [MS-XLS] 2.5.10 Bes (boolean or error) */
function parse_Bes(blob) {
	var v = blob.read_shift(1), t = blob.read_shift(1);
	return t === 0x01 ? v : v === 0x01;
}

/* [MS-XLS] 2.5.240 ShortXLUnicodeString */
function parse_ShortXLUnicodeString(blob, length, opts) {
	var cch = blob.read_shift(opts && opts.biff >= 12 ? 2 : 1);
	var width = 1, encoding = 'sbcs-cont';
	var cp = current_codepage;
	if(opts && opts.biff >= 8) current_codepage = 1200;
	if(!opts || opts.biff == 8 ) {
		var fHighByte = blob.read_shift(1);
		if(fHighByte) { width = 2; encoding = 'dbcs-cont'; }
	} else if(opts.biff == 12) {
		width = 2; encoding = 'wstr';
	}
	var o = cch ? blob.read_shift(cch, encoding) : "";
	current_codepage = cp;
	return o;
}

/* 2.5.293 XLUnicodeRichExtendedString */
function parse_XLUnicodeRichExtendedString(blob) {
	var cp = current_codepage;
	current_codepage = 1200;
	var cch = blob.read_shift(2), flags = blob.read_shift(1);
	var fHighByte = flags & 0x1, fExtSt = flags & 0x4, fRichSt = flags & 0x8;
	var width = 1 + (flags & 0x1); // 0x0 -> utf8, 0x1 -> dbcs
	var cRun = 0, cbExtRst;
	var z = {};
	if(fRichSt) cRun = blob.read_shift(2);
	if(fExtSt) cbExtRst = blob.read_shift(4);
	var encoding = (flags & 0x1) ? 'dbcs-cont' : 'sbcs-cont';
	var msg = cch === 0 ? "" : blob.read_shift(cch, encoding);
	if(fRichSt) blob.l += 4 * cRun; //TODO: parse this
	if(fExtSt) blob.l += cbExtRst; //TODO: parse this
	z.t = msg;
	if(!fRichSt) { z.raw = "<t>" + z.t + "</t>"; z.r = z.t; }
	current_codepage = cp;
	return z;
}

/* 2.5.296 XLUnicodeStringNoCch */
function parse_XLUnicodeStringNoCch(blob, cch, opts) {
	var retval;
	if(opts) {
		if(opts.biff >= 2 && opts.biff <= 5) return blob.read_shift(cch, 'sbcs-cont');
		if(opts.biff >= 12) return blob.read_shift(cch, 'dbcs-cont');
	}
	var fHighByte = blob.read_shift(1);
	if(fHighByte===0) { retval = blob.read_shift(cch, 'sbcs-cont'); }
	else { retval = blob.read_shift(cch, 'dbcs-cont'); }
	return retval;
}

/* 2.5.294 XLUnicodeString */
function parse_XLUnicodeString(blob, length, opts) {
	var cch = blob.read_shift(opts && opts.biff == 2 ? 1 : 2);
	if(cch === 0) { blob.l++; return ""; }
	return parse_XLUnicodeStringNoCch(blob, cch, opts);
}
/* BIFF5 override */
function parse_XLUnicodeString2(blob, length, opts) {
	if(opts.biff > 5) return parse_XLUnicodeString(blob, length, opts);
	var cch = blob.read_shift(1);
	if(cch === 0) { blob.l++; return ""; }
	return blob.read_shift(cch, 'sbcs-cont');
}

/* [MS-XLS] 2.5.61 ControlInfo */
var parse_ControlInfo = parsenoop;

/* [MS-OSHARED] 2.3.7.6 URLMoniker TODO: flags */
var parse_URLMoniker = function(blob, length) {
	var len = blob.read_shift(4), start = blob.l;
	var extra = false;
	if(len > 24) {
		/* look ahead */
		blob.l += len - 24;
		if(blob.read_shift(16) === "795881f43b1d7f48af2c825dc4852763") extra = true;
		blob.l = start;
	}
	var url = blob.read_shift((extra?len-24:len)>>1, 'utf16le').replace(chr0,"");
	if(extra) blob.l += 24;
	return url;
};

/* [MS-OSHARED] 2.3.7.8 FileMoniker TODO: all fields */
var parse_FileMoniker = function(blob, length) {
	var cAnti = blob.read_shift(2);
	var ansiLength = blob.read_shift(4);
	var ansiPath = blob.read_shift(ansiLength, 'cstr');
	var endServer = blob.read_shift(2);
	var versionNumber = blob.read_shift(2);
	var cbUnicodePathSize = blob.read_shift(4);
	if(cbUnicodePathSize === 0) return ansiPath.replace(/\\/g,"/");
	var cbUnicodePathBytes = blob.read_shift(4);
	var usKeyValue = blob.read_shift(2);
	var unicodePath = blob.read_shift(cbUnicodePathBytes>>1, 'utf16le').replace(chr0,"");
	return unicodePath;
};

/* [MS-OSHARED] 2.3.7.2 HyperlinkMoniker TODO: all the monikers */
var parse_HyperlinkMoniker = function(blob, length) {
	var clsid = blob.read_shift(16); length -= 16;
	switch(clsid) {
		case "e0c9ea79f9bace118c8200aa004ba90b": return parse_URLMoniker(blob, length);
		case "0303000000000000c000000000000046": return parse_FileMoniker(blob, length);
		default: throw "unsupported moniker " + clsid;
	}
};

/* [MS-OSHARED] 2.3.7.9 HyperlinkString */
var parse_HyperlinkString = function(blob, length) {
	var len = blob.read_shift(4);
	var o = blob.read_shift(len, 'utf16le').replace(chr0, "");
	return o;
};

/* [MS-OSHARED] 2.3.7.1 Hyperlink Object TODO: unify params with XLSX */
var parse_Hyperlink = function(blob, length) {
	var end = blob.l + length;
	var sVer = blob.read_shift(4);
	if(sVer !== 2) throw new Error("Unrecognized streamVersion: " + sVer);
	var flags = blob.read_shift(2);
	blob.l += 2;
	var displayName, targetFrameName, moniker, oleMoniker, location, guid, fileTime;
	if(flags & 0x0010) displayName = parse_HyperlinkString(blob, end - blob.l);
	if(flags & 0x0080) targetFrameName = parse_HyperlinkString(blob, end - blob.l);
	if((flags & 0x0101) === 0x0101) moniker = parse_HyperlinkString(blob, end - blob.l);
	if((flags & 0x0101) === 0x0001) oleMoniker = parse_HyperlinkMoniker(blob, end - blob.l);
	if(flags & 0x0008) location = parse_HyperlinkString(blob, end - blob.l);
	if(flags & 0x0020) guid = blob.read_shift(16);
	if(flags & 0x0040) fileTime = parse_FILETIME(blob, 8);
	blob.l = end;
	var target = (targetFrameName||moniker||oleMoniker);
	if(location) target+="#"+location;
	return {Target: target};
};

/* 2.5.178 LongRGBA */
function parse_LongRGBA(blob, length) { var r = blob.read_shift(1), g = blob.read_shift(1), b = blob.read_shift(1), a = blob.read_shift(1); return [r,g,b,a]; }

/* 2.5.177 LongRGB */
function parse_LongRGB(blob, length) { var x = parse_LongRGBA(blob, length); x[3] = 0; return x; }


