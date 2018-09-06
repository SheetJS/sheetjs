var DO_NOT_EXPORT_CFB = true;
/*::
declare var Base64:any;
declare var ReadShift:any;
declare var CheckField:any;
declare var prep_blob:any;
declare var __readUInt32LE:any;
declare var __readInt32LE:any;
declare var __toBuffer:any;
declare var __utf16le:any;
declare var bconcat:any;
declare var s2a:any;
declare var chr0:any;
declare var chr1:any;
declare var has_buf:boolean;
declare var new_buf:any;
declare var new_raw_buf:any;
declare var new_unsafe_buf:any;
*/
/* cfb.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*jshint eqnull:true */
/*exported CFB */
/*global module, require:false, process:false, Buffer:false, Uint8Array:false, Uint16Array:false */

/*::
declare var DO_NOT_EXPORT_CFB:?boolean;
type SectorEntry = {
	name?:string;
	nodes?:Array<number>;
	data:RawBytes;
};
type SectorList = {
	[k:string|number]:SectorEntry;
	name:?string;
	fat_addrs:Array<number>;
	ssz:number;
}
type CFBFiles = {[n:string]:CFBEntry};
*/
/* crc32.js (C) 2014-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*exported CRC32 */
var CRC32;
(function (factory) {
	/*jshint ignore:start */
	/*eslint-disable */
	factory(CRC32 = {});
	/*eslint-enable */
	/*jshint ignore:end */
}(function(CRC32) {
CRC32.version = '1.2.0';
/* see perf/crc32table.js */
/*global Int32Array */
function signed_crc_table()/*:any*/ {
	var c = 0, table/*:Array<number>*/ = new Array(256);

	for(var n =0; n != 256; ++n){
		c = n;
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		table[n] = c;
	}

	return typeof Int32Array !== 'undefined' ? new Int32Array(table) : table;
}

var T = signed_crc_table();
function crc32_bstr(bstr/*:string*/, seed/*:number*/)/*:number*/ {
	var C = seed ^ -1, L = bstr.length - 1;
	for(var i = 0; i < L;) {
		C = (C>>>8) ^ T[(C^bstr.charCodeAt(i++))&0xFF];
		C = (C>>>8) ^ T[(C^bstr.charCodeAt(i++))&0xFF];
	}
	if(i === L) C = (C>>>8) ^ T[(C ^ bstr.charCodeAt(i))&0xFF];
	return C ^ -1;
}

function crc32_buf(buf/*:Uint8Array|Array<number>*/, seed/*:number*/)/*:number*/ {
	if(buf.length > 10000) return crc32_buf_8(buf, seed);
	var C = seed ^ -1, L = buf.length - 3;
	for(var i = 0; i < L;) {
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
	}
	while(i < L+3) C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
	return C ^ -1;
}

function crc32_buf_8(buf/*:Uint8Array|Array<number>*/, seed/*:number*/)/*:number*/ {
	var C = seed ^ -1, L = buf.length - 7;
	for(var i = 0; i < L;) {
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
	}
	while(i < L+7) C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
	return C ^ -1;
}

function crc32_str(str/*:string*/, seed/*:number*/)/*:number*/ {
	var C = seed ^ -1;
	for(var i = 0, L=str.length, c, d; i < L;) {
		c = str.charCodeAt(i++);
		if(c < 0x80) {
			C = (C>>>8) ^ T[(C ^ c)&0xFF];
		} else if(c < 0x800) {
			C = (C>>>8) ^ T[(C ^ (192|((c>>6)&31)))&0xFF];
			C = (C>>>8) ^ T[(C ^ (128|(c&63)))&0xFF];
		} else if(c >= 0xD800 && c < 0xE000) {
			c = (c&1023)+64; d = str.charCodeAt(i++)&1023;
			C = (C>>>8) ^ T[(C ^ (240|((c>>8)&7)))&0xFF];
			C = (C>>>8) ^ T[(C ^ (128|((c>>2)&63)))&0xFF];
			C = (C>>>8) ^ T[(C ^ (128|((d>>6)&15)|((c&3)<<4)))&0xFF];
			C = (C>>>8) ^ T[(C ^ (128|(d&63)))&0xFF];
		} else {
			C = (C>>>8) ^ T[(C ^ (224|((c>>12)&15)))&0xFF];
			C = (C>>>8) ^ T[(C ^ (128|((c>>6)&63)))&0xFF];
			C = (C>>>8) ^ T[(C ^ (128|(c&63)))&0xFF];
		}
	}
	return C ^ -1;
}
CRC32.table = T;
CRC32.bstr = crc32_bstr;
CRC32.buf = crc32_buf;
CRC32.str = crc32_str;
}));
/* [MS-CFB] v20171201 */
var CFB = (function _CFB(){
var exports/*:CFBModule*/ = /*::(*/{}/*:: :any)*/;
exports.version = '1.1.0';
/* [MS-CFB] 2.6.4 */
function namecmp(l/*:string*/, r/*:string*/)/*:number*/ {
	var L = l.split("/"), R = r.split("/");
	for(var i = 0, c = 0, Z = Math.min(L.length, R.length); i < Z; ++i) {
		if((c = L[i].length - R[i].length)) return c;
		if(L[i] != R[i]) return L[i] < R[i] ? -1 : 1;
	}
	return L.length - R.length;
}
function dirname(p/*:string*/)/*:string*/ {
	if(p.charAt(p.length - 1) == "/") return (p.slice(0,-1).indexOf("/") === -1) ? p : dirname(p.slice(0, -1));
	var c = p.lastIndexOf("/");
	return (c === -1) ? p : p.slice(0, c+1);
}

function filename(p/*:string*/)/*:string*/ {
	if(p.charAt(p.length - 1) == "/") return filename(p.slice(0, -1));
	var c = p.lastIndexOf("/");
	return (c === -1) ? p : p.slice(c+1);
}
/* -------------------------------------------------------------------------- */
/* DOS Date format:
   high|YYYYYYYm.mmmddddd.HHHHHMMM.MMMSSSSS|low
   add 1980 to stored year
   stored second should be doubled
*/

/* write JS date to buf as a DOS date */
function write_dos_date(buf/*:CFBlob*/, date/*:Date|string*/) {
	if(typeof date === "string") date = new Date(date);
	var hms/*:number*/ = date.getHours();
	hms = hms << 6 | date.getMinutes();
	hms = hms << 5 | (date.getSeconds()>>>1);
	buf.write_shift(2, hms);
	var ymd/*:number*/ = (date.getFullYear() - 1980);
	ymd = ymd << 4 | (date.getMonth()+1);
	ymd = ymd << 5 | date.getDate();
	buf.write_shift(2, ymd);
}

/* read four bytes from buf and interpret as a DOS date */
function parse_dos_date(buf/*:CFBlob*/)/*:Date*/ {
	var hms = buf.read_shift(2) & 0xFFFF;
	var ymd = buf.read_shift(2) & 0xFFFF;
	var val = new Date();
	var d = ymd & 0x1F; ymd >>>= 5;
	var m = ymd & 0x0F; ymd >>>= 4;
	val.setMilliseconds(0);
	val.setFullYear(ymd + 1980);
	val.setMonth(m-1);
	val.setDate(d);
	var S = hms & 0x1F; hms >>>= 5;
	var M = hms & 0x3F; hms >>>= 6;
	val.setHours(hms);
	val.setMinutes(M);
	val.setSeconds(S<<1);
	return val;
}
function parse_extra_field(blob/*:CFBlob*/)/*:any*/ {
	prep_blob(blob, 0);
	var o = /*::(*/{}/*:: :any)*/;
	var flags = 0;
	while(blob.l <= blob.length - 4) {
		var type = blob.read_shift(2);
		var sz = blob.read_shift(2), tgt = blob.l + sz;
		var p = {};
		switch(type) {
			/* UNIX-style Timestamps */
			case 0x5455: {
				flags = blob.read_shift(1);
				if(flags & 1) p.mtime = blob.read_shift(4);
				/* for some reason, CD flag corresponds to LFH */
				if(sz > 5) {
					if(flags & 2) p.atime = blob.read_shift(4);
					if(flags & 4) p.ctime = blob.read_shift(4);
				}
				if(p.mtime) p.mt = new Date(p.mtime*1000);
			}
			break;
		}
		blob.l = tgt;
		o[type] = p;
	}
	return o;
}
var fs/*:: = require('fs'); */;
function get_fs() { return fs || (fs = require('fs')); }
function parse(file/*:RawBytes*/, options/*:CFBReadOpts*/)/*:CFBContainer*/ {
if(file[0] == 0x50 && file[1] == 0x4b) return parse_zip(file, options);
if(file.length < 512) throw new Error("CFB file size " + file.length + " < 512");
var mver = 3;
var ssz = 512;
var nmfs = 0; // number of mini FAT sectors
var difat_sec_cnt = 0;
var dir_start = 0;
var minifat_start = 0;
var difat_start = 0;

var fat_addrs/*:Array<number>*/ = []; // locations of FAT sectors

/* [MS-CFB] 2.2 Compound File Header */
var blob/*:CFBlob*/ = /*::(*/file.slice(0,512)/*:: :any)*/;
prep_blob(blob, 0);

/* major version */
var mv = check_get_mver(blob);
mver = mv[0];
switch(mver) {
	case 3: ssz = 512; break; case 4: ssz = 4096; break;
	case 0: if(mv[1] == 0) return parse_zip(file, options);
	/* falls through */
	default: throw new Error("Major Version: Expected 3 or 4 saw " + mver);
}

/* reprocess header */
if(ssz !== 512) { blob = /*::(*/file.slice(0,ssz)/*:: :any)*/; prep_blob(blob, 28 /* blob.l */); }
/* Save header for final object */
var header/*:RawBytes*/ = file.slice(0,ssz);

check_shifts(blob, mver);

// Number of Directory Sectors
var dir_cnt/*:number*/ = blob.read_shift(4, 'i');
if(mver === 3 && dir_cnt !== 0) throw new Error('# Directory Sectors: Expected 0 saw ' + dir_cnt);

// Number of FAT Sectors
blob.l += 4;

// First Directory Sector Location
dir_start = blob.read_shift(4, 'i');

// Transaction Signature
blob.l += 4;

// Mini Stream Cutoff Size
blob.chk('00100000', 'Mini Stream Cutoff Size: ');

// First Mini FAT Sector Location
minifat_start = blob.read_shift(4, 'i');

// Number of Mini FAT Sectors
nmfs = blob.read_shift(4, 'i');

// First DIFAT sector location
difat_start = blob.read_shift(4, 'i');

// Number of DIFAT Sectors
difat_sec_cnt = blob.read_shift(4, 'i');

// Grab FAT Sector Locations
for(var q = -1, j = 0; j < 109; ++j) { /* 109 = (512 - blob.l)>>>2; */
	q = blob.read_shift(4, 'i');
	if(q<0) break;
	fat_addrs[j] = q;
}

/** Break the file up into sectors */
var sectors/*:Array<RawBytes>*/ = sectorify(file, ssz);

sleuth_fat(difat_start, difat_sec_cnt, sectors, ssz, fat_addrs);

/** Chains */
var sector_list/*:SectorList*/ = make_sector_list(sectors, dir_start, fat_addrs, ssz);

sector_list[dir_start].name = "!Directory";
if(nmfs > 0 && minifat_start !== ENDOFCHAIN) sector_list[minifat_start].name = "!MiniFAT";
sector_list[fat_addrs[0]].name = "!FAT";
sector_list.fat_addrs = fat_addrs;
sector_list.ssz = ssz;

/* [MS-CFB] 2.6.1 Compound File Directory Entry */
var files/*:CFBFiles*/ = {}, Paths/*:Array<string>*/ = [], FileIndex/*:CFBFileIndex*/ = [], FullPaths/*:Array<string>*/ = [];
read_directory(dir_start, sector_list, sectors, Paths, nmfs, files, FileIndex, minifat_start);

build_full_paths(FileIndex, FullPaths, Paths);
Paths.shift();

var o = {
	FileIndex: FileIndex,
	FullPaths: FullPaths
};

// $FlowIgnore
if(options && options.raw) o.raw = {header: header, sectors: sectors};
return o;
} // parse

/* [MS-CFB] 2.2 Compound File Header -- read up to major version */
function check_get_mver(blob/*:CFBlob*/)/*:[number, number]*/ {
	if(blob[blob.l] == 0x50 && blob[blob.l + 1] == 0x4b) return [0, 0];
	// header signature 8
	blob.chk(HEADER_SIGNATURE, 'Header Signature: ');

	// clsid 16
	blob.chk(HEADER_CLSID, 'CLSID: ');

	// minor version 2
	var mver/*:number*/ = blob.read_shift(2, 'u');

	return [blob.read_shift(2,'u'), mver];
}
function check_shifts(blob/*:CFBlob*/, mver/*:number*/)/*:void*/ {
	var shift = 0x09;

	// Byte Order
	//blob.chk('feff', 'Byte Order: '); // note: some writers put 0xffff
	blob.l += 2;

	// Sector Shift
	switch((shift = blob.read_shift(2))) {
		case 0x09: if(mver != 3) throw new Error('Sector Shift: Expected 9 saw ' + shift); break;
		case 0x0c: if(mver != 4) throw new Error('Sector Shift: Expected 12 saw ' + shift); break;
		default: throw new Error('Sector Shift: Expected 9 or 12 saw ' + shift);
	}

	// Mini Sector Shift
	blob.chk('0600', 'Mini Sector Shift: ');

	// Reserved
	blob.chk('000000000000', 'Reserved: ');
}

/** Break the file up into sectors */
function sectorify(file/*:RawBytes*/, ssz/*:number*/)/*:Array<RawBytes>*/ {
	var nsectors = Math.ceil(file.length/ssz)-1;
	var sectors/*:Array<RawBytes>*/ = [];
	for(var i=1; i < nsectors; ++i) sectors[i-1] = file.slice(i*ssz,(i+1)*ssz);
	sectors[nsectors-1] = file.slice(nsectors*ssz);
	return sectors;
}

/* [MS-CFB] 2.6.4 Red-Black Tree */
function build_full_paths(FI/*:CFBFileIndex*/, FP/*:Array<string>*/, Paths/*:Array<string>*/)/*:void*/ {
	var i = 0, L = 0, R = 0, C = 0, j = 0, pl = Paths.length;
	var dad/*:Array<number>*/ = [], q/*:Array<number>*/ = [];

	for(; i < pl; ++i) { dad[i]=q[i]=i; FP[i]=Paths[i]; }

	for(; j < q.length; ++j) {
		i = q[j];
		L = FI[i].L; R = FI[i].R; C = FI[i].C;
		if(dad[i] === i) {
			if(L !== -1 /*NOSTREAM*/ && dad[L] !== L) dad[i] = dad[L];
			if(R !== -1 && dad[R] !== R) dad[i] = dad[R];
		}
		if(C !== -1 /*NOSTREAM*/) dad[C] = i;
		if(L !== -1) { dad[L] = dad[i]; if(q.lastIndexOf(L) < j) q.push(L); }
		if(R !== -1) { dad[R] = dad[i]; if(q.lastIndexOf(R) < j) q.push(R); }
	}
	for(i=1; i < pl; ++i) if(dad[i] === i) {
		if(R !== -1 /*NOSTREAM*/ && dad[R] !== R) dad[i] = dad[R];
		else if(L !== -1 && dad[L] !== L) dad[i] = dad[L];
	}

	for(i=1; i < pl; ++i) {
		if(FI[i].type === 0 /* unknown */) continue;
		j = dad[i];
		if(j === 0) FP[i] = FP[0] + "/" + FP[i];
		else while(j !== 0 && j !== dad[j]) {
			FP[i] = FP[j] + "/" + FP[i];
			j = dad[j];
		}
		dad[i] = 0;
	}

	FP[0] += "/";
	for(i=1; i < pl; ++i) {
		if(FI[i].type !== 2 /* stream */) FP[i] += "/";
	}
}

function get_mfat_entry(entry/*:CFBEntry*/, payload/*:RawBytes*/, mini/*:?RawBytes*/)/*:CFBlob*/ {
	var start = entry.start, size = entry.size;
	//return (payload.slice(start*MSSZ, start*MSSZ + size)/*:any*/);
	var o = [];
	var idx = start;
	while(mini && size > 0 && idx >= 0) {
		o.push(payload.slice(idx * MSSZ, idx * MSSZ + MSSZ));
		size -= MSSZ;
		idx = __readInt32LE(mini, idx * 4);
	}
	if(o.length === 0) return (new_buf(0)/*:any*/);
	return (bconcat(o).slice(0, entry.size)/*:any*/);
}

/** Chase down the rest of the DIFAT chain to build a comprehensive list
    DIFAT chains by storing the next sector number as the last 32 bits */
function sleuth_fat(idx/*:number*/, cnt/*:number*/, sectors/*:Array<RawBytes>*/, ssz/*:number*/, fat_addrs)/*:void*/ {
	var q/*:number*/ = ENDOFCHAIN;
	if(idx === ENDOFCHAIN) {
		if(cnt !== 0) throw new Error("DIFAT chain shorter than expected");
	} else if(idx !== -1 /*FREESECT*/) {
		var sector = sectors[idx], m = (ssz>>>2)-1;
		if(!sector) return;
		for(var i = 0; i < m; ++i) {
			if((q = __readInt32LE(sector,i*4)) === ENDOFCHAIN) break;
			fat_addrs.push(q);
		}
		sleuth_fat(__readInt32LE(sector,ssz-4),cnt - 1, sectors, ssz, fat_addrs);
	}
}

/** Follow the linked list of sectors for a given starting point */
function get_sector_list(sectors/*:Array<RawBytes>*/, start/*:number*/, fat_addrs/*:Array<number>*/, ssz/*:number*/, chkd/*:?Array<boolean>*/)/*:SectorEntry*/ {
	var buf/*:Array<number>*/ = [], buf_chain/*:Array<any>*/ = [];
	if(!chkd) chkd = [];
	var modulus = ssz - 1, j = 0, jj = 0;
	for(j=start; j>=0;) {
		chkd[j] = true;
		buf[buf.length] = j;
		buf_chain.push(sectors[j]);
		var addr = fat_addrs[Math.floor(j*4/ssz)];
		jj = ((j*4) & modulus);
		if(ssz < 4 + jj) throw new Error("FAT boundary crossed: " + j + " 4 "+ssz);
		if(!sectors[addr]) break;
		j = __readInt32LE(sectors[addr], jj);
	}
	return {nodes: buf, data:__toBuffer([buf_chain])};
}

/** Chase down the sector linked lists */
function make_sector_list(sectors/*:Array<RawBytes>*/, dir_start/*:number*/, fat_addrs/*:Array<number>*/, ssz/*:number*/)/*:SectorList*/ {
	var sl = sectors.length, sector_list/*:SectorList*/ = ([]/*:any*/);
	var chkd/*:Array<boolean>*/ = [], buf/*:Array<number>*/ = [], buf_chain/*:Array<RawBytes>*/ = [];
	var modulus = ssz - 1, i=0, j=0, k=0, jj=0;
	for(i=0; i < sl; ++i) {
		buf = ([]/*:Array<number>*/);
		k = (i + dir_start); if(k >= sl) k-=sl;
		if(chkd[k]) continue;
		buf_chain = [];
		for(j=k; j>=0;) {
			chkd[j] = true;
			buf[buf.length] = j;
			buf_chain.push(sectors[j]);
			var addr/*:number*/ = fat_addrs[Math.floor(j*4/ssz)];
			jj = ((j*4) & modulus);
			if(ssz < 4 + jj) throw new Error("FAT boundary crossed: " + j + " 4 "+ssz);
			if(!sectors[addr]) break;
			j = __readInt32LE(sectors[addr], jj);
		}
		sector_list[k] = ({nodes: buf, data:__toBuffer([buf_chain])}/*:SectorEntry*/);
	}
	return sector_list;
}

/* [MS-CFB] 2.6.1 Compound File Directory Entry */
function read_directory(dir_start/*:number*/, sector_list/*:SectorList*/, sectors/*:Array<RawBytes>*/, Paths/*:Array<string>*/, nmfs, files, FileIndex, mini) {
	var minifat_store = 0, pl = (Paths.length?2:0);
	var sector = sector_list[dir_start].data;
	var i = 0, namelen = 0, name;
	for(; i < sector.length; i+= 128) {
		var blob/*:CFBlob*/ = /*::(*/sector.slice(i, i+128)/*:: :any)*/;
		prep_blob(blob, 64);
		namelen = blob.read_shift(2);
		name = __utf16le(blob,0,namelen-pl);
		Paths.push(name);
		var o/*:CFBEntry*/ = ({
			name:  name,
			type:  blob.read_shift(1),
			color: blob.read_shift(1),
			L:     blob.read_shift(4, 'i'),
			R:     blob.read_shift(4, 'i'),
			C:     blob.read_shift(4, 'i'),
			clsid: blob.read_shift(16),
			state: blob.read_shift(4, 'i'),
			start: 0,
			size: 0
		});
		var ctime/*:number*/ = blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2);
		if(ctime !== 0) o.ct = read_date(blob, blob.l-8);
		var mtime/*:number*/ = blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2);
		if(mtime !== 0) o.mt = read_date(blob, blob.l-8);
		o.start = blob.read_shift(4, 'i');
		o.size = blob.read_shift(4, 'i');
		if(o.size < 0 && o.start < 0) { o.size = o.type = 0; o.start = ENDOFCHAIN; o.name = ""; }
		if(o.type === 5) { /* root */
			minifat_store = o.start;
			if(nmfs > 0 && minifat_store !== ENDOFCHAIN) sector_list[minifat_store].name = "!StreamData";
			/*minifat_size = o.size;*/
		} else if(o.size >= 4096 /* MSCSZ */) {
			o.storage = 'fat';
			if(sector_list[o.start] === undefined) sector_list[o.start] = get_sector_list(sectors, o.start, sector_list.fat_addrs, sector_list.ssz);
			sector_list[o.start].name = o.name;
			o.content = (sector_list[o.start].data.slice(0,o.size)/*:any*/);
		} else {
			o.storage = 'minifat';
			if(o.size < 0) o.size = 0;
			else if(minifat_store !== ENDOFCHAIN && o.start !== ENDOFCHAIN && sector_list[minifat_store]) {
				o.content = get_mfat_entry(o, sector_list[minifat_store].data, (sector_list[mini]||{}).data);
			}
		}
		if(o.content) prep_blob(o.content, 0);
		files[name] = o;
		FileIndex.push(o);
	}
}

function read_date(blob/*:RawBytes|CFBlob*/, offset/*:number*/)/*:Date*/ {
	return new Date(( ( (__readUInt32LE(blob,offset+4)/1e7)*Math.pow(2,32)+__readUInt32LE(blob,offset)/1e7 ) - 11644473600)*1000);
}

function read_file(filename/*:string*/, options/*:CFBReadOpts*/) {
	get_fs();
	return parse(fs.readFileSync(filename), options);
}

function read(blob/*:RawBytes|string*/, options/*:CFBReadOpts*/) {
	switch(options && options.type || "base64") {
		case "file": /*:: if(typeof blob !== 'string') throw "Must pass a filename when type='file'"; */return read_file(blob, options);
		case "base64": /*:: if(typeof blob !== 'string') throw "Must pass a base64-encoded binary string when type='file'"; */return parse(s2a(Base64.decode(blob)), options);
		case "binary": /*:: if(typeof blob !== 'string') throw "Must pass a binary string when type='file'"; */return parse(s2a(blob), options);
	}
	return parse(/*::typeof blob == 'string' ? new Buffer(blob, 'utf-8') : */blob, options);
}

function init_cfb(cfb/*:CFBContainer*/, opts/*:?any*/)/*:void*/ {
	var o = opts || {}, root = o.root || "Root Entry";
	if(!cfb.FullPaths) cfb.FullPaths = [];
	if(!cfb.FileIndex) cfb.FileIndex = [];
	if(cfb.FullPaths.length !== cfb.FileIndex.length) throw new Error("inconsistent CFB structure");
	if(cfb.FullPaths.length === 0) {
		cfb.FullPaths[0] = root + "/";
		cfb.FileIndex[0] = ({ name: root, type: 5 }/*:any*/);
	}
	if(o.CLSID) cfb.FileIndex[0].clsid = o.CLSID;
	seed_cfb(cfb);
}
function seed_cfb(cfb/*:CFBContainer*/)/*:void*/ {
	var nm = "\u0001Sh33tJ5";
	if(CFB.find(cfb, "/" + nm)) return;
	var p = new_buf(4); p[0] = 55; p[1] = p[3] = 50; p[2] = 54;
	cfb.FileIndex.push(({ name: nm, type: 2, content:p, size:4, L:69, R:69, C:69 }/*:any*/));
	cfb.FullPaths.push(cfb.FullPaths[0] + nm);
	rebuild_cfb(cfb);
}
function rebuild_cfb(cfb/*:CFBContainer*/, f/*:?boolean*/)/*:void*/ {
	init_cfb(cfb);
	var gc = false, s = false;
	for(var i = cfb.FullPaths.length - 1; i >= 0; --i) {
		var _file = cfb.FileIndex[i];
		switch(_file.type) {
			case 0:
				if(s) gc = true;
				else { cfb.FileIndex.pop(); cfb.FullPaths.pop(); }
				break;
			case 1: case 2: case 5:
				s = true;
				if(isNaN(_file.R * _file.L * _file.C)) gc = true;
				if(_file.R > -1 && _file.L > -1 && _file.R == _file.L) gc = true;
				break;
			default: gc = true; break;
		}
	}
	if(!gc && !f) return;

	var now = new Date(1987, 1, 19), j = 0;
	var data/*:Array<[string, CFBEntry]>*/ = [];
	for(i = 0; i < cfb.FullPaths.length; ++i) {
		if(cfb.FileIndex[i].type === 0) continue;
		data.push([cfb.FullPaths[i], cfb.FileIndex[i]]);
	}
	for(i = 0; i < data.length; ++i) {
		var dad = dirname(data[i][0]);
		s = false;
		for(j = 0; j < data.length; ++j) if(data[j][0] === dad) s = true;
		if(!s) data.push([dad, ({
			name: filename(dad).replace("/",""),
			type: 1,
			clsid: HEADER_CLSID,
			ct: now, mt: now,
			content: null
		}/*:any*/)]);
	}

	data.sort(function(x,y) { return namecmp(x[0], y[0]); });
	cfb.FullPaths = []; cfb.FileIndex = [];
	for(i = 0; i < data.length; ++i) { cfb.FullPaths[i] = data[i][0]; cfb.FileIndex[i] = data[i][1]; }
	for(i = 0; i < data.length; ++i) {
		var elt = cfb.FileIndex[i];
		var nm = cfb.FullPaths[i];

		elt.name =  filename(nm).replace("/","");
		elt.L = elt.R = elt.C = -(elt.color = 1);
		elt.size = elt.content ? elt.content.length : 0;
		elt.start = 0;
		elt.clsid = (elt.clsid || HEADER_CLSID);
		if(i === 0) {
			elt.C = data.length > 1 ? 1 : -1;
			elt.size = 0;
			elt.type = 5;
		} else if(nm.slice(-1) == "/") {
			for(j=i+1;j < data.length; ++j) if(dirname(cfb.FullPaths[j])==nm) break;
			elt.C = j >= data.length ? -1 : j;
			for(j=i+1;j < data.length; ++j) if(dirname(cfb.FullPaths[j])==dirname(nm)) break;
			elt.R = j >= data.length ? -1 : j;
			elt.type = 1;
		} else {
			if(dirname(cfb.FullPaths[i+1]||"") == dirname(nm)) elt.R = i + 1;
			elt.type = 2;
		}
	}

}

function _write(cfb/*:CFBContainer*/, options/*:CFBWriteOpts*/)/*:RawBytes*/ {
	var _opts = options || {};
	rebuild_cfb(cfb);
	if(_opts.fileType == 'zip') return write_zip(cfb, _opts);
	var L = (function(cfb/*:CFBContainer*/)/*:Array<number>*/{
		var mini_size = 0, fat_size = 0;
		for(var i = 0; i < cfb.FileIndex.length; ++i) {
			var file = cfb.FileIndex[i];
			if(!file.content) continue;
			/*:: if(file.content == null) throw new Error("unreachable"); */
			var flen = file.content.length;
			if(flen > 0){
				if(flen < 0x1000) mini_size += (flen + 0x3F) >> 6;
				else fat_size += (flen + 0x01FF) >> 9;
			}
		}
		var dir_cnt = (cfb.FullPaths.length +3) >> 2;
		var mini_cnt = (mini_size + 7) >> 3;
		var mfat_cnt = (mini_size + 0x7F) >> 7;
		var fat_base = mini_cnt + fat_size + dir_cnt + mfat_cnt;
		var fat_cnt = (fat_base + 0x7F) >> 7;
		var difat_cnt = fat_cnt <= 109 ? 0 : Math.ceil((fat_cnt-109)/0x7F);
		while(((fat_base + fat_cnt + difat_cnt + 0x7F) >> 7) > fat_cnt) difat_cnt = ++fat_cnt <= 109 ? 0 : Math.ceil((fat_cnt-109)/0x7F);
		var L =  [1, difat_cnt, fat_cnt, mfat_cnt, dir_cnt, fat_size, mini_size, 0];
		cfb.FileIndex[0].size = mini_size << 6;
		L[7] = (cfb.FileIndex[0].start=L[0]+L[1]+L[2]+L[3]+L[4]+L[5])+((L[6]+7) >> 3);
		return L;
	})(cfb);
	var o = new_buf(L[7] << 9);
	var i = 0, T = 0;
	{
		for(i = 0; i < 8; ++i) o.write_shift(1, HEADER_SIG[i]);
		for(i = 0; i < 8; ++i) o.write_shift(2, 0);
		o.write_shift(2, 0x003E);
		o.write_shift(2, 0x0003);
		o.write_shift(2, 0xFFFE);
		o.write_shift(2, 0x0009);
		o.write_shift(2, 0x0006);
		for(i = 0; i < 3; ++i) o.write_shift(2, 0);
		o.write_shift(4, 0);
		o.write_shift(4, L[2]);
		o.write_shift(4, L[0] + L[1] + L[2] + L[3] - 1);
		o.write_shift(4, 0);
		o.write_shift(4, 1<<12);
		o.write_shift(4, L[3] ? L[0] + L[1] + L[2] - 1: ENDOFCHAIN);
		o.write_shift(4, L[3]);
		o.write_shift(-4, L[1] ? L[0] - 1: ENDOFCHAIN);
		o.write_shift(4, L[1]);
		for(i = 0; i < 109; ++i) o.write_shift(-4, i < L[2] ? L[1] + i : -1);
	}
	if(L[1]) {
		for(T = 0; T < L[1]; ++T) {
			for(; i < 236 + T * 127; ++i) o.write_shift(-4, i < L[2] ? L[1] + i : -1);
			o.write_shift(-4, T === L[1] - 1 ? ENDOFCHAIN : T + 1);
		}
	}
	var chainit = function(w/*:number*/)/*:void*/ {
		for(T += w; i<T-1; ++i) o.write_shift(-4, i+1);
		if(w) { ++i; o.write_shift(-4, ENDOFCHAIN); }
	};
	T = i = 0;
	for(T+=L[1]; i<T; ++i) o.write_shift(-4, consts.DIFSECT);
	for(T+=L[2]; i<T; ++i) o.write_shift(-4, consts.FATSECT);
	chainit(L[3]);
	chainit(L[4]);
	var j/*:number*/ = 0, flen/*:number*/ = 0;
	var file/*:CFBEntry*/ = cfb.FileIndex[0];
	for(; j < cfb.FileIndex.length; ++j) {
		file = cfb.FileIndex[j];
		if(!file.content) continue;
		/*:: if(file.content == null) throw new Error("unreachable"); */
		flen = file.content.length;
		if(flen < 0x1000) continue;
		file.start = T;
		chainit((flen + 0x01FF) >> 9);
	}
	chainit((L[6] + 7) >> 3);
	while(o.l & 0x1FF) o.write_shift(-4, consts.ENDOFCHAIN);
	T = i = 0;
	for(j = 0; j < cfb.FileIndex.length; ++j) {
		file = cfb.FileIndex[j];
		if(!file.content) continue;
		/*:: if(file.content == null) throw new Error("unreachable"); */
		flen = file.content.length;
		if(!flen || flen >= 0x1000) continue;
		file.start = T;
		chainit((flen + 0x3F) >> 6);
	}
	while(o.l & 0x1FF) o.write_shift(-4, consts.ENDOFCHAIN);
	for(i = 0; i < L[4]<<2; ++i) {
		var nm = cfb.FullPaths[i];
		if(!nm || nm.length === 0) {
			for(j = 0; j < 17; ++j) o.write_shift(4, 0);
			for(j = 0; j < 3; ++j) o.write_shift(4, -1);
			for(j = 0; j < 12; ++j) o.write_shift(4, 0);
			continue;
		}
		file = cfb.FileIndex[i];
		if(i === 0) file.start = file.size ? file.start - 1 : ENDOFCHAIN;
		var _nm/*:string*/ = (i === 0 && _opts.root) || file.name;
		flen = 2*(_nm.length+1);
		o.write_shift(64, _nm, "utf16le");
		o.write_shift(2, flen);
		o.write_shift(1, file.type);
		o.write_shift(1, file.color);
		o.write_shift(-4, file.L);
		o.write_shift(-4, file.R);
		o.write_shift(-4, file.C);
		if(!file.clsid) for(j = 0; j < 4; ++j) o.write_shift(4, 0);
		else o.write_shift(16, file.clsid, "hex");
		o.write_shift(4, file.state || 0);
		o.write_shift(4, 0); o.write_shift(4, 0);
		o.write_shift(4, 0); o.write_shift(4, 0);
		o.write_shift(4, file.start);
		o.write_shift(4, file.size); o.write_shift(4, 0);
	}
	for(i = 1; i < cfb.FileIndex.length; ++i) {
		file = cfb.FileIndex[i];
		/*:: if(!file.content) throw new Error("unreachable"); */
		if(file.size >= 0x1000) {
			o.l = (file.start+1) << 9;
			for(j = 0; j < file.size; ++j) o.write_shift(1, file.content[j]);
			for(; j & 0x1FF; ++j) o.write_shift(1, 0);
		}
	}
	for(i = 1; i < cfb.FileIndex.length; ++i) {
		file = cfb.FileIndex[i];
		/*:: if(!file.content) throw new Error("unreachable"); */
		if(file.size > 0 && file.size < 0x1000) {
			for(j = 0; j < file.size; ++j) o.write_shift(1, file.content[j]);
			for(; j & 0x3F; ++j) o.write_shift(1, 0);
		}
	}
	while(o.l < o.length) o.write_shift(1, 0);
	return o;
}
/* [MS-CFB] 2.6.4 (Unicode 3.0.1 case conversion) */
function find(cfb/*:CFBContainer*/, path/*:string*/)/*:?CFBEntry*/ {
	var UCFullPaths/*:Array<string>*/ = cfb.FullPaths.map(function(x) { return x.toUpperCase(); });
	var UCPaths/*:Array<string>*/ = UCFullPaths.map(function(x) { var y = x.split("/"); return y[y.length - (x.slice(-1) == "/" ? 2 : 1)]; });
	var k/*:boolean*/ = false;
	if(path.charCodeAt(0) === 47 /* "/" */) { k = true; path = UCFullPaths[0].slice(0, -1) + path; }
	else k = path.indexOf("/") !== -1;
	var UCPath/*:string*/ = path.toUpperCase();
	var w/*:number*/ = k === true ? UCFullPaths.indexOf(UCPath) : UCPaths.indexOf(UCPath);
	if(w !== -1) return cfb.FileIndex[w];

	var m = !UCPath.match(chr1);
	UCPath = UCPath.replace(chr0,'');
	if(m) UCPath = UCPath.replace(chr1,'!');
	for(w = 0; w < UCFullPaths.length; ++w) {
		if((m ? UCFullPaths[w].replace(chr1,'!') : UCFullPaths[w]).replace(chr0,'') == UCPath) return cfb.FileIndex[w];
		if((m ? UCPaths[w].replace(chr1,'!') : UCPaths[w]).replace(chr0,'') == UCPath) return cfb.FileIndex[w];
	}
	return null;
}
/** CFB Constants */
var MSSZ = 64; /* Mini Sector Size = 1<<6 */
//var MSCSZ = 4096; /* Mini Stream Cutoff Size */
/* 2.1 Compound File Sector Numbers and Types */
var ENDOFCHAIN = -2;
/* 2.2 Compound File Header */
var HEADER_SIGNATURE = 'd0cf11e0a1b11ae1';
var HEADER_SIG = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
var HEADER_CLSID = '00000000000000000000000000000000';
var consts = {
	/* 2.1 Compund File Sector Numbers and Types */
	MAXREGSECT: -6,
	DIFSECT: -4,
	FATSECT: -3,
	ENDOFCHAIN: ENDOFCHAIN,
	FREESECT: -1,
	/* 2.2 Compound File Header */
	HEADER_SIGNATURE: HEADER_SIGNATURE,
	HEADER_MINOR_VERSION: '3e00',
	MAXREGSID: -6,
	NOSTREAM: -1,
	HEADER_CLSID: HEADER_CLSID,
	/* 2.6.1 Compound File Directory Entry */
	EntryTypes: ['unknown','storage','stream','lockbytes','property','root']
};

function write_file(cfb/*:CFBContainer*/, filename/*:string*/, options/*:CFBWriteOpts*/)/*:void*/ {
	get_fs();
	var o = _write(cfb, options);
	/*:: if(typeof Buffer == 'undefined' || !Buffer.isBuffer(o) || !(o instanceof Buffer)) throw new Error("unreachable"); */
	fs.writeFileSync(filename, o);
}

function a2s(o/*:RawBytes*/)/*:string*/ {
	var out = new Array(o.length);
	for(var i = 0; i < o.length; ++i) out[i] = String.fromCharCode(o[i]);
	return out.join("");
}

function write(cfb/*:CFBContainer*/, options/*:CFBWriteOpts*/)/*:RawBytes|string*/ {
	var o = _write(cfb, options);
	switch(options && options.type) {
		case "file": get_fs(); fs.writeFileSync(options.filename, (o/*:any*/)); return o;
		case "binary": return a2s(o);
		case "base64": return Base64.encode(a2s(o));
	}
	return o;
}
/* node < 8.1 zlib does not expose bytesRead, so default to pure JS */
var _zlib;
function use_zlib(zlib) { try {
	var InflateRaw = zlib.InflateRaw;
	var InflRaw = new InflateRaw();
	InflRaw._processChunk(new Uint8Array([3, 0]), InflRaw._finishFlushFlag);
	if(InflRaw.bytesRead) _zlib = zlib;
	else throw new Error("zlib does not expose bytesRead");
} catch(e) {console.error("cannot use native zlib: " + (e.message || e)); } }

function _inflateRawSync(payload, usz) {
	if(!_zlib) return _inflate(payload, usz);
	var InflateRaw = _zlib.InflateRaw;
	var InflRaw = new InflateRaw();
	var out = InflRaw._processChunk(payload.slice(payload.l), InflRaw._finishFlushFlag);
	payload.l += InflRaw.bytesRead;
	return out;
}

function _deflateRawSync(payload) {
	return _zlib ? _zlib.deflateRawSync(payload) : _deflate(payload);
}
var CLEN_ORDER = [ 16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15 ];

/*  LEN_ID = [ 257, 258, 259, 260, 261, 262, 263, 264, 265, 266, 267, 268, 269, 270, 271, 272, 273, 274, 275, 276, 277, 278, 279, 280, 281, 282, 283, 284, 285 ]; */
var LEN_LN = [   3,   4,   5,   6,   7,   8,   9,  10,  11,  13 , 15,  17,  19,  23,  27,  31,  35,  43,  51,  59,  67,  83,  99, 115, 131, 163, 195, 227, 258 ];

/*  DST_ID = [  0,  1,  2,  3,  4,  5,  6,  7,  8,  9, 10, 11, 12, 13,  14,  15,  16,  17,  18,  19,   20,   21,   22,   23,   24,   25,   26,    27,    28,    29 ]; */
var DST_LN = [  1,  2,  3,  4,  5,  7,  9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577 ];

function bit_swap_8(n) { var t = (((((n<<1)|(n<<11)) & 0x22110) | (((n<<5)|(n<<15)) & 0x88440))); return ((t>>16) | (t>>8) |t)&0xFF; }

var use_typed_arrays = typeof Uint8Array !== 'undefined';

var bitswap8 = use_typed_arrays ? new Uint8Array(1<<8) : [];
for(var q = 0; q < (1<<8); ++q) bitswap8[q] = bit_swap_8(q);

function bit_swap_n(n, b) {
	var rev = bitswap8[n & 0xFF];
	if(b <= 8) return rev >>> (8-b);
	rev = (rev << 8) | bitswap8[(n>>8)&0xFF];
	if(b <= 16) return rev >>> (16-b);
	rev = (rev << 8) | bitswap8[(n>>16)&0xFF];
	return rev >>> (24-b);
}

/* helpers for unaligned bit reads */
function read_bits_2(buf, bl) { var w = (bl&7), h = (bl>>>3); return ((buf[h]|(w <= 6 ? 0 : buf[h+1]<<8))>>>w)& 0x03; }
function read_bits_3(buf, bl) { var w = (bl&7), h = (bl>>>3); return ((buf[h]|(w <= 5 ? 0 : buf[h+1]<<8))>>>w)& 0x07; }
function read_bits_4(buf, bl) { var w = (bl&7), h = (bl>>>3); return ((buf[h]|(w <= 4 ? 0 : buf[h+1]<<8))>>>w)& 0x0F; }
function read_bits_5(buf, bl) { var w = (bl&7), h = (bl>>>3); return ((buf[h]|(w <= 3 ? 0 : buf[h+1]<<8))>>>w)& 0x1F; }
function read_bits_7(buf, bl) { var w = (bl&7), h = (bl>>>3); return ((buf[h]|(w <= 1 ? 0 : buf[h+1]<<8))>>>w)& 0x7F; }

/* works up to n = 3 * 8 + 1 = 25 */
function read_bits_n(buf, bl, n) {
	var w = (bl&7), h = (bl>>>3), f = ((1<<n)-1);
	var v = buf[h] >>> w;
	if(n < 8 - w) return v & f;
	v |= buf[h+1]<<(8-w);
	if(n < 16 - w) return v & f;
	v |= buf[h+2]<<(16-w);
	if(n < 24 - w) return v & f;
	v |= buf[h+3]<<(24-w);
	return v & f;
}

/* until ArrayBuffer#realloc is a thing, fake a realloc */
function realloc(b, sz/*:number*/) {
	var L = b.length, M = 2*L > sz ? 2*L : sz + 5, i = 0;
	if(L >= sz) return b;
	if(has_buf) {
		var o = new_unsafe_buf(M);
		// $FlowIgnore
		if(b.copy) b.copy(o);
		else for(; i < b.length; ++i) o[i] = b[i];
		return o;
	} else if(use_typed_arrays) {
		var a = new Uint8Array(M);
		if(a.set) a.set(b);
		else for(; i < b.length; ++i) a[i] = b[i];
		return a;
	}
	b.length = M;
	return b;
}

/* zero-filled arrays for older browsers */
function zero_fill_array(n) {
	var o = new Array(n);
	for(var i = 0; i < n; ++i) o[i] = 0;
	return o;
}var _deflate = (function() {
var _deflateRaw = (function() {
	return function deflateRaw(data, out) {
		var boff = 0;
		while(boff < data.length) {
			var L = Math.min(0xFFFF, data.length - boff);
			var h = boff + L == data.length;
			/* TODO: this is only type 0 stored */
			out.write_shift(1, +h);
			out.write_shift(2, L);
			out.write_shift(2, (~L) & 0xFFFF);
			while(L-- > 0) out[out.l++] = data[boff++];
		}
		return out.l;
	};
})();

return function(data) {
	var buf = new_buf(50+Math.floor(data.length*1.1));
	var off = _deflateRaw(data, buf);
	return buf.slice(0, off);
};
})();
/* modified inflate function also moves original read head */

/* build tree (used for literals and lengths) */
function build_tree(clens, cmap, MAX/*:number*/)/*:number*/ {
	var maxlen = 1, w = 0, i = 0, j = 0, ccode = 0, L = clens.length;

	var bl_count  = use_typed_arrays ? new Uint16Array(32) : zero_fill_array(32);
	for(i = 0; i < 32; ++i) bl_count[i] = 0;

	for(i = L; i < MAX; ++i) clens[i] = 0;
	L = clens.length;

	var ctree = use_typed_arrays ? new Uint16Array(L) : zero_fill_array(L); // []

	/* build code tree */
	for(i = 0; i < L; ++i) {
		bl_count[(w = clens[i])]++;
		if(maxlen < w) maxlen = w;
		ctree[i] = 0;
	}
	bl_count[0] = 0;
	for(i = 1; i <= maxlen; ++i) bl_count[i+16] = (ccode = (ccode + bl_count[i-1])<<1);
	for(i = 0; i < L; ++i) {
		ccode = clens[i];
		if(ccode != 0) ctree[i] = bl_count[ccode+16]++;
	}

	/* cmap[maxlen + 4 bits] = (off&15) + (lit<<4) reverse mapping */
	var cleni = 0;
	for(i = 0; i < L; ++i) {
		cleni = clens[i];
		if(cleni != 0) {
			ccode = bit_swap_n(ctree[i], maxlen)>>(maxlen-cleni);
			for(j = (1<<(maxlen + 4 - cleni)) - 1; j>=0; --j)
				cmap[ccode|(j<<cleni)] = (cleni&15) | (i<<4);
		}
	}
	return maxlen;
}

var fix_lmap = use_typed_arrays ? new Uint16Array(512) : zero_fill_array(512);
var fix_dmap = use_typed_arrays ? new Uint16Array(32)  : zero_fill_array(32);
if(!use_typed_arrays) {
	for(var i = 0; i < 512; ++i) fix_lmap[i] = 0;
	for(i = 0; i < 32; ++i) fix_dmap[i] = 0;
}
(function() {
	var dlens/*:Array<number>*/ = [];
	var i = 0;
	for(;i<32; i++) dlens.push(5);
	build_tree(dlens, fix_dmap, 32);

	var clens/*:Array<number>*/ = [];
	i = 0;
	for(; i<=143; i++) clens.push(8);
	for(; i<=255; i++) clens.push(9);
	for(; i<=279; i++) clens.push(7);
	for(; i<=287; i++) clens.push(8);
	build_tree(clens, fix_lmap, 288);
})();

var dyn_lmap = use_typed_arrays ? new Uint16Array(32768) : zero_fill_array(32768);
var dyn_dmap = use_typed_arrays ? new Uint16Array(32768) : zero_fill_array(32768);
var dyn_cmap = use_typed_arrays ? new Uint16Array(128)   : zero_fill_array(128);
var dyn_len_1 = 1, dyn_len_2 = 1;

/* 5.5.3 Expanding Huffman Codes */
function dyn(data, boff/*:number*/) {
	/* nomenclature from RFC1951 refers to bit values; these are offset by the implicit constant */
	var _HLIT = read_bits_5(data, boff) + 257; boff += 5;
	var _HDIST = read_bits_5(data, boff) + 1; boff += 5;
	var _HCLEN = read_bits_4(data, boff) + 4; boff += 4;
	var w = 0;

	/* grab and store code lengths */
	var clens = use_typed_arrays ? new Uint8Array(19) : zero_fill_array(19);
	var ctree = [ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 ];
	var maxlen = 1;
	var bl_count =  use_typed_arrays ? new Uint8Array(8) : zero_fill_array(8);
	var next_code = use_typed_arrays ? new Uint8Array(8) : zero_fill_array(8);
	var L = clens.length; /* 19 */
	for(var i = 0; i < _HCLEN; ++i) {
		clens[CLEN_ORDER[i]] = w = read_bits_3(data, boff);
		if(maxlen < w) maxlen = w;
		bl_count[w]++;
		boff += 3;
	}

	/* build code tree */
	var ccode = 0;
	bl_count[0] = 0;
	for(i = 1; i <= maxlen; ++i) next_code[i] = ccode = (ccode + bl_count[i-1])<<1;
	for(i = 0; i < L; ++i) if((ccode = clens[i]) != 0) ctree[i] = next_code[ccode]++;
	/* cmap[7 bits from stream] = (off&7) + (lit<<3) */
	var cleni = 0;
	for(i = 0; i < L; ++i) {
		cleni = clens[i];
		if(cleni != 0) {
			ccode = bitswap8[ctree[i]]>>(8-cleni);
			for(var j = (1<<(7-cleni))-1; j>=0; --j) dyn_cmap[ccode|(j<<cleni)] = (cleni&7) | (i<<3);
		}
	}

	/* read literal and dist codes at once */
	var hcodes/*:Array<number>*/ = [];
	maxlen = 1;
	for(; hcodes.length < _HLIT + _HDIST;) {
		ccode = dyn_cmap[read_bits_7(data, boff)];
		boff += ccode & 7;
		switch((ccode >>>= 3)) {
			case 16:
				w = 3 + read_bits_2(data, boff); boff += 2;
				ccode = hcodes[hcodes.length - 1];
				while(w-- > 0) hcodes.push(ccode);
				break;
			case 17:
				w = 3 + read_bits_3(data, boff); boff += 3;
				while(w-- > 0) hcodes.push(0);
				break;
			case 18:
				w = 11 + read_bits_7(data, boff); boff += 7;
				while(w -- > 0) hcodes.push(0);
				break;
			default:
				hcodes.push(ccode);
				if(maxlen < ccode) maxlen = ccode;
				break;
		}
	}

	/* build literal / length trees */
	var h1 = hcodes.slice(0, _HLIT), h2 = hcodes.slice(_HLIT);
	for(i = _HLIT; i < 286; ++i) h1[i] = 0;
	for(i = _HDIST; i < 30; ++i) h2[i] = 0;
	dyn_len_1 = build_tree(h1, dyn_lmap, 286);
	dyn_len_2 = build_tree(h2, dyn_dmap, 30);
	return boff;
}

/* return [ data, bytesRead ] */
function inflate(data, usz/*:number*/) {
	/* shortcircuit for empty buffer [0x03, 0x00] */
	if(data[0] == 3 && !(data[1] & 0x3)) { return [new_raw_buf(usz), 2]; }

	/* bit offset */
	var boff = 0;

	/* header includes final bit and type bits */
	var header = 0;

	var outbuf = new_unsafe_buf(usz ? usz : (1<<18));
	var woff = 0;
	var OL = outbuf.length>>>0;
	var max_len_1 = 0, max_len_2 = 0;

	while((header&1) == 0) {
		header = read_bits_3(data, boff); boff += 3;
		if((header >>> 1) == 0) {
			/* Stored block */
			if(boff & 7) boff += 8 - (boff&7);
			/* 2 bytes sz, 2 bytes bit inverse */
			var sz = data[boff>>>3] | data[(boff>>>3)+1]<<8;
			boff += 32;
			/* push sz bytes */
			if(!usz && OL < woff + sz) { outbuf = realloc(outbuf, woff + sz); OL = outbuf.length; }
			if(typeof data.copy === 'function') {
				// $FlowIgnore
				data.copy(outbuf, woff, boff>>>3, (boff>>>3)+sz);
				woff += sz; boff += 8*sz;
			} else while(sz-- > 0) { outbuf[woff++] = data[boff>>>3]; boff += 8; }
			continue;
		} else if((header >>> 1) == 1) {
			/* Fixed Huffman */
			max_len_1 = 9; max_len_2 = 5;
		} else {
			/* Dynamic Huffman */
			boff = dyn(data, boff);
			max_len_1 = dyn_len_1; max_len_2 = dyn_len_2;
		}
		if(!usz && (OL < woff + 32767)) { outbuf = realloc(outbuf, woff + 32767); OL = outbuf.length; }
		for(;;) { // while(true) is apparently out of vogue in modern JS circles
			/* ingest code and move read head */
			var bits = read_bits_n(data, boff, max_len_1);
			var code = (header>>>1) == 1 ? fix_lmap[bits] : dyn_lmap[bits];
			boff += code & 15; code >>>= 4;
			/* 0-255 are literals, 256 is end of block token, 257+ are copy tokens */
			if(((code>>>8)&0xFF) === 0) outbuf[woff++] = code;
			else if(code == 256) break;
			else {
				code -= 257;
				var len_eb = (code < 8) ? 0 : ((code-4)>>2); if(len_eb > 5) len_eb = 0;
				var tgt = woff + LEN_LN[code];
				/* length extra bits */
				if(len_eb > 0) {
					tgt += read_bits_n(data, boff, len_eb);
					boff += len_eb;
				}

				/* dist code */
				bits = read_bits_n(data, boff, max_len_2);
				code = (header>>>1) == 1 ? fix_dmap[bits] : dyn_dmap[bits];
				boff += code & 15; code >>>= 4;
				var dst_eb = (code < 4 ? 0 : (code-2)>>1);
				var dst = DST_LN[code];
				/* dist extra bits */
				if(dst_eb > 0) {
					dst += read_bits_n(data, boff, dst_eb);
					boff += dst_eb;
				}

				/* in the common case, manual byte copy is faster than TA set / Buffer copy */
				if(!usz && OL < tgt) { outbuf = realloc(outbuf, tgt); OL = outbuf.length; }
				while(woff < tgt) { outbuf[woff] = outbuf[woff - dst]; ++woff; }
			}
		}
	}
	return [usz ? outbuf : outbuf.slice(0, woff), (boff+7)>>>3];
}

function _inflate(payload, usz) {
	var data = payload.slice(payload.l||0);
	var out = inflate(data, usz);
	payload.l += out[1];
	return out[0];
}

function warn_or_throw(wrn, msg) {
	if(wrn) { if(typeof console !== 'undefined') console.error(msg); }
	else throw new Error(msg);
}

function parse_zip(file/*:RawBytes*/, options/*:CFBReadOpts*/)/*:CFBContainer*/ {
	var blob/*:CFBlob*/ = /*::(*/file/*:: :any)*/;
	prep_blob(blob, 0);

	var FileIndex/*:CFBFileIndex*/ = [], FullPaths/*:Array<string>*/ = [];
	var o = {
		FileIndex: FileIndex,
		FullPaths: FullPaths
	};
	init_cfb(o, { root: options.root });

	/* find end of central directory, start just after signature */
	var i = blob.length - 4;
	while((blob[i] != 0x50 || blob[i+1] != 0x4b || blob[i+2] != 0x05 || blob[i+3] != 0x06) && i >= 0) --i;
	blob.l = i + 4;

	/* parse end of central directory */
	blob.l += 4;
	var fcnt = blob.read_shift(2);
	blob.l += 6;
	var start_cd = blob.read_shift(4);

	/* parse central directory */
	blob.l = start_cd;

	for(i = 0; i < fcnt; ++i) {
		/* trust local file header instead of CD entry */
		blob.l += 20;
		var csz = blob.read_shift(4);
		var usz = blob.read_shift(4);
		var namelen = blob.read_shift(2);
		var efsz = blob.read_shift(2);
		var fcsz = blob.read_shift(2);
		blob.l += 8;
		var offset = blob.read_shift(4);
		var EF = parse_extra_field(/*::(*/blob.slice(blob.l+namelen, blob.l+namelen+efsz)/*:: :any)*/);
		blob.l += namelen + efsz + fcsz;

		var L = blob.l;
		blob.l = offset + 4;
		parse_local_file(blob, csz, usz, o, EF);
		blob.l = L;
	}

	return o;
}


/* head starts just after local file header signature */
function parse_local_file(blob/*:CFBlob*/, csz/*:number*/, usz/*:number*/, o/*:CFBContainer*/, EF) {
	/* [local file header] */
	blob.l += 2;
	var flags = blob.read_shift(2);
	var meth = blob.read_shift(2);
	var date = parse_dos_date(blob);

	if(flags & 0x2041) throw new Error("Unsupported ZIP encryption");
	var crc32 = blob.read_shift(4);
	var _csz = blob.read_shift(4);
	var _usz = blob.read_shift(4);

	var namelen = blob.read_shift(2);
	var efsz = blob.read_shift(2);

	// TODO: flags & (1<<11) // UTF8
	var name = ""; for(var i = 0; i < namelen; ++i) name += String.fromCharCode(blob[blob.l++]);
	if(efsz) {
		var ef = parse_extra_field(/*::(*/blob.slice(blob.l, blob.l + efsz)/*:: :any)*/);
		if((ef[0x5455]||{}).mt) date = ef[0x5455].mt;
		if(((EF||{})[0x5455]||{}).mt) date = EF[0x5455].mt;
	}
	blob.l += efsz;

	/* [encryption header] */

	/* [file data] */
	var data = blob.slice(blob.l, blob.l + _csz);
	switch(meth) {
		case 8: data = _inflateRawSync(blob, _usz); break;
		case 0: break;
		default: throw new Error("Unsupported ZIP Compression method " + meth);
	}

	/* [data descriptor] */
	var wrn = false;
	if(flags & 8) {
		crc32 = blob.read_shift(4);
		if(crc32 == 0x08074b50) { crc32 = blob.read_shift(4); wrn = true; }
		_csz = blob.read_shift(4);
		_usz = blob.read_shift(4);
	}

	if(_csz != csz) warn_or_throw(wrn, "Bad compressed size: " + csz + " != " + _csz);
	if(_usz != usz) warn_or_throw(wrn, "Bad uncompressed size: " + usz + " != " + _usz);
	var _crc32 = CRC32.buf(data, 0);
	if(crc32 != _crc32) warn_or_throw(wrn, "Bad CRC32 checksum: " + crc32 + " != " + _crc32);
	cfb_add(o, name, data, {unsafe: true, mt: date});
}
function write_zip(cfb/*:CFBContainer*/, options/*:CFBWriteOpts*/)/*:RawBytes*/ {
	var _opts = options || {};
	var out = [], cdirs = [];
	var o/*:CFBlob*/ = new_buf(1);
	var method = (_opts.compression ? 8 : 0), flags = 0;
	var desc = false;
	if(desc) flags |= 8;
	var i = 0, j = 0;

	var start_cd = 0, fcnt = 0;
	var root = cfb.FullPaths[0], fp = root, fi = cfb.FileIndex[0];
	var crcs = [];
	var sz_cd = 0;

	for(i = 1; i < cfb.FullPaths.length; ++i) {
		fp = cfb.FullPaths[i].slice(root.length); fi = cfb.FileIndex[i];
		if(!fi.size || !fi.content || fp == "\u0001Sh33tJ5") continue;
		var start = start_cd;

		/* TODO: CP437 filename */
		var namebuf = new_buf(fp.length);
		for(j = 0; j < fp.length; ++j) namebuf.write_shift(1, fp.charCodeAt(j) & 0x7F);
		namebuf = namebuf.slice(0, namebuf.l);
		crcs[fcnt] = CRC32.buf(/*::((*/fi.content/*::||[]):any)*/, 0);

		var outbuf = fi.content/*::||[]*/;
		if(method == 8) outbuf = _deflateRawSync(outbuf);

		/* local file header */
		o = new_buf(30);
		o.write_shift(4, 0x04034b50);
		o.write_shift(2, 20);
		o.write_shift(2, flags);
		o.write_shift(2, method);
		/* TODO: last mod file time/date */
		if(fi.mt) write_dos_date(o, fi.mt);
		else o.write_shift(4, 0);
		o.write_shift(-4, (flags & 8) ? 0 : crcs[fcnt]);
		o.write_shift(4,  (flags & 8) ? 0 : outbuf.length);
		o.write_shift(4,  (flags & 8) ? 0 : /*::(*/fi.content/*::||[])*/.length);
		o.write_shift(2, namebuf.length);
		o.write_shift(2, 0);

		start_cd += o.length;
		out.push(o);
		start_cd += namebuf.length;
		out.push(namebuf);

		/* TODO: encryption header ? */
		start_cd += outbuf.length;
		out.push(outbuf);

		/* data descriptor */
		if(flags & 8) {
			o = new_buf(12);
			o.write_shift(-4, crcs[fcnt]);
			o.write_shift(4, outbuf.length);
			o.write_shift(4, /*::(*/fi.content/*::||[])*/.length);
			start_cd += o.l;
			out.push(o);
		}

		/* central directory */
		o = new_buf(46);
		o.write_shift(4, 0x02014b50);
		o.write_shift(2, 0);
		o.write_shift(2, 20);
		o.write_shift(2, flags);
		o.write_shift(2, method);
		o.write_shift(4, 0); /* TODO: last mod file time/date */
		o.write_shift(-4, crcs[fcnt]);

		o.write_shift(4, outbuf.length);
		o.write_shift(4, /*::(*/fi.content/*::||[])*/.length);
		o.write_shift(2, namebuf.length);
		o.write_shift(2, 0);
		o.write_shift(2, 0);
		o.write_shift(2, 0);
		o.write_shift(2, 0);
		o.write_shift(4, 0);
		o.write_shift(4, start);

		sz_cd += o.l;
		cdirs.push(o);
		sz_cd += namebuf.length;
		cdirs.push(namebuf);
		++fcnt;
	}

	/* end of central directory */
	o = new_buf(22);
	o.write_shift(4, 0x06054b50);
	o.write_shift(2, 0);
	o.write_shift(2, 0);
	o.write_shift(2, fcnt);
	o.write_shift(2, fcnt);
	o.write_shift(4, sz_cd);
	o.write_shift(4, start_cd);
	o.write_shift(2, 0);

	return bconcat(([bconcat((out/*:any*/)), bconcat(cdirs), o]/*:any*/));
}
function cfb_new(opts/*:?any*/)/*:CFBContainer*/ {
	var o/*:CFBContainer*/ = ({}/*:any*/);
	init_cfb(o, opts);
	return o;
}

function cfb_add(cfb/*:CFBContainer*/, name/*:string*/, content/*:?RawBytes*/, opts/*:?any*/)/*:CFBEntry*/ {
	var unsafe = opts && opts.unsafe;
	if(!unsafe) init_cfb(cfb);
	var file = !unsafe && CFB.find(cfb, name);
	if(!file) {
		var fpath/*:string*/ = cfb.FullPaths[0];
		if(name.slice(0, fpath.length) == fpath) fpath = name;
		else {
			if(fpath.slice(-1) != "/") fpath += "/";
			fpath = (fpath + name).replace("//","/");
		}
		file = ({name: filename(name), type: 2}/*:any*/);
		cfb.FileIndex.push(file);
		cfb.FullPaths.push(fpath);
		if(!unsafe) CFB.utils.cfb_gc(cfb);
	}
	/*:: if(!file) throw new Error("unreachable"); */
	file.content = (content/*:any*/);
	file.size = content ? content.length : 0;
	if(opts) {
		if(opts.CLSID) file.clsid = opts.CLSID;
		if(opts.mt) file.mt = opts.mt;
		if(opts.ct) file.ct = opts.ct;
	}
	return file;
}

function cfb_del(cfb/*:CFBContainer*/, name/*:string*/)/*:boolean*/ {
	init_cfb(cfb);
	var file = CFB.find(cfb, name);
	if(file) for(var j = 0; j < cfb.FileIndex.length; ++j) if(cfb.FileIndex[j] == file) {
		cfb.FileIndex.splice(j, 1);
		cfb.FullPaths.splice(j, 1);
		return true;
	}
	return false;
}

function cfb_mov(cfb/*:CFBContainer*/, old_name/*:string*/, new_name/*:string*/)/*:boolean*/ {
	init_cfb(cfb);
	var file = CFB.find(cfb, old_name);
	if(file) for(var j = 0; j < cfb.FileIndex.length; ++j) if(cfb.FileIndex[j] == file) {
		cfb.FileIndex[j].name = filename(new_name);
		cfb.FullPaths[j] = new_name;
		return true;
	}
	return false;
}

function cfb_gc(cfb/*:CFBContainer*/)/*:void*/ { rebuild_cfb(cfb, true); }

exports.find = find;
exports.read = read;
exports.parse = parse;
exports.write = write;
exports.writeFile = write_file;
exports.utils = {
	cfb_new: cfb_new,
	cfb_add: cfb_add,
	cfb_del: cfb_del,
	cfb_mov: cfb_mov,
	cfb_gc: cfb_gc,
	ReadShift: ReadShift,
	CheckField: CheckField,
	prep_blob: prep_blob,
	bconcat: bconcat,
	use_zlib: use_zlib,
	_deflateRaw: _deflate,
	_inflateRaw: _inflate,
	consts: consts
};

return exports;
})();

if(typeof require !== 'undefined' && typeof module !== 'undefined' && typeof DO_NOT_EXPORT_CFB === 'undefined') { module.exports = CFB; }
