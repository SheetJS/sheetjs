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
*/
/* cfb.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*jshint eqnull:true */

/*::
declare var DO_NOT_EXPORT_CFB:any;
type SectorEntry = any;
type SectorList = {
	(k:string|number):SectorEntry;
	name:?string;
	fat_addrs:any;
	ssz:number;
}
*/
/* [MS-CFB] v20130118 */
var CFB = (function _CFB(){
var exports = {};
exports.version = '0.11.0';
function parse(file) {
var mver = 3; // major version
var ssz = 512; // sector size
var nmfs = 0; // number of mini FAT sectors
var ndfs = 0; // number of DIFAT sectors
var dir_start = 0; // first directory sector location
var minifat_start = 0; // first mini FAT sector location
var difat_start = 0; // first mini FAT sector location

var fat_addrs = []; // locations of FAT sectors

/* [MS-CFB] 2.2 Compound File Header */
var blob/*:any*/ = file.slice(0,512);
prep_blob(blob, 0);

/* major version */
var mv = check_get_mver(blob);
mver = mv[0];
switch(mver) {
	case 3: ssz = 512; break; case 4: ssz = 4096; break;
	default: throw "Major Version: Expected 3 or 4 saw " + mver;
}

/* reprocess header */
if(ssz !== 512) { blob = file.slice(0,ssz); prep_blob(blob, 28 /* blob.l */); }
/* Save header for final object */
var header = file.slice(0,ssz);

check_shifts(blob, mver);

// Number of Directory Sectors
var nds = blob.read_shift(4, 'i');
if(mver === 3 && nds !== 0) throw '# Directory Sectors: Expected 0 saw ' + nds;

// Number of FAT Sectors
//var nfs = blob.read_shift(4, 'i');
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
ndfs = blob.read_shift(4, 'i');

// Grab FAT Sector Locations
for(var q, j = 0; j < 109; ++j) { /* 109 = (512 - blob.l)>>>2; */
	q = blob.read_shift(4, 'i');
	if(q<0) break;
	fat_addrs[j] = q;
}

/** Break the file up into sectors */
var sectors = sectorify(file, ssz);

sleuth_fat(difat_start, ndfs, sectors, ssz, fat_addrs);

/** Chains */
var sector_list/*:SectorList*/ = make_sector_list(sectors, dir_start, fat_addrs, ssz);

sector_list[dir_start].name = "!Directory";
if(nmfs > 0 && minifat_start !== ENDOFCHAIN) sector_list[minifat_start].name = "!MiniFAT";
sector_list[fat_addrs[0]].name = "!FAT";
sector_list.fat_addrs = fat_addrs;
sector_list.ssz = ssz;

/* [MS-CFB] 2.6.1 Compound File Directory Entry */
var files = {}, Paths/*:any*/ = [], FileIndex = [], FullPaths = [], FullPathDir = {};
read_directory(dir_start, sector_list, sectors, Paths, nmfs, files, FileIndex);

build_full_paths(FileIndex, FullPathDir, FullPaths, Paths);

var root_name = Paths.shift();
Paths.root = root_name;

/* [MS-CFB] 2.6.4 (Unicode 3.0.1 case conversion) */
var find_path = make_find_path(FullPaths, Paths, FileIndex, files, root_name);

return {
	raw: {header: header, sectors: sectors},
	FileIndex: FileIndex,
	FullPaths: FullPaths,
	FullPathDir: FullPathDir,
	find: find_path
};
} // parse

/* [MS-CFB] 2.2 Compound File Header -- read up to major version */
function check_get_mver(blob) {
	// header signature 8
	blob.chk(HEADER_SIGNATURE, 'Header Signature: ');

	// clsid 16
	blob.chk(HEADER_CLSID, 'CLSID: ');

	// minor version 2
	var mver = blob.read_shift(2, 'u');

	return [blob.read_shift(2,'u'), mver];
}
function check_shifts(blob, mver) {
	var shift = 0x09;

	// Byte Order
	blob.chk('feff', 'Byte Order: ');

	// Sector Shift
	switch((shift = blob.read_shift(2))) {
		case 0x09: if(mver !== 3) throw 'MajorVersion/SectorShift Mismatch'; break;
		case 0x0c: if(mver !== 4) throw 'MajorVersion/SectorShift Mismatch'; break;
		default: throw 'Sector Shift: Expected 9 or 12 saw ' + shift;
	}

	// Mini Sector Shift
	blob.chk('0600', 'Mini Sector Shift: ');

	// Reserved
	blob.chk('000000000000', 'Reserved: ');
}

/** Break the file up into sectors */
function sectorify(file, ssz) {
	var nsectors = Math.ceil(file.length/ssz)-1;
	var sectors = new Array(nsectors);
	for(var i=1; i < nsectors; ++i) sectors[i-1] = file.slice(i*ssz,(i+1)*ssz);
	sectors[nsectors-1] = file.slice(nsectors*ssz);
	return sectors;
}

/* [MS-CFB] 2.6.4 Red-Black Tree */
function build_full_paths(FI, FPD, FP, Paths) {
	var i = 0, L = 0, R = 0, C = 0, j = 0, pl = Paths.length;
	var dad = new Array(pl), q = new Array(pl);

	for(; i < pl; ++i) { dad[i]=q[i]=i; FP[i]=Paths[i]; }

	for(; j < q.length; ++j) {
		i = q[j];
		L = FI[i].L; R = FI[i].R; C = FI[i].C;
		if(dad[i] === i) {
			if(L !== -1 /*NOSTREAM*/ && dad[L] !== L) dad[i] = dad[L];
			if(R !== -1 && dad[R] !== R) dad[i] = dad[R];
		}
		if(C !== -1 /*NOSTREAM*/) dad[C] = i;
		if(L !== -1) { dad[L] = dad[i]; q.push(L); }
		if(R !== -1) { dad[R] = dad[i]; q.push(R); }
	}
	for(i=1; i !== pl; ++i) if(dad[i] === i) {
		if(R !== -1 /*NOSTREAM*/ && dad[R] !== R) dad[i] = dad[R];
		else if(L !== -1 && dad[L] !== L) dad[i] = dad[L];
	}

	for(i=1; i < pl; ++i) {
		if(FI[i].type === 0 /* unknown */) continue;
		j = dad[i];
		if(j === 0) FP[i] = FP[0] + "/" + FP[i];
		else while(j !== 0) {
			FP[i] = FP[j] + "/" + FP[i];
			j = dad[j];
		}
		dad[i] = 0;
	}

	FP[0] += "/";
	for(i=1; i < pl; ++i) {
		if(FI[i].type !== 2 /* stream */) FP[i] += "/";
		FPD[FP[i]] = FI[i];
	}
}

/* [MS-CFB] 2.6.4 */
function make_find_path(FullPaths, Paths, FileIndex, files, root_name) {
	var UCFullPaths = new Array(FullPaths.length);
	var UCPaths = new Array(Paths.length), i;
	for(i = 0; i < FullPaths.length; ++i) UCFullPaths[i] = FullPaths[i].toUpperCase().replace(chr0,'').replace(chr1,'!');
	for(i = 0; i < Paths.length; ++i) UCPaths[i] = Paths[i].toUpperCase().replace(chr0,'').replace(chr1,'!');
	return function find_path(path/*:string*/) {
		var k;
		if(path.charCodeAt(0) === 47 /* "/" */) { k=true; path = root_name + path; }
		else k = path.indexOf("/") !== -1;
		var UCPath = path.toUpperCase().replace(chr0,'').replace(chr1,'!');
		var w = k === true ? UCFullPaths.indexOf(UCPath) : UCPaths.indexOf(UCPath);
		if(w === -1) return null;
		return k === true ? FileIndex[w] : files[Paths[w]];
	};
}

/** Chase down the rest of the DIFAT chain to build a comprehensive list
    DIFAT chains by storing the next sector number as the last 32 bytes */
function sleuth_fat(idx, cnt, sectors, ssz, fat_addrs) {
	var q;
	if(idx === ENDOFCHAIN) {
		if(cnt !== 0) throw "DIFAT chain shorter than expected";
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
function get_sector_list(sectors, start, fat_addrs, ssz, chkd) {
	var sl = sectors.length;
	var buf, buf_chain;
	if(!chkd) chkd = new Array(sl);
	var modulus = ssz - 1, j, jj;
	buf = [];
	buf_chain = [];
	for(j=start; j>=0;) {
		chkd[j] = true;
		buf[buf.length] = j;
		buf_chain.push(sectors[j]);
		var addr = fat_addrs[Math.floor(j*4/ssz)];
		jj = ((j*4) & modulus);
		if(ssz < 4 + jj) throw "FAT boundary crossed: " + j + " 4 "+ssz;
		if(!sectors[addr]) break;
		j = __readInt32LE(sectors[addr], jj);
	}
	return {nodes: buf, data:__toBuffer([buf_chain])};
}

/** Chase down the sector linked lists */
function make_sector_list(sectors, dir_start, fat_addrs, ssz/*:number*/)/*:any*/ {
	var sl = sectors.length, sector_list = new Array(sl);
	var chkd = new Array(sl), buf, buf_chain;
	var modulus = ssz - 1, i, j, k, jj;
	for(i=0; i < sl; ++i) {
		buf = [];
		k = (i + dir_start); if(k >= sl) k-=sl;
		if(chkd[k] === true) continue;
		buf_chain = [];
		for(j=k; j>=0;) {
			chkd[j] = true;
			buf[buf.length] = j;
			buf_chain.push(sectors[j]);
			var addr = fat_addrs[Math.floor(j*4/ssz)];
			jj = ((j*4) & modulus);
			if(ssz < 4 + jj) throw "FAT boundary crossed: " + j + " 4 "+ssz;
			if(!sectors[addr]) break;
			j = __readInt32LE(sectors[addr], jj);
		}
		sector_list[k] = {nodes: buf, data:__toBuffer([buf_chain])};
	}
	return sector_list;
}

/* [MS-CFB] 2.6.1 Compound File Directory Entry */
function read_directory(dir_start, sector_list, sectors, Paths, nmfs, files, FileIndex) {
	var blob;
	var minifat_store = 0, pl = (Paths.length?2:0);
	var sector = sector_list[dir_start].data;
	var i = 0, namelen = 0, name, o, ctime, mtime;
	for(; i < sector.length; i+= 128) {
		blob = sector.slice(i, i+128);
		prep_blob(blob, 64);
		namelen = blob.read_shift(2);
		if(namelen === 0) continue;
		name = __utf16le(blob,0,namelen-pl);
		Paths.push(name);
		o = ({
			name:  name,
			type:  blob.read_shift(1),
			color: blob.read_shift(1),
			L:     blob.read_shift(4, 'i'),
			R:     blob.read_shift(4, 'i'),
			C:     blob.read_shift(4, 'i'),
			clsid: blob.read_shift(16),
			state: blob.read_shift(4, 'i')
		}/*:any*/);
		ctime = blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2);
		if(ctime !== 0) {
			o.ctime = ctime; o.ct = read_date(blob, blob.l-8);
		}
		mtime = blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2);
		if(mtime !== 0) {
			o.mtime = mtime; o.mt = read_date(blob, blob.l-8);
		}
		o.start = blob.read_shift(4, 'i');
		o.size = blob.read_shift(4, 'i');
		if(o.type === 5) { /* root */
			minifat_store = o.start;
			if(nmfs > 0 && minifat_store !== ENDOFCHAIN) sector_list[minifat_store].name = "!StreamData";
			/*minifat_size = o.size;*/
		} else if(o.size >= 4096 /* MSCSZ */) {
			o.storage = 'fat';
			if(sector_list[o.start] === undefined) sector_list[o.start] = get_sector_list(sectors, o.start, sector_list.fat_addrs, sector_list.ssz);
			sector_list[o.start].name = o.name;
			o.content = sector_list[o.start].data.slice(0,o.size);
			prep_blob(o.content, 0);
		} else {
			o.storage = 'minifat';
			if(minifat_store !== ENDOFCHAIN && o.start !== ENDOFCHAIN) {
				o.content = sector_list[minifat_store].data.slice(o.start*MSSZ,o.start*MSSZ+o.size);
				prep_blob(o.content, 0);
			}
		}
		files[name] = o;
		FileIndex.push(o);
	}
}

function read_date(blob, offset) {
	return new Date(( ( (__readUInt32LE(blob,offset+4)/1e7)*Math.pow(2,32)+__readUInt32LE(blob,offset)/1e7 ) - 11644473600)*1000);
}

var fs;
function readFileSync(filename/*:string*/, options/*:any*/) {
	if(fs === undefined) fs = require('fs');
	return parse(fs.readFileSync(filename), options);
}

function readSync(blob/*:any*/, options/*:any*/) {
	switch(options !== undefined && options.type !== undefined ? options.type : "base64") {
		case "file": return readFileSync(blob, options);
		case "base64": return parse(s2a(Base64.decode(blob)), options);
		case "binary": return parse(s2a(blob), options);
	}
	return parse(blob);
}

/** CFB Constants */
var MSSZ = 64; /* Mini Sector Size = 1<<6 */
//var MSCSZ = 4096; /* Mini Stream Cutoff Size */
/* 2.1 Compound File Sector Numbers and Types */
var ENDOFCHAIN = -2;
/* 2.2 Compound File Header */
var HEADER_SIGNATURE = 'd0cf11e0a1b11ae1';
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

exports.read = readSync;
exports.parse = parse;
exports.utils = {
	ReadShift: ReadShift,
	CheckField: CheckField,
	prep_blob: prep_blob,
	bconcat: bconcat,
	consts: consts
};

return exports;
})();

if(typeof require !== 'undefined' && typeof module !== 'undefined' && typeof DO_NOT_EXPORT_CFB === 'undefined') { module.exports = CFB; }
