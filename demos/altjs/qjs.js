#!/usr/bin/env qjs
/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
/* load XLSX */
std.global.global = std.global;
std.loadScript("xlsx.full.min.js");

/* read contents of file */
var rh = std.open("sheetjs.xlsx", "rb");
rh.seek(0, std.SEEK_END);
var sz = rh.tell();
var ab = new ArrayBuffer(sz);
rh.seek();
rh.read(ab, 0, sz);
rh.close();

/* parse file */
var wb = XLSX.read(ab, {type: 'array'});

/* write array */
var out = XLSX.write(wb, {type: 'array'});

/* write contents to file */
var wh = std.open("sheetjs.qjs.xlsx", "wb");
wh.write(out, 0, out.byteLength);
wh.close();
