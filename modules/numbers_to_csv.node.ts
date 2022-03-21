#!/usr/bin/env ts-node
/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
import { read } from 'cfb';
import { utils } from 'xlsx';
import { parse_numbers_iwa } from './src/numbers';

var f = process.argv[2];
var cfb = read(f, {type: "file"});
var wb = parse_numbers_iwa(cfb);
var sn = process.argv[3];
if(typeof sn == "undefined") {
	wb.SheetNames.forEach(sn => console.log(utils.sheet_to_csv(wb.Sheets[sn])));
} else {
	if(sn && !isNaN(+sn)) sn = wb.SheetNames[+sn];
	if(wb.SheetNames.indexOf(sn) == -1) sn = wb.SheetNames[0];
	console.log(utils.sheet_to_csv(wb.Sheets[sn]));
}
