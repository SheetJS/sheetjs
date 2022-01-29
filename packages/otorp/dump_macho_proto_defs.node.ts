#!/usr/bin/env ts-node
/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */

import { readFileSync, writeFileSync } from 'fs';
import { resolve } from 'path';
import { otorp } from './';

if(!process.argv[2] || process.argv[2] == "-h" || process.argv[2] == "--help") {
	[
		"usage: otorp <path/to/bin> [output/folder]",
		"  if no output folder specified, log all discovered defs",
		"  if output folder specified, attempt to write defs in the folder"
	].map(x => console.error(x));
	process.exit(1);
}
var buf = readFileSync(process.argv[2]);

var otorps = otorp(buf);

otorps.forEach(({name, proto}) => {
	if(!process.argv[3]) {
		console.log(proto);
	} else {
		var pth = resolve(process.argv[3] || "./", name.replace(/[/]/g, "$"));
		console.error(`writing ${name} to ${pth}`);
		writeFileSync(pth, proto);
	}
});
