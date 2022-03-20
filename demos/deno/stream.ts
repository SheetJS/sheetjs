import {utils, stream, set_cptable} from '../../xlsx.mjs';
import * as cpexcel from '../../dist/cpexcel.full.mjs';
set_cptable(cpexcel);

function NodeReadableCB(cb:(d:any)=>void) {
	var rd = {
		__done: false,
		_read: function() {},
		push: function(d: any) { if(!this.__done) cb(d); if(d == null) this.__done = true; },
		resume: function pump() {for(var i = 0; i < 10000 && !this.__done; ++i) rd._read(); if(!rd.__done) setTimeout(pump, 0); }
	};
	return rd;
}

function NodeReadable(rd: any) { return function() { return rd; }; }

const L = 1_000_000;
const W = 30;

console.time("prep");
const ws = utils.aoa_to_sheet([Array.from({length: W}, (_, C) => utils.encode_col(C))], {dense: true});
for(let l = 1; l < L; ++l) utils.sheet_add_aoa(ws, [Array.from({length: W}, (_,j) => j == 0 ? String(l) : l+j)], {origin: -1});
console.timeEnd("prep");

console.time("stream");
var cnt = 0;
const rt = NodeReadableCB((d: any) => {
	++cnt; if((cnt%10000) == 0) console.log(cnt); if(d == null) console.timeEnd("stream");
});
stream.set_readable(NodeReadable(rt));
const rd = stream.to_csv(ws);
rd.resume();
