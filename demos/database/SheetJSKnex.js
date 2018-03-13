/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
var XLSX = require("xlsx");

async function book_append_knex(wb, knex, tbl) {
	const aoo = await knex.select("*").from(tbl);
	XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(aoo), tbl);
}

const TYPES = {
  b: "boolean",
  n: "float",
  t: "string",
  d: "dateTime"
};
async function ws_to_knex(ws, knex, n) {
  if(!ws || !ws['!ref']) return;
  var range = XLSX.utils.decode_range(ws['!ref']);
  if(!range || !range.s || !range.e || range.s > range.e) return;
  var R = range.s.r, C = range.s.c;

  var names = new Array(range.e.c-range.s.c+1);
  for(C = range.s.c; C<= range.e.c; ++C){
    var addr = XLSX.utils.encode_cell({c:C,r:R});
    names[C-range.s.c] = ws[addr] ? ws[addr].v : XLSX.utils.encode_col(C);
  }

  for(var i = 0; i < names.length; ++i) if(names.indexOf(names[i]) < i)
    for(var j = 0; j < names.length; ++j) {
      var _name = names[i] + "_" + (j+1);
      if(names.indexOf(_name) > -1) continue;
      names[i] = _name;
    }

  var types = new Array(range.e.c-range.s.c+1);
  for(C = range.s.c; C<= range.e.c; ++C) {
    var seen = {}, _type = "";
    for(R = range.s.r+1; R<= range.e.r; ++R)
      seen[(ws[XLSX.utils.encode_cell({c:C,r:R})]||{t:"z"}).t] = true;
    if(seen.s || seen.str) _type = TYPES.t;
    else if(seen.n + seen.b + seen.d + seen.e > 1) _type = TYPES.t;
    else switch(true) {
      case seen.b: _type = TYPES.b; break;
      case seen.n: _type = TYPES.n; break;
      case seen.e: _type = TYPES.t; break;
      case seen.d: _type = TYPES.d; break;
    }
    types[C-range.s.c] = _type || TYPES.t;
  }

  await knex.schema.dropTableIfExists(n);
	await knex.schema.createTable(n, (table) => { names.forEach((n, i) => { table[types[i] || "text"](n); }); });

  for(R = range.s.r+1; R<= range.e.r; ++R) {
    var row = {};
    for(C = range.s.c; C<= range.e.c; ++C) {
      var cell = ws[XLSX.utils.encode_cell({c:C,r:R})];
      if(!cell) continue;
      var key = names[C-range.s.c], val = cell.v;
      if(types[C-range.s.c] == TYPES.n) if(cell.t == 'b' || typeof val == 'boolean' ) val = +val;
      row[key] = val;
    }
    await knex.insert(row).into(n);;
  }
}

async function wb_to_knex(wb, knex) {
	for(var i = 0; i < wb.SheetNames.length; ++i) {
		var n = wb.SheetNames[i];
		var ws = wb.Sheets[n];
		await ws_to_knex(ws, knex, n);
	}
}

module.exports = {
	book_append_knex,
	wb_to_knex
};
