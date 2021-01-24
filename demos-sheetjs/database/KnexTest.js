/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
/* global Promise */
const XLSX = require('xlsx');
const assert = require('assert');
const SheetJSKnex = require("./SheetJSKnex");
const Knex = require('knex');

/* Connection to both databases are passed around */
let P = Promise.resolve([
	Knex({ client: 'sqlite3', connection: { filename: './knex1.db' } }),
	Knex({ client: 'sqlite3', connection: { filename: './knex2.db' } })
]);

/* Sample data table */
P = P.then(async (_) => {
	const [knex] = _;
	await knex.schema.dropTableIfExists('pres');
	await knex.schema.createTable('pres', (table) => {
		table.string('name');
		table.integer('idx');
	});
	await knex.insert([
		{ name: "Barack Obama", idx: 44 },
		{ name: "Donald Trump", idx: 45 }
	]).into('pres');

	await knex.schema.dropTableIfExists('fmts');
	await knex.schema.createTable('fmts', (table) => {
		table.string('ext');
		table.string('ctr');
		table.integer('multi');
	});
	await knex.insert([
		{ ext: 'XLSB', ctr: 'ZIP', multi: 1 },
		{ ext: 'XLS',  ctr: 'CFB', multi: 1 },
		{ ext: 'XLML', ctr: '',    multi: 1 },
		{ ext: 'CSV',  ctr: 'ZIP', multi: 0 },
	]).into('fmts');

	return _;
});

/* Export database to XLSX */
P = P.then(async (_) => {
	const [knex] = _;
	const wb = XLSX.utils.book_new();
	await SheetJSKnex.book_append_knex(wb, knex, "pres");
	await SheetJSKnex.book_append_knex(wb, knex, "fmts");
	XLSX.writeFile(wb, "knex.xlsx");
	return _;
});

/* Import XLSX to database */
P = P.then(async (_) => {
	const [, knex] = _;
	const wb = XLSX.readFile("knex.xlsx");
	await SheetJSKnex.wb_to_knex(wb, knex);
	return _;
});

/* Compare databases */
P = P.then(async (_) => {
	const [k1, k2] = _;
	const P1 = await k1.select("*").from('pres');
	const P2 = await k2.select("*").from('pres');
	const F1 = await k1.select("*").from('fmts');
	const F2 = await k2.select("*").from('fmts');
	assert.deepEqual(P1, P2);
	assert.deepEqual(F1, F2);
});

P.then(async () => { process.exit(); });
