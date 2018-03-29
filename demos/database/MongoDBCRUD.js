/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
/* global Promise */
const XLSX = require('xlsx');
const SheetJSMongo = require("./SheetJSMongo");
const MongoClient = require('mongodb').MongoClient;

const url = 'mongodb://localhost:27017/sheetjs';
const db_name = 'sheetjs';

let P = Promise.resolve("sheetjs");

/* Connect to mongodb server */
P = P.then(async () => {
	const client = await MongoClient.connect(url);
	return [client];
});

/* Sample data table */
P = P.then(async ([client]) => {
	const db = client.db(db_name);

	try { await db.collection('pres').drop(); } catch(e) {}
	const pres = db.collection('pres');
	await pres.insertMany([
		{ name: "Barack Obama", idx: 44 },
		{ name: "Donald Trump", idx: 45 }
	], {ordered: true});

	try { await db.collection('fmts').drop(); } catch(e) {}
	const fmts = db.collection('fmts');
	await fmts.insertMany([
		{ ext: 'XLSB', ctr: 'ZIP', multi: 1 },
		{ ext: 'XLS',  ctr: 'CFB', multi: 1 },
		{ ext: 'XLML',             multi: 1 },
		{ ext: 'CSV',  ctr: 'ZIP', multi: 0 },
	], {ordered: true});

	return [client, pres, fmts];
});

/* Export database to XLSX */
P = P.then(async ([client, pres, fmts]) => {
	const wb = XLSX.utils.book_new();
	await SheetJSMongo.book_append_mongo(wb, pres, "pres");
	await SheetJSMongo.book_append_mongo(wb, fmts, "fmts");
	XLSX.writeFile(wb, "mongocrud.xlsx");
	return [client, pres, fmts];
});

/* Read the new file and dump all of the data */
P = P.then(() => {
	const wb = XLSX.readFile('mongocrud.xlsx');
	wb.SheetNames.forEach((n,i) => {
		console.log(`Sheet #${i+1}: ${n}`);
		const ws = wb.Sheets[n];
		console.log(XLSX.utils.sheet_to_csv(ws));
	});
});

/* Close connection */
P.then(async ([client]) => { client.close(); });
