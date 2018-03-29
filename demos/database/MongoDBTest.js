/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
/* global Promise */
const XLSX = require('xlsx');
const assert = require('assert');
const MongoClient = require('mongodb').MongoClient;

const url = 'mongodb://localhost:27017/sheetjs';
const db_name = 'sheetjs';

/* make new workbook object from CSV */
const wb = XLSX.read('a,b,c\n1,2,3', {type:"binary", raw:true});

let P = Promise.resolve("sheetjs");

/* Connect to mongodb server and initialize collection */
P = P.then(async () => {
	const client = await MongoClient.connect(url);
	const db = client.db(db_name);
	try { await db.collection('wb').drop(); } catch(e) {}
	const coll = db.collection('wb');
	return [client, coll];
});

/* Insert entire workbook object as a document */
P = P.then(async ([client, coll]) => {
	const res = await coll.insertOne(wb);
	assert.equal(res.insertedCount, 1);
	return [client, coll];
});

/* Change cell A1 of Sheet1 to "J" and change A2 to 5 */
P = P.then(async ([client, coll]) => {
	const res = await coll.updateOne({}, { $set: {
		"Sheets.Sheet1.A1": {"t": "s", "v": "J"},
		"Sheets.Sheet1.A2": {"t": "n", "v": 5},
	}});
	assert.equal(res.matchedCount, 1);
	assert.equal(res.modifiedCount, 1);
	return [client, coll];
});

/* Write to file */
P = P.then(async ([client, coll]) => {
	const res = await coll.find({}).toArray();
	const wb = res[0];
	XLSX.writeFile(wb, "mongo.xlsx");
	const ws = XLSX.readFile("mongo.xlsx").Sheets.Sheet1;
	console.log(XLSX.utils.sheet_to_csv(ws));
	return [client, coll];
});

/* Close connection */
P.then(async ([client]) => { client.close(); });
