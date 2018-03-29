/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
const XLSX = require('xlsx');
const assert = require('assert');
const Firebase = require('firebase-admin');

const config = {
	credential: {
		getAccessToken: () => ({
			expires_in: 0,
			access_token: '',
		}),
	},
	databaseURL: 'ws://localhost:5555'
};

/* make new workbook object from CSV */
const wb = XLSX.read('a,b,c\n1,2,3', {type:"binary", raw:true});

let P = Promise.resolve("sheetjs");

/* Connect to Firebase server and initialize collection */
P = P.then(async () => {
	Firebase.initializeApp(config);
	const database = Firebase.database();
	await database.ref('foo').set(null);
	return [database];
});

/* Insert entire workbook object into `foo` ref */
P = P.then(async ([database]) => {
	await database.ref('foo').set(wb);
	return [database];
});

/* Change cell A1 of Sheet1 to "J" and change A2 to 5 */
P = P.then(async ([database]) => {
	database.ref('foo').update({
		"Sheets/Sheet1/A1": {"t": "s", "v": "J"},
		"Sheets/Sheet1/A2": {"t": "n", "v": 5},
	});
	return [database];
});

/* Write to file */
P = P.then(async ([database]) => {
	const val = await database.ref('foo').once('value');
	const wb = await val.val();
	XLSX.writeFile(wb, "firebase.xlsx");
	const ws = XLSX.readFile("firebase.xlsx").Sheets.Sheet1;
	const csv = XLSX.utils.sheet_to_csv(ws);
	assert.equal(csv, "J,b,c\n5,2,3\n");
	console.log(csv);
	return [database];
});

/* Close connection */
P = P.then(async ([database]) => { database.app.delete(); });
