/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
var XLSX = require("xlsx");
	
async function book_append_mongo(wb, coll, name) {
	const aoo = await coll.find({}).toArray();
	aoo.forEach((x) => delete x._id);
	const ws = XLSX.utils.json_to_sheet(aoo);
	XLSX.utils.book_append_sheet(wb, ws, name);
	return ws;
}

module.exports = {
	book_append_mongo
};
