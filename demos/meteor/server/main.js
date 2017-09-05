/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
import { Meteor } from 'meteor/meteor';

const XLSX = require('xlsx');

Meteor.methods({
	upload: (bstr, name) => {
		/* read the data and return the workbook object to the frontend */
		return XLSX.read(bstr, {type:'binary'});
	},
	download: () => {
		/* generate a workbook object and return to the frontend */
		const data = [
			["a", "b", "c"],
			[ 1 ,  2 ,  3 ]
		];
		const ws = XLSX.utils.aoa_to_sheet(data);
		const wb = {SheetNames: ["Sheet1"], Sheets:{Sheet1:ws }};
		return wb;
	}
});

Meteor.startup(() => { });
