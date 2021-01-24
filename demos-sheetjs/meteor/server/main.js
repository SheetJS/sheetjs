/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
import { Meteor } from 'meteor/meteor';
import { check } from 'meteor/check';
import XLSX from 'xlsx';

Meteor.methods({
  /* read the data and return the workbook object to the frontend */
  uploadS: (bstr, name) => {
    check(bstr, String);
    check(name, String);
    return XLSX.read(bstr, { type: 'binary' });
  },
  uploadU: (ab, name) => {
    check(ab, Uint8Array);
    check(name, String);
    return XLSX.read(ab, { type: 'array' });
  },
  download: (html) => {
    check(html, String);
    let wb;
    if (html.length > 3) {
      /* parse workbook if html is available */
      wb = XLSX.read(html, { type: 'binary' });
    } else {
      /* generate a workbook object otherwise */
      const data = [['a', 'b', 'c'], [1, 2, 3]];
      const ws = XLSX.utils.aoa_to_sheet(data);
      wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'SheetJS');
    }
    return wb;
  },
});

Meteor.startup(() => { });
