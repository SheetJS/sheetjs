/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
import { Meteor } from 'meteor/meteor';
import { check } from 'meteor/check';
import XLSX from 'xlsx';

Meteor.methods({
  upload: (bstr, name) => {
    /* read the data and return the workbook object to the frontend */
    check(bstr, String);
    check(name, String);
    return XLSX.read(bstr, { type: 'binary' });
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
