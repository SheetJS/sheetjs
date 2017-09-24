/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
import XLSX from 'xlsx';
/* note: saveAs is made available via the smart package */
/* global saveAs */

import { Meteor } from 'meteor/meteor';
import { Template } from 'meteor/templating';

import './main.html';

Template.sheetjs.events({
  'change input' (event) {
    /* "Browser file upload form element" from SheetJS README */
    const file = event.currentTarget.files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
      const data = e.target.result;
      const name = file.name;
      /* Meteor magic */
      Meteor.call('upload', data, name, function(err, wb) {
        if (err) throw err;
        /* load the first worksheet */
        const ws = wb.Sheets[wb.SheetNames[0]];
        /* generate HTML table and enable export */
        const html = XLSX.utils.sheet_to_html(ws, { editable: true });
        document.getElementById('out').innerHTML = html;
        document.getElementById('dnload').disabled = false;
      });
    };
    reader.readAsBinaryString(file);
  },
  'click button' () {
    const html = document.getElementById('out').innerHTML;
    Meteor.call('download', html, function(err, wb) {
      if (err) throw err;
      /* "Browser download file" from SheetJS README */
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
      saveAs(new Blob([s2ab(wbout)], { type: 'application/octet-stream' }), 'sheetjs.xlsx');
    });
  },
});

/* eslint no-bitwise:0, no-plusplus:0 */
function s2ab(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}
