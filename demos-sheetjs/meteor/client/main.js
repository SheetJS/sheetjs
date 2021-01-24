/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
import XLSX from 'xlsx';

import { Meteor } from 'meteor/meteor';
import { Template } from 'meteor/templating';

import './main.html';

Template.sheetjs.events({
  'change input' (event) {
    /* "Browser file upload form element" from SheetJS README */
    const file = event.currentTarget.files[0];
    const reader = new FileReader();
    const rABS = !!reader.readAsBinaryString;
    reader.onload = function(e) {
      const data = e.target.result;
      const name = file.name;
      /* Meteor magic */
      Meteor.call(rABS ? 'uploadS' : 'uploadU', rABS ? data : new Uint8Array(data), name, function(err, wb) {
        if (err) throw err;
        /* load the first worksheet */
        const ws = wb.Sheets[wb.SheetNames[0]];
        /* generate HTML table and enable export */
        const html = XLSX.utils.sheet_to_html(ws, { editable: true });
        document.getElementById('out').innerHTML = html;
        document.getElementById('dnload').disabled = false;
      });
    };
    if(rABS) reader.readAsBinaryString(file); else reader.readAsArrayBuffer(file);
  },
  'click button' () {
    const html = document.getElementById('out').innerHTML;
    Meteor.call('download', html, function(err, wb) {
      if (err) throw err;
      /* "Browser download file" from SheetJS README */
      XLSX.writeFile(wb, 'sheetjs.xlsx');
    });
  },
});

