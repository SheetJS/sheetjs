import { Template } from 'meteor/templating';
import { ReactiveVar } from 'meteor/reactive-var';

import './main.html';

const XLSX = require('xlsx');

Template.read.events({
	'change input' (evt, instance) {
		/* "Browser file upload form element" from SheetJS README */
		const file = evt.currentTarget.files[0];
		const reader = new FileReader();
		reader.onload = function(e) {
			const data = e.target.result;
			const name = file.name;
			/* Meteor magic */
			Meteor.call('upload', data, name, function(err, wb) {
				if(err) console.error(err);
				else {
					/* do something here -- this just dumps an array of arrays to console */
					console.log(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1}));
					document.getElementById('out').innerHTML = (XLSX.utils.sheet_to_html(wb.Sheets[wb.SheetNames[0]]));
				}
			});
		};
		reader.readAsBinaryString(file);
	},
});

Template.write.events({
	'click button' (evt, instance) {
		Meteor.call('download', function(err, wb) {
			if(err) console.error(err);
			else {
				console.log(wb);
				/* "Browser download file" from SheetJS README */
				var wopts = { bookType:'xlsx', bookSST:false, type:'binary' };
				var wbout = XLSX.write(wb, wopts);
				saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), "meteor.xlsx");
			}
		});
	},
});

function s2ab(s) {
	var buf = new ArrayBuffer(s.length);
	var view = new Uint8Array(buf);
	for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
	return buf;
}
