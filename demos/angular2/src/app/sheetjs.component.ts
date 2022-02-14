/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
import { Component } from '@angular/core';

import { WorkBook, WorkSheet, WritingOptions, read, writeFileXLSX as writeFile, utils, version, set_cptable } from 'xlsx';
//import * as cpexcel from 'xlsx/dist/cpexcel.full.mjs';
//set_cptable(cpexcel);

type AOA = any[][];

@Component({
	selector: 'sheetjs',
	template: `
	<pre><b>Version: {{ver}}</b></pre>
	<input type="file" (change)="onFileChange($event)" multiple="false" />
	<table class="sjs-table">
		<tr *ngFor="let row of data">
			<td *ngFor="let val of row">
				{{val}}
			</td>
		</tr>
	</table>
	<button (click)="export()">Export!</button>
	`
})

export class SheetJSComponent {
	data: AOA = [ [1, 2], [3, 4] ];
	wopts: WritingOptions = { bookType: 'xlsx', type: 'array' };
	fileName: string = 'SheetJS.xlsx';
	ver: string = version;

	onFileChange(evt: any) {
		/* wire up file reader */
		const target: DataTransfer = <DataTransfer>(evt.target);
		if (target.files.length !== 1) throw new Error('Cannot use multiple files');
		const reader: FileReader = new FileReader();
		reader.onload = (e: any) => {
			/* read workbook */
			const ab: ArrayBuffer = e.target.result;
			const wb: WorkBook = read(ab);

			/* grab first sheet */
			const wsname: string = wb.SheetNames[0];
			const ws: WorkSheet = wb.Sheets[wsname];

			/* save data */
			this.data = <AOA>(utils.sheet_to_json(ws, {header: 1}));
		};
		reader.readAsArrayBuffer(target.files[0]);
	}

	export(): void {
		/* generate worksheet */
		const ws: WorkSheet = utils.aoa_to_sheet(this.data);

		/* generate workbook and add the worksheet */
		const wb: WorkBook = utils.book_new();
		utils.book_append_sheet(wb, ws, 'Sheet1');

		/* save to file */
		writeFile(wb, this.fileName);
	}
}
