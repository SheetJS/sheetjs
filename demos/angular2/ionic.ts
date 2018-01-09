/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
import { Component } from '@angular/core';

import * as XLSX from 'xlsx';

import { File } from '@ionic-native/file';
import { saveAs } from 'file-saver';

type AOA = any[][];

@Component({
  selector: 'page-home',
  template: `
<ion-header><ion-navbar><ion-title>SheetJS Ionic Demo</ion-title></ion-navbar></ion-header>

<ion-content padding>
  <ion-grid>
    <ion-row *ngFor="let row of data">
      <ion-col *ngFor="let val of row">
        {{val}}
      </ion-col>
    </ion-row>
  </ion-grid>
</ion-content>

<ion-footer>
  <input type="file" (change)="onFileChange($event)" multiple="false" />
  <button ion-button color="secondary" (click)="import()">Import Data</button>
  <button ion-button color="secondary" (click)="export()">Export Data</button>
</ion-footer>
`
})

export class HomePage {
  data: any[][] = [[1,2,3],[4,5,6]];
  constructor(public file: File) {};

  read(bstr: string) {
    /* read workbook */
    const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});

    /* grab first sheet */
    const wsname: string = wb.SheetNames[0];
    const ws: XLSX.WorkSheet = wb.Sheets[wsname];

    /* save data */
    this.data = <AOA>(XLSX.utils.sheet_to_json(ws, {header: 1}));
  };

  write(): ArrayBuffer {
    /* generate worksheet */
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);

    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'SheetJS');

    /* save to ArrayBuffer */
    const wbout: ArrayBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    return wbout;
  };

  /* File Input element for browser */
  onFileChange(evt: any) {
    /* wire up file reader */
    const target: DataTransfer = <DataTransfer>(evt.target);
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      const bstr: string = e.target.result;
      this.read(bstr);
    };
    reader.readAsBinaryString(target.files[0]);
  };

  /* Import button for mobile */
  async import() {
    try {
      const target: string = this.file.documentsDirectory || this.file.externalDataDirectory || this.file.dataDirectory || '';
      const dentry = await this.file.resolveDirectoryUrl(target);
      const url: string = dentry.nativeURL || '';
      alert(`Attempting to read SheetJSIonic.xlsx from ${url}`)
      const bstr: string = await this.file.readAsBinaryString(url, "SheetJSIonic.xlsx");
      this.read(bstr);
    } catch(e) {
      const m: string = e.message;
      alert(m.match(/It was determined/) ? "Use File Input control" : `Error: ${m}`);
    }
  };

  /* Export button */
  async export() {
    const wbout: ArrayBuffer = this.write();
    const filename: string = "SheetJSIonic.xlsx";
    const blob: Blob = new Blob([wbout], {type: 'application/octet-stream'});
    try {
      const target: string = this.file.documentsDirectory || this.file.externalDataDirectory || this.file.dataDirectory || '';
      const dentry = await this.file.resolveDirectoryUrl(target);
      const url: string = dentry.nativeURL || '';
      await this.file.writeFile(url, filename, blob, {replace: true});
      alert(`Wrote to SheetJSIonic.xlsx in ${url}`);
    } catch(e) {
      if(e.message.match(/It was determined/)) saveAs(blob, filename);
      else alert(`Error: ${e.message}`);
    }
  };
}

