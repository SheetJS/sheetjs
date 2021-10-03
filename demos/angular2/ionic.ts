/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
import { Component } from '@angular/core';
import { File } from '@ionic-native/file/ngx';
import * as XLSX from 'xlsx';


type AOA = any[][];

@Component({
  selector: 'app-home',
  //templateUrl: 'home.page.html',
  styleUrls: ['home.page.scss'],
  template: `
<ion-header>
  <ion-toolbar>
    <ion-title>SheetJS Ionic Demo</ion-title>
  </ion-toolbar>
</ion-header>

<ion-content [fullscreen]="true">
  <ion-header collapse="condense">
    <ion-toolbar>
      <ion-title>SheetJS Demo</ion-title>
    </ion-toolbar>
  </ion-header>

  <ion-grid>
    <ion-row *ngFor="let row of data">
      <ion-col *ngFor="let val of row">
        {{val}}
      </ion-col>
    </ion-row>
  </ion-grid>
</ion-content>

<ion-footer padding>
  <input type="file" (change)="onFileChange($event)" multiple="false" />
  <button ion-button color="secondary" (click)="import()">Import Data</button>
  <button ion-button color="secondary" (click)="export()">Export Data</button>
</ion-footer>
`
})

export class HomePage {
  data: any[][] = [[1,2,3],[4,5,6]];
  constructor(public file: File) {}

  read(ab: ArrayBuffer) {
    /* read workbook */
    const wb: XLSX.WorkBook = XLSX.read(new Uint8Array(ab), {type: 'array'});

    /* grab first sheet */
    const wsname: string = wb.SheetNames[0];
    const ws: XLSX.WorkSheet = wb.Sheets[wsname];

    /* save data */
    this.data = (XLSX.utils.sheet_to_json(ws, {header: 1}) as AOA);
  };

  write(): XLSX.WorkBook {
    /* generate worksheet */
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);

    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'SheetJS');

    return wb;
  };

  /* File Input element for browser */
  onFileChange(evt: any) {
    /* wire up file reader */
    const target: DataTransfer = (evt.target as DataTransfer);
    if (target.files.length !== 1) { throw new Error('Cannot use multiple files'); }
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      const ab: ArrayBuffer = e.target.result;
      this.read(ab);
    };
    reader.readAsArrayBuffer(target.files[0]);
  };

  /* Import button for mobile */
  async import() {
    try {
      const target: string = this.file.documentsDirectory || this.file.externalDataDirectory || this.file.dataDirectory || '';
      const dentry = await this.file.resolveDirectoryUrl(target);
      const url: string = dentry.nativeURL || '';
      alert(`Attempting to read SheetJSIonic.xlsx from ${url}`);
      const ab: ArrayBuffer = await this.file.readAsArrayBuffer(url, 'SheetJSIonic.xlsx');
      this.read(ab);
    } catch(e) {
      const m: string = e.message;
      alert(m.match(/It was determined/) ? 'Use File Input control' : `Error: ${m}`);
    }
  };

  /* Export button */
  async export() {
    const wb: XLSX.WorkBook = this.write();
    const filename = 'SheetJSIonic.xlsx';
    try {
      /* generate Blob */
      const wbout: ArrayBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

      /* find appropriate path for mobile */
      const target: string = this.file.documentsDirectory || this.file.externalDataDirectory || this.file.dataDirectory || '';
      const dentry = await this.file.resolveDirectoryUrl(target);
      const url: string = dentry.nativeURL || '';

      /* attempt to save blob to file */
      await this.file.writeFile(url, filename, wbout, {replace: true});
      alert(`Wrote to SheetJSIonic.xlsx in ${url}`);
    } catch(e) {
      if(e.message.match(/It was determined/)) {
        /* in the browser, use writeFile */
        XLSX.writeFile(wb, filename);
      } else {
        alert(`Error: ${e.message}`);
      }
    }
  };
}

