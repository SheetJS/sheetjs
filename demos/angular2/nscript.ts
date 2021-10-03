/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */

import { Component } from '@angular/core';
import { encoding } from '@nativescript/core/text';
import { File, Folder, knownFolders, path } from '@nativescript/core/file-system';
import { Dialogs } from '@nativescript/core';
import { Page, GridLayout, WebView, DockLayout, Button } from '@nativescript/core';

import * as XLSX from './xlsx.full.min';

@Component({
  selector: 'ns-app',
  template: `
  <Page>
    <GridLayout rows="auto, *, auto">

      <!-- data converted to HTML and rendered in web view -->
      <WebView row="1" src="{{html}}"></WebView>

      <DockLayout row="2" dock="bottom" stretchLastChild="false">
        <Button text="Import File" (tap)="import()" style="padding: 10px"></Button>
        <Button text="Export File" (tap)="export()" style="padding: 10px"></Button>
      </DockLayout>
    </GridLayout>
  </Page>
  `
})

export class AppComponent {
  html: string = "";
  constructor() {
    const ws = XLSX.utils.aoa_to_sheet([[1,2],[3,4]]);
    this.html = XLSX.utils.sheet_to_html(ws);
  };

  /* Import button */
  async import() {
    const filename: string = "SheetJSNS.csv";

    /* find appropriate path */
    const target: Folder = knownFolders.documents() || knownFolders.ios.sharedPublic();
    const url: string = path.normalize(target.path + "///" + filename);
    const file: File = File.fromPath(url);

    try {
      /* get binary string */
      const bstr: string = await file.readText(encoding.ISO_8859_1);

      /* read workbook */
      const wb = XLSX.read(bstr, { type: "binary" });

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];

      /* update table */
      this.html = XLSX.utils.sheet_to_html(ws);
      Dialogs.alert(`Attempting to read to ${filename} in ${url}`);
    } catch(e) {
      Dialogs.alert(e.message);
    }
  };

  /* Export button */
  async export() {
    const wb = XLSX.read(this.html, { type: "string" });
    const filename: string = "SheetJSNS.csv";

    /* generate binary string */
    const wbout: string = XLSX.write(wb, { bookType: 'csv', type: 'binary' });

    /* find appropriate path */
    const target: Folder = knownFolders.documents() || knownFolders.ios.sharedPublic();
    const url: string = path.normalize(target.path + "///" + filename);
    const file: File = File.fromPath(url);

    /* attempt to save binary string to file */
    await file.writeText(wbout, encoding.ISO_8859_1);
    Dialogs.alert(`Wrote to ${filename} in ${url}`);
  };
}
