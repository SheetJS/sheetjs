/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */

import { Component } from "@angular/core";
import * as dockModule from "tns-core-modules/ui/layouts/dock-layout";
import * as buttonModule from "tns-core-modules/ui/button";
import * as textModule from "tns-core-modules/text";
import * as dialogs from "ui/dialogs";
import * as fs from "tns-core-modules/file-system";

/* NativeScript does not support import syntax for npm modules */
const XLSX = require("./xlsx.full.min.js");

@Component({
  selector: "my-app",
  template: `
    <GridLayout rows="auto, *, auto">
      <ActionBar row="0" title="SheetJS NativeScript Demo" class="action-bar"></ActionBar>

      <!-- data converted to HTML and rendered in web view -->
      <WebView row="1" src="{{html}}"></WebView>

      <DockLayout row="2" dock="bottom" stretchLastChild="false">
        <Button text="Import File" (tap)="import()" style="padding: 10px"></Button>
        <Button text="Export File" (tap)="export()" style="padding: 10px"></Button>
      </DockLayout>
    </GridLayout>
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
    const filename: string = "SheetJSNS.xlsx";

    /* find appropriate path */
    const target: fs.Folder = fs.knownFolders.documents() || fs.knownFolders.ios.sharedPublic();
    const url: string = fs.path.normalize(target.path + "///" + filename);
    const file: fs.File = fs.File.fromPath(url);

    try {
      /* get binary string */
      const bstr: string = await file.readText(textModule.encoding.ISO_8859_1);

      /* read workbook */
      const wb = XLSX.read(bstr, { type: "binary" });

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];

      /* update table */
      this.html = XLSX.utils.sheet_to_html(ws);
      dialogs.alert(`Attempting to read to SheetJSNS.xlsx in ${url}`);
    } catch(e) {
      dialogs.alert(e.message);
    }
  };

  /* Export button */
  async export() {
    const wb = XLSX.read(this.html, { type: "string" });
    const filename: string = "SheetJSNS.xlsx";

    /* generate binary string */
    const wbout: string = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

    /* find appropriate path */
    const target: fs.Folder = fs.knownFolders.documents() || fs.knownFolders.ios.sharedPublic();
    const url: string = fs.path.normalize(target.path + "///" + filename);
    const file: fs.File = fs.File.fromPath(url);

    /* attempt to save binary string to file */
    await file.writeText(wbout, textModule.encoding.ISO_8859_1);
    dialogs.alert(`Wrote to SheetJSNS.xlsx in ${url}`);
  };
}
