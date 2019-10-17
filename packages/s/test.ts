/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
/* vim: set ts=2: */

/// <reference path="src/xlsx.d.ts"/>

import { Workbook, set_XLSX, get_XLSX } from "./";
import * as assert from 'assert';
const XLSXLib: typeof XLSX = require("../../");
set_XLSX(XLSXLib);

import 'mocha';

describe('Defined Names', () => {
  let wb = new Workbook();

  it('should add names to blank workbook', () => {
    let cnt = wb.names.count;
    assert.equal(cnt, 0);
    assert.throws(() => { const newname = wb.names.getName("wtf"); });
    wb.names.add("wtf", "Sheet1!A1:A3", "dafuq");
    assert.doesNotThrow(() => { const newname = wb.names.getName("wtf"); });
  });
});
