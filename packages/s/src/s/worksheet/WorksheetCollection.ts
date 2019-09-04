/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
/* vim: set ts=2: */

/// <reference path="../../xlsx.d.ts"/>

export class WorksheetCollection {
  private readonly _wb: XLSX.WorkBook;

  constructor(wb: XLSX.WorkBook) {
    this._wb = wb;
  };
};
