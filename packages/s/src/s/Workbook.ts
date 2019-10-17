/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
/* vim: set ts=2: */

/// <reference path="../xlsx.d.ts"/>

import { get_XLSX } from "./XLSXWrapper";
import { WorksheetCollection } from "./worksheet/WorksheetCollection";
import { DefinedNameCollection } from "./names/DefinedNameCollection";
import { WorkbookDefinedNameCollection } from "./names/WorkbookDefinedNameCollection";

export class Workbook {
  private readonly _wb: XLSX.WorkBook;
  private readonly _ws: WorksheetCollection;
  private readonly _names: WorkbookDefinedNameCollection;

  constructor(wb?: XLSX.WorkBook) {
    this._wb = wb || get_XLSX().utils.book_new();
    this._ws = new WorksheetCollection(this._wb);
    this._names = new WorkbookDefinedNameCollection(this._wb);
  };

  get wb(): XLSX.WorkBook { return this._wb; };

  get names(): DefinedNameCollection { return this._names; }

  get worksheets(): WorksheetCollection { return this._ws; };

};