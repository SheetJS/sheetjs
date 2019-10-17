/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
/* vim: set ts=2: */
/// <reference path="../xlsx.d.ts"/>
import { get_XLSX } from "./XLSXWrapper.mjs";
import { WorksheetCollection } from "./worksheet/WorksheetCollection.mjs";
import { WorkbookDefinedNameCollection } from "./names/WorkbookDefinedNameCollection.mjs";
export class Workbook {
    constructor(wb) {
        this._wb = wb || get_XLSX().utils.book_new();
        this._ws = new WorksheetCollection(this._wb);
        this._names = new WorkbookDefinedNameCollection(this._wb);
    }
    ;
    get wb() { return this._wb; }
    ;
    get names() { return this._names; }
    get worksheets() { return this._ws; }
    ;
}
;
//# sourceMappingURL=Workbook.js.map
