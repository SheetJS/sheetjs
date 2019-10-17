/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
/// <reference path="../../../src/xlsx.d.ts" />
import { WorksheetCollection } from "./worksheet/WorksheetCollection";
import { DefinedNameCollection } from "./names/DefinedNameCollection";
export declare class Workbook {
    private readonly _wb;
    private readonly _ws;
    private readonly _names;
    constructor(wb?: XLSX.WorkBook);
    readonly wb: XLSX.WorkBook;
    readonly names: DefinedNameCollection;
    readonly worksheets: WorksheetCollection;
}
