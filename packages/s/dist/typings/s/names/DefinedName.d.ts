/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
/// <reference path="../../../../src/xlsx.d.ts" />
export declare class DefinedName {
    readonly _name: XLSX.DefinedName;
    constructor(name: XLSX.DefinedName);
    readonly raw: XLSX.DefinedName;
}
