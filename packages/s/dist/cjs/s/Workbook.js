"use strict";
/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
/* vim: set ts=2: */
Object.defineProperty(exports, "__esModule", { value: true });
/// <reference path="../xlsx.d.ts"/>
var XLSXWrapper_1 = require("./XLSXWrapper");
var WorksheetCollection_1 = require("./worksheet/WorksheetCollection");
var WorkbookDefinedNameCollection_1 = require("./names/WorkbookDefinedNameCollection");
var Workbook = /** @class */ (function () {
    function Workbook(wb) {
        this._wb = wb || XLSXWrapper_1.get_XLSX().utils.book_new();
        this._ws = new WorksheetCollection_1.WorksheetCollection(this._wb);
        this._names = new WorkbookDefinedNameCollection_1.WorkbookDefinedNameCollection(this._wb);
    }
    ;
    Object.defineProperty(Workbook.prototype, "wb", {
        get: function () { return this._wb; },
        enumerable: true,
        configurable: true
    });
    ;
    Object.defineProperty(Workbook.prototype, "names", {
        get: function () { return this._names; },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Workbook.prototype, "worksheets", {
        get: function () { return this._ws; },
        enumerable: true,
        configurable: true
    });
    ;
    return Workbook;
}());
exports.Workbook = Workbook;
;
//# sourceMappingURL=Workbook.js.map