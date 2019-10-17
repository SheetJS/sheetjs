"use strict";
/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
/* vim: set ts=2: */
Object.defineProperty(exports, "__esModule", { value: true });
/// <reference path="../../xlsx.d.ts"/>
var DefinedName = /** @class */ (function () {
    function DefinedName(name) {
        this._name = name;
    }
    ;
    Object.defineProperty(DefinedName.prototype, "raw", {
        get: function () { return this._name; },
        enumerable: true,
        configurable: true
    });
    ;
    return DefinedName;
}());
exports.DefinedName = DefinedName;
;
//# sourceMappingURL=DefinedName.js.map