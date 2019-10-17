"use strict";
/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
/* vim: set ts=2: */
Object.defineProperty(exports, "__esModule", { value: true });
var DefinedName_1 = require("./DefinedName");
//import { Range } from "../Range";
var WorkbookDefinedNameCollection = /** @class */ (function () {
    function WorkbookDefinedNameCollection(wb) {
        this._wb = wb;
    }
    ;
    Object.defineProperty(WorkbookDefinedNameCollection.prototype, "items", {
        /**
         * Get read-only array of global defined names
         */
        get: function () {
            if (!this._wb.Workbook)
                return [];
            if (!this._wb.Workbook.Names)
                return [];
            return this._wb.Workbook.Names.filter(function (name) { name.Sheet == null; }).map(function (name) { return new DefinedName_1.DefinedName(name); });
        },
        enumerable: true,
        configurable: true
    });
    ;
    /**
     * Get defined name object
     */
    WorkbookDefinedNameCollection.prototype.getName = function (name) {
        if (this._wb.Workbook && this._wb.Workbook.Names) {
            var names = this._wb.Workbook.Names;
            for (var i = 0; i < names.length; ++i) {
                if (names[i].Name.toLowerCase() != name.toLowerCase())
                    continue;
                if (names[i].Sheet != null)
                    continue;
                return new DefinedName_1.DefinedName(names[i]);
            }
        }
        throw new Error("Cannot find defined name |" + name + "|");
    };
    Object.defineProperty(WorkbookDefinedNameCollection.prototype, "count", {
        /**
         * Number of global defined names
         */
        get: function () {
            if (!this._wb.Workbook)
                return 0;
            if (!this._wb.Workbook.Names)
                return 0;
            return this._wb.Workbook.Names.filter(function (name) { typeof name.Sheet == "undefined"; }).length;
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Add or update defined name
     * @param name String name
     * @param ref Range object or string range/formula
     * @param comment Optional comment
     */
    WorkbookDefinedNameCollection.prototype.add = function (name, ref /*TODO: | Range */, comment) {
        try {
            return this.getName(name);
        }
        catch (e) {
            var nm = { Name: name, Ref: ref.toString(), Comment: comment || "" };
            if (!this._wb.Workbook)
                this._wb.Workbook = {};
            if (!this._wb.Workbook.Names)
                this._wb.Workbook.Names = [];
            this._wb.Workbook.Names.push(nm);
            return new DefinedName_1.DefinedName(nm);
        }
    };
    return WorkbookDefinedNameCollection;
}());
exports.WorkbookDefinedNameCollection = WorkbookDefinedNameCollection;
;
//# sourceMappingURL=WorkbookDefinedNameCollection.js.map