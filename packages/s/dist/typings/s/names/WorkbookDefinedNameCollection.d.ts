/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
/// <reference path="../../../../src/xlsx.d.ts" />
import { DefinedNameCollection } from "./DefinedNameCollection";
import { DefinedName } from "./DefinedName";
export declare class WorkbookDefinedNameCollection implements DefinedNameCollection {
    private readonly _wb;
    constructor(wb: XLSX.WorkBook);
    /**
     * Get read-only array of global defined names
     */
    readonly items: DefinedName[];
    /**
     * Get defined name object
     */
    getName(name: string): DefinedName;
    /**
     * Number of global defined names
     */
    readonly count: number;
    /**
     * Add or update defined name
     * @param name String name
     * @param ref Range object or string range/formula
     * @param comment Optional comment
     */
    add(name: string, ref: string, comment?: string): DefinedName;
}
