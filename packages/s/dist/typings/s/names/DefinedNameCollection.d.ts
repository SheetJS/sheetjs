/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
import { DefinedName } from "./DefinedName";
import { Range } from "../Range";
export interface DefinedNameCollection {
    readonly items: DefinedName[];
    readonly count: number;
    add(name: string, ref: string | Range, comment?: string): DefinedName;
    getName(name: string): DefinedName;
}
