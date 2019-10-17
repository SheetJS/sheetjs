/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
/* vim: set ts=2: */

/// <reference path="../../xlsx.d.ts"/>

import { DefinedNameCollection } from "./DefinedNameCollection";
import { DefinedName } from "./DefinedName";
//import { Range } from "../Range";

export class WorkbookDefinedNameCollection implements DefinedNameCollection {
  private readonly _wb: XLSX.WorkBook;

  constructor(wb: XLSX.WorkBook) {
    this._wb = wb;
  };

  /**
   * Get read-only array of global defined names
   */
  get items(): DefinedName[] {
    if(!this._wb.Workbook) return [];
    if(!this._wb.Workbook.Names) return [];
    return this._wb.Workbook.Names.filter((name) => {name.Sheet == null}).map((name) => new DefinedName(name));
  };

  /**
   * Get defined name object
   */
  getName(name: string): DefinedName {
    if(this._wb.Workbook && this._wb.Workbook.Names) {
      const names = this._wb.Workbook.Names;
      for(let i = 0; i < names.length; ++i) {
        if(names[i].Name.toLowerCase() != name.toLowerCase()) continue;
        if(names[i].Sheet != null) continue;
        return new DefinedName(names[i]);
      }
    }
    throw new Error(`Cannot find defined name |${name}|`);
  }

  /**
   * Number of global defined names
   */
  get count(): number {
    if(!this._wb.Workbook) return 0;
    if(!this._wb.Workbook.Names) return 0;
    return this._wb.Workbook.Names.filter((name) => {typeof name.Sheet == "undefined"}).length;
  }

  /**
   * Add or update defined name
   * @param name String name
   * @param ref Range object or string range/formula
   * @param comment Optional comment
   */
  add(name: string, ref: string /*TODO: | Range */, comment?: string): DefinedName {
    try {
      return this.getName(name);
    } catch(e) {
      const nm = { Name: name, Ref: ref.toString(), Comment: comment || "" };
      if(!this._wb.Workbook) this._wb.Workbook = {};
      if(!this._wb.Workbook.Names) this._wb.Workbook.Names = [];
      this._wb.Workbook.Names.push(nm);
      return new DefinedName(nm);
    }
  }
};