/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
/* vim: set ts=2: */

/// <reference path="../../xlsx.d.ts"/>

export class DefinedName {
  readonly _name: XLSX.DefinedName;

  constructor(name: XLSX.DefinedName) {
    this._name = name;
  };

  get raw(): XLSX.DefinedName { return this._name; };

};