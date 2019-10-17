/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
/* vim: set ts=2: */

/// <reference path="../xlsx.d.ts"/>

let _XLSX: typeof XLSX;
if(typeof XLSX !== "undefined") _XLSX = XLSX;

export function get_XLSX() { return _XLSX; };
export function set_XLSX(xlsx: typeof XLSX) { _XLSX = xlsx; };