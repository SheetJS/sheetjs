/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
/* vim: set ts=2: */

import assert from "assert";
import * as S from "../esm";

/* song and dance for node 12 esm */
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const XLSX = require("../../../");

assert(S != null);
S.set_XLSX(XLSX);
assert(S.get_XLSX() == XLSX);
assert(S.get_XLSX().version);