import * as S from "../esm";
const XLSX = require("./xlsx.mini.min.js");
function assert(x) { if(!x) throw "assert failed"; }

assert(S != null);
S.set_XLSX(XLSX);
assert(S.get_XLSX() == XLSX);
assert(S.get_XLSX().version);