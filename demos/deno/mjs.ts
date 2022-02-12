import * as XLSX from '../../xlsx.mjs';
import * as cpexcel from '../../dist/cpexcel.full.mjs';
XLSX.set_cptable(cpexcel);

import doit from './doit.ts';
doit(XLSX, "mjs");
