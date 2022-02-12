import { createRequire } from 'https://deno.land/std/node/module.ts';
const require = createRequire(import.meta.url);
const XLSX = require('../../');

import doit from './doit.ts';
doit(XLSX, "node");
