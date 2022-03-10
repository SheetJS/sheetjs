// nuxt.config.js
import { readFile, utils } from 'xlsx';

const parseXLSX = (file, { path }) => {
  const wb = readFile(path);
  const o = wb.SheetNames.map(name => ({ name, data: utils.sheet_to_json(wb.Sheets[name])}));
  return { data: o };
}

export default {
  modules: [ '@nuxt/content' ],
  content: {
    extendParser: {
      ".numbers": parseXLSX,
      ".xlsx": parseXLSX,
      ".xls": parseXLSX
      // ...
    }
  },
}
