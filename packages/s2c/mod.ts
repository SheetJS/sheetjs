// @deno-types="https://cdn.sheetjs.com/xlsx-latest/package/types/index.d.ts"
import { read, set_cptable } from 'https://cdn.sheetjs.com/xlsx-latest/package/xlsx.mjs';
import * as cptable from 'https://cdn.sheetjs.com/xlsx-latest/package/dist/cpexcel.full.mjs';
set_cptable(cptable);

import type { Request } from "https://deno.land/x/drash@v2.5.4/mod.ts";
import { Types } from "https://deno.land/x/drash@v2.5.4/mod.ts";

/**
 * Parse a workbook from an uploaded file
 *
 * This works with Deno Deploy (Drash body parser does not use temp files)
 *
 * request is a Drash.Request object
 * field is the name of the field to read
 */
export function parse_book_from_request(request: Request, field: string) {
  const file = request.bodyParam<Types.BodyFile>(field);
  if(!file) throw new Error(`Field ${field} is missing!`);
  return read(file.content, { type: "buffer" });
}
