/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
// @deno-types="https://unpkg.com/xlsx/types/index.d.ts"
import { read, utils, set_cptable } from 'https://unpkg.com/xlsx/xlsx.mjs';
import * as cptable from 'https://unpkg.com/xlsx/dist/cpexcel.full.mjs';
set_cptable(cptable);

import * as Drash from "https://deno.land/x/drash@v2.5.4/mod.ts";


// Create your resource

class HomeResource extends Drash.Resource {
  public paths = ["/"];

  public POST(request: Drash.Request, response: Drash.Response) {
    const file = request.bodyParam<Drash.Types.BodyFile>("file");
    if (!file) throw new Error("File is required!");
    var wb = read(file.content, {type: "buffer"});
    return response.html( utils.sheet_to_html(wb.Sheets[wb.SheetNames[0]]));
  }

  public GET(request: Drash.Request, response: Drash.Response): void {
    return response.html(`\
<!DOCTYPE html>
<html>
  <head>
    <title>SheetJS Spreadsheet to HTML Conversion Service</title>
    <meta charset="utf-8" />
  </head>
  <body>
<pre><h3><a href="//sheetjs.com/">SheetJS</a> Spreadsheet Conversion Service</h3>
<b>API</b>

Send a POST request to https://s2c.deno.dev/ with the file in the "file" body parameter:

$ curl -X POST -F"file=@test.xlsx" https://s2c.deno.dev/

The response will be an HTML TABLE generated from the first worksheet.

<b>Try it out!</b><form action="/" method="post" enctype="multipart/form-data">

<input type="file" name="file" />

Use the file input element to select a file, then click "Submit"

<button type="submit">Submit</button>
</form>
</pre>
  </body>
</html>`,
    );
  }
}

// Create and run your server
const server = new Drash.Server({
  hostname: "",
  port: 3000,
  protocol: "http",
  resources: [
    HomeResource,
  ],
});

server.run();

console.log(`Server running at ${server.address}.`);

