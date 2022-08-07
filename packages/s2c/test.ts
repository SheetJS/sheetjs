// @deno-types="https://cdn.sheetjs.com/xlsx-latest/package/types/index.d.ts"
import { utils, set_cptable, version } from 'https://cdn.sheetjs.com/xlsx-latest/package/xlsx.mjs';
import { parse_book_from_request } from "./mod.ts";
import * as Drash from "https://deno.land/x/drash@v2.5.4/mod.ts";

class S2CResource extends Drash.Resource {
  public paths = ["/"];

  // see https://github.com/drashland/drash/issues/194
  public OPTIONS(request: Drash.Request, response: Drash.Response) {
    const allHttpMethods: string[] = [ "GET", "POST", "PUT", "DELETE" ];
    response.headers.set("Allow", allHttpMethods.join()); // Use this
    response.headers.set("Access-Control-Allow-Methods", allHttpMethods.join()); // or this
    response.headers.set("access-control-allow-origin", "*");
    response.status_code = 204;
    return response;
  }

  public POST(request: Drash.Request, response: Drash.Response) {
    try { response.headers.set("access-control-allow-origin", "*"); } catch(e) {}
    var wb = parse_book_from_request(request, "file");
    return response.html( utils.sheet_to_html(wb.Sheets[wb.SheetNames[0]]));
  }

  public GET(request: Drash.Request, response: Drash.Response): void {
    try { response.headers.set("access-control-allow-origin", "*"); } catch(e) {}
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

Library Version: ${version}
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
    S2CResource,
  ],
});

server.run();

console.log(`Server running at ${server.address}.`);

