# Headless Browsers

The library, eschewing unstable and nascent ECMAScript features, plays nicely
with most headless browsers.  This demo shows a few common headless scenarios.

NodeJS does not ship with its own layout engine.  For advanced HTML exports, a
headless browser is generally indistinguishable from a browser process.

## Chromium Automation with Puppeteer

[Puppeteer](https://npm.im/puppeteer) enables headless Chromium automation.

[`html.js`](./html.js) shows a dedicated script for converting an HTML file to
XLSB using puppeteer.  The first argument is the path to the HTML file.  The
script writes to `output.xlsb`:

```bash
# read from test.html and write to output.xlsb
$ node html.js test.html
```

The script pulls up the webpage using headless Chromium and adds a script tag
reference to the standalone browser build.  That will make the `XLSX` variable
available to future scripts added in the page!  The browser context is not able
to save the file using `writeFile`, so the demo generates the XLSB spreadsheet
bytes with the `base64` type, sends the string back to the main process, and
uses `fs.writeFileSync` to write the file.

## WebKit Automation with PhantomJS

This was tested using [PhantomJS 2.1.1](https://phantomjs.org/download.html)

```bash
$ phantomjs phantomjs.js
```

The flow is similar to the Puppeteer flow (scrape table and generate workbook in
website context, copy string back, write string to file from main process).

The `binary` type generates strings that can be written in PhantomJS using the
`fs.write` method with mode `"wb"`.

## wkhtmltopdf

This was tested in wkhtmltopdf 0.12.4, installed using the official binaries:

```bash
$ wkhtmltopdf --javascript-delay 20000 http://oss.sheetjs.com/sheetjs/tests/ test.pdf
```


[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
