# Headless Browsers

The library, eschewing unstable and nascent ECMAScript features, plays nicely
with most headless browsers.  This demo shows a few common headless scenarios.

## PhantomJS

This was tested in PhantomJS 2.1.1, installed using the node module:

```bash
$ npm install -g phantomjs
$ phantomjs phantomjs.js
```

## Chrome Automation

This was tested in puppeteer 0.9.0 (Chromium revision 494755) and `chromeless`:

```bash
$ npm install puppeteer
$ node puppeteer.js

$ npm install -g chromeless
$ node chromeless.js
```

Since the main process is node, the read and write features should be placed in
the webpage.  The `dist` versions are suitable for web pages.


## wkhtmltopdf

This was tested in wkhtmltopdf 0.12.4, installed using the official binaries:

```bash
$ wkhtmltopdf --javascript-delay 20000 http://oss.sheetjs.com/js-xlsx/tests/ test.pdf
```

## SlimerJS

This was tested in SlimerJS 0.10.3 and FF 52.0, installed using `brew` on OSX:

```bash
$ brew install slimerjs
$ slimerjs slimerjs.js
```

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
