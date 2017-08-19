# Headless Browsers

The library, intentionally conservative in the use of ES5+ feaures, plays nicely
with most headless browsers.  This demo shows a few common headless scenarios.

## PhantomJS

This was tested in phantomjs 2.1.1, installed using the node module:

```bash
$ npm install -g phantomjs
$ phantomjs phantomjs.js
```

## wkhtmltopdf

This was tested in wkhtmltopdf 0.12.4, installed using the official binaries:

```bash
$ wkhtmltopdf --javascript-delay 60000 http://localhost:8000/ test.pdf
``` 

## Puppeteer

This was tested in puppeteer 0.9.0 and Chromium r494755, installed using node:

```bash
$ npm install puppeteer
$ node puppeteer.js
```

Since the main process is node, the read and write features should be placed in
the webpage.  The `dist` versions are suitable for web pages.

## SlimerJS

This was tested in slimerjs 0.10.3 and FF 52.0, installed using `brew` on OSX:

```bash
$ brew install slimerjs
$ slimerjs slimerjs.js
```

