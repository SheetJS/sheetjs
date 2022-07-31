# Demos

These demos are intended to demonstrate how to load this library in various
ecosystems.  The library is designed to be used in the web browser and in node
contexts, using dynamic feature tests to pull in features when necessary.  This
works extremely well in common use cases: script tag insertion and node require.

Systems like webpack try to be clever by performing simple static analysis to
pull in code.  However, they do not support dynamic type tests, breaking
compatibility with traditional scripts.  Configuration is required.  The demos
cover basic configuration steps for various systems and should "just work".

Mobile app and other larger demos do not include the full build structure. The
demos have `Makefile` scripts that show how to reproduce the full projects.  The
scripts have been tested against iOS and OSX.  For Windows platforms, GNU make
can be installed with Bash on Windows or with `cygwin`.

### Included Demos

**JavaScript APIs**
- [`XMLHttpRequest and fetch`](xhr/)
- [`Clipboard Data`](https://docs.sheetjs.com/docs/getting-started/demos/clipboard)
- [`Typed Arrays for Machine Learning`](https://docs.sheetjs.com/docs/getting-started/demos/ml)
- [`LocalStorage and SessionStorage`](https://docs.sheetjs.com/docs/getting-started/demos/database#localstorage-and-sessionstorage)
- [`Web SQL Database`](https://docs.sheetjs.com/docs/getting-started/demos/database#websql)
- [`IndexedDB`](https://docs.sheetjs.com/docs/getting-started/demos/database#indexeddb)

**Frameworks**
- [`Angular.JS`](angular/)
- [`Angular 2+ and Ionic`](angular2/)
- [`Knockout`](https://docs.sheetjs.com/docs/getting-started/demos/legacy#knockoutjs)
- [`Meteor`](meteor/)
- [`React, React Native and NextJS`](react/)
- [`VueJS, WeeX and NuxtJS`](vue/)

**Front-End UI Components**
- [`canvas-datagrid`](datagrid/)
- [`x-spreadsheet`](xspreadsheet/)
- [`react-data-grid`](react/modify/)
- [`vue3-table-light`](vue/modify/)

**Platforms and Integrations**
- [`NodeJS Server-Side Processing`](server/)
- [`Deno`](deno/)
- [`Electron`](electron/)
- [`NW.js`](nwjs/)
- [`Chrome / Chromium Extension`](chrome/)
- [`Google Sheets API`](https://docs.sheetjs.com/docs/getting-started/demos/gsheet)
- [`ExtendScript for Adobe Apps`](https://docs.sheetjs.com/docs/getting-started/demos/extendscript)
- [`NetSuite SuiteScript`](https://docs.sheetjs.com/docs/getting-started/demos/netsuite)
- [`SalesForce Lightning Web Components`](https://docs.sheetjs.com/docs/getting-started/demos/salesforce)
- [`Excel JavaScript API`](https://docs.sheetjs.com/docs/getting-started/demos/excel)
- [`Headless Automation`](https://docs.sheetjs.com/docs/getting-started/demos/headless)
- [`Swift JSC and Other JavaScript Engines`](altjs/)
- [`"serverless" functions`](function/)
- [`databases and key/value stores`](database/)
- [`Databases and Structured Data Stores`](https://docs.sheetjs.com/docs/getting-started/demos/database)
- [`NoSQL, K/V, and Unstructured Data Stores`](https://docs.sheetjs.com/docs/getting-started/demos/nosql)
- [`internet explorer`](oldie/)

**Bundlers and Tooling**
- [`browserify`](https://docs.sheetjs.com/docs/getting-started/demos/bundler#browserify)
- [`bun`](https://docs.sheetjs.com/docs/getting-started/demos/bundler#bun)
- [`esbuild`](https://docs.sheetjs.com/docs/getting-started/demos/bundler#esbuild)
- [`parcel`](https://docs.sheetjs.com/docs/getting-started/demos/bundler#parcel)
- [`requirejs`](https://docs.sheetjs.com/docs/getting-started/demos/bundler#requirejs)
- [`rollup`](https://docs.sheetjs.com/docs/getting-started/demos/bundler#rollup)
- [`snowpack`](https://docs.sheetjs.com/docs/getting-started/demos/bundler#snowpack)
- [`swc`](https://docs.sheetjs.com/docs/getting-started/demos/bundler#swc)
- [`systemjs`](systemjs/)
- [`typescript`](typescript/)
- [`vite`](https://docs.sheetjs.com/docs/getting-started/demos/bundler#vite)
- [`webpack 2.x`](webpack/)
- [`wmr`](https://docs.sheetjs.com/docs/getting-started/demos/bundler#wmr)

Other examples are included in the [showcase](demos/showcase/).

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
