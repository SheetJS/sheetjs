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
- [`XMLHttpRequest and fetch`](https://docs.sheetjs.com/docs/demos/network)
- [`Clipboard Data`](https://docs.sheetjs.com/docs/demos/clipboard)
- [`Typed Arrays for Machine Learning`](https://docs.sheetjs.com/docs/demos/ml)
- [`LocalStorage and SessionStorage`](https://docs.sheetjs.com/docs/demos/database#localstorage-and-sessionstorage)
- [`Web SQL Database`](https://docs.sheetjs.com/docs/demos/database#websql)
- [`IndexedDB`](https://docs.sheetjs.com/docs/demos/database#indexeddb)

**Frameworks**
- [`Angular 2+ and Ionic`](https://docs.sheetjs.com/docs/demos/angular)
- [`React`](https://docs.sheetjs.com/docs/demos/react)
- [`VueJS`](https://docs.sheetjs.com/docs/demos/vue)
- [`Angular.JS`](https://docs.sheetjs.com/docs/demos/legacy#angularjs)
- [`Knockout`](https://docs.sheetjs.com/docs/demos/legacy#knockoutjs)

**Front-End UI Components**
- [`canvas-datagrid`](https://docs.sheetjs.com/docs/demos/grid#canvas-datagrid)
- [`x-spreadsheet`](https://docs.sheetjs.com/docs/demos/grid#x-spreadsheet)
- [`react-data-grid`](https://docs.sheetjs.com/docs/demos/grid#react-data-grid)
- [`vue3-table-lite`](https://docs.sheetjs.com/docs/demos/grid#vue3-table-lite)
- [`angular-ui-grid`](https://docs.sheetjs.com/docs/demos/grid#angular-ui-grid)

**Platforms and Integrations**
- [`Command-Line Tools`](https://docs.sheetjs.com/docs/demos/cli)
- [`iOS and Android Mobile Applications`](https://docs.sheetjs.com/docs/demos/mobile)
- [`NodeJS Server-Side Processing`](https://docs.sheetjs.com/docs/demos/server#nodejs)
- [`Content Management and Static Sites`](https://docs.sheetjs.com/docs/demos/content)
- [`Electron`](https://docs.sheetjs.com/docs/demos/desktop#electron)
- [`NW.js`](https://docs.sheetjs.com/docs/demos/desktop#nwjs)
- [`Tauri`](https://docs.sheetjs.com/docs/demos/desktop#tauri)
- [`Chrome and Chromium Extensions`](https://docs.sheetjs.com/docs/demos/chromium)
- [`Google Sheets API`](https://docs.sheetjs.com/docs/demos/gsheet)
- [`ExtendScript for Adobe Apps`](https://docs.sheetjs.com/docs/demos/extendscript)
- [`NetSuite SuiteScript`](https://docs.sheetjs.com/docs/demos/netsuite)
- [`SalesForce Lightning Web Components`](https://docs.sheetjs.com/docs/demos/salesforce)
- [`Excel JavaScript API`](https://docs.sheetjs.com/docs/demos/excel)
- [`Headless Automation`](https://docs.sheetjs.com/docs/demos/headless)
- [`Other JavaScript Engines`](https://docs.sheetjs.com/docs/demos/engines)
- [`Azure Functions and Storage`](https://docs.sheetjs.com/docs/demos/azure)
- [`Amazon Web Services`](https://docs.sheetjs.com/docs/demos/aws)
- [`Databases and Structured Data Stores`](https://docs.sheetjs.com/docs/demos/database)
- [`NoSQL and Unstructured Data Stores`](https://docs.sheetjs.com/docs/demos/nosql)
- [`Legacy Internet Explorer`](https://docs.sheetjs.com/docs/demos/legacy#internet-explorer)

**Bundlers and Tooling**
- [`browserify`](https://docs.sheetjs.com/docs/demos/bundler#browserify)
- [`bun`](https://docs.sheetjs.com/docs/demos/bundler#bun)
- [`esbuild`](https://docs.sheetjs.com/docs/demos/bundler#esbuild)
- [`parcel`](https://docs.sheetjs.com/docs/demos/bundler#parcel)
- [`requirejs`](https://docs.sheetjs.com/docs/demos/bundler#requirejs)
- [`rollup`](https://docs.sheetjs.com/docs/demos/bundler#rollup)
- [`snowpack`](https://docs.sheetjs.com/docs/demos/bundler#snowpack)
- [`swc`](https://docs.sheetjs.com/docs/demos/bundler#swc)
- [`systemjs`](https://docs.sheetjs.com/docs/demos/bundler#systemjs)
- [`vite`](https://docs.sheetjs.com/docs/demos/bundler#vite)
- [`webpack`](https://docs.sheetjs.com/docs/demos/bundler#webpack)
- [`wmr`](https://docs.sheetjs.com/docs/demos/bundler#wmr)

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
