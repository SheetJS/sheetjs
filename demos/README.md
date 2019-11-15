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

**Frameworks and APIs**
- [`angularjs`](angular/)
- [`angular 2 / 4 / 5 / 6 and ionic`](angular2/)
- [`knockout`](knockout/)
- [`meteor`](meteor/)
- [`react and react-native`](react/)
- [`vue 2.x and weex`](vue/)
- [`XMLHttpRequest and fetch`](xhr/)
- [`nodejs server`](server/)
- [`databases and key/value stores`](database/)
- [`typed arrays and math`](array/)

**Bundlers and Tooling**
- [`browserify`](browserify/)
- [`fusebox`](fusebox/)
- [`parcel`](parcel/)
- [`requirejs`](requirejs/)
- [`rollup`](rollup/)
- [`systemjs`](systemjs/)
- [`typescript`](typescript/)
- [`webpack 2.x`](webpack/)

**Platforms and Integrations**
- [`electron application`](electron/)
- [`nw.js application`](nwjs/)
- [`Chrome / Chromium extensions`](chrome/)
- [`Adobe ExtendScript`](extendscript/)
- [`Headless Browsers`](headless/)
- [`canvas-datagrid`](datagrid/)
- [`Swift JSC and other engines`](altjs/)
- [`"serverless" functions`](function/)
- [`internet explorer`](oldie/)

Other examples are included in the [showcase](demos/showcase/).

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
