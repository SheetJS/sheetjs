# Other JS Engines and Deployments

There are many JS engines and deployments outside of web browsers. NodeJS is the
most popular deployment, but there are many others for special use cases.  Some
optimize for low overhead and others optimize for ease of embedding within other
applications.  Since it was designed for ES3 engines, the library can be used in
those settings!  This demo tries to demonstrate a few alternative deployments.

Some engines provide no default global object.  To create a global reference:

```js
var global = (function(){ return this; }).call(null);
```


## Swift + JavaScriptCore

iOS and OSX ship with the JavaScriptCore framework, enabling easy JS access from
Swift and Objective-C.  Hybrid function invocation is tricky, but explicit data
passing is straightforward.

Binary strings can be passed back and forth using `String.Encoding.ascii`.


## Nashorn

Nashorn ships with Java 8.  It includes a command-line tool `jjs` for running JS
scripts.  It is somewhat limited but does offer access to the full Java runtime.

`jjs` does not provide a CommonJS `require` implementation.  This demo uses a
[`shim`](https://rawgit.com/nodyn/jvm-npm/master/src/main/javascript/jvm-npm.js)
and manually requires the library.

The Java `nio` API provides the `Files.readAllBytes` method to read a file into
a byte array.  To use in `XLSX.read`, the demo copies the bytes into a plain JS
array and calls `XLSX.read` with type `"array"`.


## Rhino

[Rhino](http://www.mozilla.org/rhino) is an ES3+ engine written in Java.  The
`SheetJSRhino` class and `com.sheetjs` package show a complete JAR deployment,
including the full XLSX source.

Due to code generation errors, optimization must be disabled:

```java
Context context = Context.enter();
context.setOptimizationLevel(-1);
```


## ChakraCore

ChakraCore is an embeddable JS engine written in C++.  The library and binary
distributions include a command-line tool `chakra` for running JS scripts.

The simplest way to interop with the engine is to pass Base64 strings.  The make
target builds a very simple payload with the data.


## Duktape

[Duktape](http://duktape.org/) is an embeddable JS engine written in C.  The
amalgamation makes integration extremely simple!  It supports `Buffer` natively:

```C
/* parse a C char array as a workbook object */
duk_push_external_buffer(ctx);
duk_config_buffer(ctx, -1, buf, len);
duk_put_global_string(ctx, "buf");
duk_eval_string_noresult("workbook = XLSX.read(buf, {type:'buffer'});");

/* write a workbook object to a C char array */
duk_eval_string(ctx, "XLSX.write(workbook, {type:'buffer', bookType:'xlsx'})");
duk_size_t sz;
char *buf = (char *)duk_get_buffer_data(ctx, -1, sz);
duk_pop(ctx);
```
