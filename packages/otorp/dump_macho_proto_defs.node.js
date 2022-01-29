#!/usr/bin/env node

// dump_macho_proto_defs.node.ts
var import_fs = require("fs");
var import_path = require("path");
var import__ = require("./");
if (!process.argv[2] || process.argv[2] == "-h" || process.argv[2] == "--help") {
  [
    "usage: otorp <path/to/bin> [output/folder]",
    "  if no output folder specified, log all discovered defs",
    "  if output folder specified, attempt to write defs in the folder"
  ].map((x) => console.error(x));
  process.exit(1);
}
var buf = (0, import_fs.readFileSync)(process.argv[2]);
var otorps = (0, import__.otorp)(buf);
otorps.forEach(({ name, proto }) => {
  if (!process.argv[3]) {
    console.log(proto);
  } else {
    var pth = (0, import_path.resolve)(process.argv[3] || "./", name.replace(/[/]/g, "$"));
    console.error(`writing ${name} to ${pth}`);
    (0, import_fs.writeFileSync)(pth, proto);
  }
});
/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
