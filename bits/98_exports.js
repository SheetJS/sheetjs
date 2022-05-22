if(typeof parse_xlscfb !== "undefined") XLSX.parse_xlscfb = parse_xlscfb;
XLSX.parse_zip = parse_zip;
XLSX.read = readSync; //xlsread
XLSX.readFile = readFileSync; //readFile
XLSX.readFileSync = readFileSync;
XLSX.write = writeSync;
XLSX.writeFile = writeFileSync;
XLSX.writeFileSync = writeFileSync;
XLSX.writeFileAsync = writeFileAsync;
XLSX.utils = utils;
XLSX.writeXLSX = writeSyncXLSX;
XLSX.writeFileXLSX = writeFileSyncXLSX;
XLSX.SSF = SSF;
if(typeof __stream !== "undefined") XLSX.stream = __stream;
if(typeof CFB !== "undefined") XLSX.CFB = CFB;
if(typeof require !== "undefined") {
  var strmod = require('stream');
  if((strmod||{}).Readable) set_readable(strmod.Readable);
	try { _fs = require('fs'); } catch(e) {}
}
