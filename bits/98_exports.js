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
XLSX.SSF = SSF;
if(typeof CFB !== "undefined") XLSX.CFB = CFB;
