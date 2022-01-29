}
/*global define */
/*:: declare var define:any; */
if(typeof exports !== 'undefined') make_xlsx_lib(exports);
else if(typeof module !== 'undefined' && module.exports) make_xlsx_lib(module.exports);
else if(typeof define === 'function' && define.amd) define('xlsx', function() { if(!XLSX.version) make_xlsx_lib(XLSX); return XLSX; });
else make_xlsx_lib(XLSX);
/* NOTE: the following extra line is needed for "Lightning Locker Service" */
if(typeof window !== 'undefined' && !window.XLSX) try { window.XLSX = XLSX; } catch(e) {}
