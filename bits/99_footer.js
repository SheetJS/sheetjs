}
/*global define */
/*:: declare var define:any; */
if(typeof exports !== 'undefined') make_xlsx_lib(exports);
else if(typeof module !== 'undefined' && module.exports) make_xlsx_lib(module.exports);
else if(typeof define === 'function' && define.amd) define(function() { if(!XLSX.version) make_xlsx_lib(XLSX); return XLSX; });
else make_xlsx_lib(XLSX);
/*exported XLS, ODS */
var XLS = XLSX, ODS = XLSX;
