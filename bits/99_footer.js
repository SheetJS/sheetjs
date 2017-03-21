};
make_ssf(SSF);
/*global module */
/*:: declare var DO_NOT_EXPORT_SSF: any; */
if(typeof module !== 'undefined' && typeof DO_NOT_EXPORT_SSF === 'undefined') module.exports = SSF;
