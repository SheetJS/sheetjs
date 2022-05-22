if(typeof cptable !== 'undefined') set_cptable(cptable);
else if(typeof module !== "undefined" && typeof require !== 'undefined') {
	set_cptable(require('./dist/cpexcel.js'));
}
