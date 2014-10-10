/* copied from js-xls (C) SheetJS Apache2 license */
function xlml_normalize(d) {
	if(has_buf && Buffer.isBuffer(d)) return d.toString('utf8');
	if(typeof d === 'string') return d;
	throw "badf";
}

var xlmlregex = /<(\/?)([a-z0-9]*:|)([\w-]+)[^>]*>/mg;
