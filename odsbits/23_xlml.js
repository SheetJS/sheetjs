/* copied from js-xls (C) SheetJS Apache2 license */
function xlml_normalize(d) {
	if(has_buf &&/*::typeof Buffer !== "undefined" && d != null &&*/ Buffer.isBuffer(d)) return d.toString('utf8');
	if(typeof d === 'string') return d;
	throw "badf";
}

/* UOS uses CJK in tags, original regex /<(\/?)([a-z0-9]*:|)([\w-]+)[^>]*>/ */
var xlmlregex = /<(\/?)([^\s?>\/:]*:|)([^\s?>]*[^\s?>\/])[^>]*>/mg;
