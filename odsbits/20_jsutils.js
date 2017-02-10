var has_buf = (typeof Buffer !== 'undefined');

function cc2str(arr)/*:string*/ {
	var o = "";
	for(var i = 0; i != arr.length; ++i) o += String.fromCharCode(arr[i]);
	return o;
}

function dup(o/*:any*/)/*:any*/ {
	if(typeof JSON != 'undefined') return JSON.parse(JSON.stringify(o));
	if(typeof o != 'object' || !o) return o;
	var out = {};
	for(var k in o) if(o.hasOwnProperty(k)) out[k] = dup(o[k]);
	return out;
}
