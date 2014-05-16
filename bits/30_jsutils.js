function isval(x) { return typeof x !== "undefined" && x !== null; }

function keys(o) { return Object.keys(o).filter(function(x) { return o.hasOwnProperty(x); }); }

function evert(obj, arr) {
	var o = {};
	keys(obj).forEach(function(k) {
		if(!obj.hasOwnProperty(k)) return;
		if(!arr) o[obj[k]] = k;
		else (o[obj[k]]=o[obj[k]]||[]).push(k);
	});
	return o;
}
