function isval(x) { return typeof x !== "undefined" && x !== null; }

function keys(o) { return Object.keys(o).filter(function(x) { return o.hasOwnProperty(x); }); }

function evert(obj, arr) {
	var o = {};
	keys(obj).forEach(function(k) {
		if(!obj.hasOwnProperty(k)) return;
		if(arr && typeof arr === "string") o[obj[k][arr]] = k;
		if(!arr) o[obj[k]] = k;
		else (o[obj[k]]=o[obj[k]]||[]).push(k);
	});
	return o;
}

/* TODO: date1904 logic */
function datenum(v, date1904) {
	if(date1904) v+=1462;
	var epoch = Date.parse(v);
	return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}
