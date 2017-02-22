function isval(x/*:?any*/)/*:boolean*/ { return x !== undefined && x !== null; }

function keys(o/*:any*/)/*:Array<any>*/ { return Object.keys(o); }

function evert_key(obj/*:any*/, key/*:string*/)/*:EvertType*/ {
	var o = ([]/*:any*/), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) o[obj[K[i]][key]] = K[i];
	return o;
}

function evert(obj/*:any*/)/*:EvertType*/ {
	var o = ([]/*:any*/), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) o[obj[K[i]]] = K[i];
	return o;
}

function evert_num(obj/*:any*/)/*:EvertNumType*/ {
	var o = ([]/*:any*/), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) o[obj[K[i]]] = parseInt(K[i],10);
	return o;
}

function evert_arr(obj/*:any*/)/*:EvertArrType*/ {
	var o/*:EvertArrType*/ = ([]/*:any*/), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) {
		if(o[obj[K[i]]] == null) o[obj[K[i]]] = [];
		o[obj[K[i]]].push(K[i]);
	}
	return o;
}

/* TODO: date1904 logic */
function datenum(v/*:number*/, date1904/*:?boolean*/)/*:number*/ {
	if(date1904) v+=1462;
	var epoch = Date.parse(v);
	return (epoch + 2209161600000) / (24 * 60 * 60 * 1000);
}

function cc2str(arr/*:Array<number>*/)/*:string*/ {
	var o = "";
	for(var i = 0; i != arr.length; ++i) o += String.fromCharCode(arr[i]);
	return o;
}

function str2cc(str) {
	var o = [];
	for(var i = 0; i != str.length; ++i) o.push(str.charCodeAt(i));
	return o;
}

function dup(o/*:any*/)/*:any*/ {
	if(typeof JSON != 'undefined') return JSON.parse(JSON.stringify(o));
	if(typeof o != 'object' || !o) return o;
	var out = {};
	for(var k in o) if(o.hasOwnProperty(k)) out[k] = dup(o[k]);
	return out;
}

function fill(c/*:string*/,l/*:number*/)/*:string*/ { var o = ""; while(o.length < l) o+=c; return o; }
