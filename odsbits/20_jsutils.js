var has_buf = (typeof Buffer !== 'undefined');

function cc2str(arr) {
	var o = "";
	for(var i = 0; i != arr.length; ++i) o += String.fromCharCode(arr[i]);
	return o;
}
