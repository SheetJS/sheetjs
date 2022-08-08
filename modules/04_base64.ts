var Base64_map = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
function Base64_encode(input: string): string {
	var o = "";
	var c1=0, c2=0, c3=0, e1=0, e2=0, e3=0, e4=0;
	for(var i = 0; i < input.length; ) {
		c1 = input.charCodeAt(i++);
		e1 = (c1 >> 2);

		c2 = input.charCodeAt(i++);
		e2 = ((c1 & 3) << 4) | (c2 >> 4);

		c3 = input.charCodeAt(i++);
		e3 = ((c2 & 15) << 2) | (c3 >> 6);
		e4 = (c3 & 63);
		if (isNaN(c2)) { e3 = e4 = 64; }
		else if (isNaN(c3)) { e4 = 64; }
		o += Base64_map.charAt(e1) + Base64_map.charAt(e2) + Base64_map.charAt(e3) + Base64_map.charAt(e4);
	}
	return o;
}
function Base64_encode_pass(input: string): string {
	var o = "";
	var c1=0, c2=0, c3=0, e1=0, e2=0, e3=0, e4=0;
	for(var i = 0; i < input.length; ) {
		c1 = input.charCodeAt(i++); if(c1 > 0xFF) c1 = 0x5F;
		e1 = (c1 >> 2);

		c2 = input.charCodeAt(i++); if(c2 > 0xFF) c2 = 0x5F;
		e2 = ((c1 & 3) << 4) | (c2 >> 4);

		c3 = input.charCodeAt(i++); if(c3 > 0xFF) c3 = 0x5F;
		e3 = ((c2 & 15) << 2) | (c3 >> 6);
		e4 = (c3 & 63);
		if (isNaN(c2)) { e3 = e4 = 64; }
		else if (isNaN(c3)) { e4 = 64; }
		o += Base64_map.charAt(e1) + Base64_map.charAt(e2) + Base64_map.charAt(e3) + Base64_map.charAt(e4);
	}
	return o;
}
function Base64_decode(input: string): string {
	var o = "";
	var c1=0, c2=0, c3=0, e1=0, e2=0, e3=0, e4=0;
	input = input.replace(/^data:([^\/]+\/[^\/]+)?;base64\,/,'')
	             .replace(/[^\w\+\/\=]/g, "")
	for(var i = 0; i < input.length;) {
		e1 = Base64_map.indexOf(input.charAt(i++));
		e2 = Base64_map.indexOf(input.charAt(i++));
		c1 = (e1 << 2) | (e2 >> 4);
		o += String.fromCharCode(c1);

		e3 = Base64_map.indexOf(input.charAt(i++));
		c2 = ((e2 & 15) << 4) | (e3 >> 2);
		if (e3 !== 64) { o += String.fromCharCode(c2); }

		e4 = Base64_map.indexOf(input.charAt(i++));
		c3 = ((e3 & 3) << 6) | e4;
		if (e4 !== 64) { o += String.fromCharCode(c3); }
	}
	return o;
}
