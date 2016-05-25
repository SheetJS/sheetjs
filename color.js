function hex2RGB(h) {
	var o = h.substr(h[0]==="#"?1:0,6);
        var R = o.substr(0,2);
        var G = o.substr(2,2);
        var B = o.substr(4,2);
	return [parseInt(R,16),parseInt(G,16),parseInt(B,16)];
}
function rgb2Hex(rgb) {
	for(var i=0,o=1; i!=3; ++i) o = o*256 + (rgb[i]>255?255:rgb[i]<0?0:rgb[i]);
	return o.toString(16).toUpperCase().substr(1);
}

function rgb2HSL(rgb) {
	var R = rgb[0]/255, G = rgb[1]/255, B=rgb[2]/255;
	var M = Math.max(R, G, B), m = Math.min(R, G, B), C = M - m;
	if(C === 0) return [0, 0, R];

	var H6 = 0, S = 0, L2 = (M + m);
	S = C / (L2 > 1 ? 2 - L2 : L2);
	switch(M){
		case R: H6 = ((G - B) / C + 6)%6; break;
		case G: H6 = ((B - R) / C + 2); break;
		case B: H6 = ((R - G) / C + 4); break;
	}
	return [H6 / 6, S, L2 / 2];
}

function hsl2RGB(hsl){
	var H = hsl[0], S = hsl[1], L = hsl[2];
	var C = S * 2 * (L < 0.5 ? L : 1 - L), m = L - C/2;
	var rgb = [m,m,m], h6 = 6*H;

	var X;
	if(S !== 0) switch(h6|0) {
		case 0: case 6: X = C * h6; rgb[0] += C; rgb[1] += X; break;
		case 1: X = C * (2 - h6);   rgb[0] += X; rgb[1] += C; break;
		case 2: X = C * (h6 - 2);   rgb[1] += C; rgb[2] += X; break;
		case 3: X = C * (4 - h6);   rgb[1] += X; rgb[2] += C; break;
		case 4: X = C * (h6 - 4);   rgb[2] += C; rgb[0] += X; break;
		case 5: X = C * (6 - h6);   rgb[2] += X; rgb[0] += C; break;
	}
	for(var i = 0; i != 3; ++i) rgb[i] = Math.round(rgb[i]*255);
	return rgb;
}
/* 18.8.3 bgColor tint algorithm */
function rgb_tint(hex, tint) {
	if(tint === 0) return hex;
	var hsl = rgb2HSL(hex2RGB(hex));
	if (tint < 0) hsl[2] = hsl[2] * (1 + tint);
	else hsl[2] = 1 - (1 - hsl[2]) * (1 - tint);
	return rgb2Hex(hsl2RGB(hsl));
}

 var exp = [
      { patternType: 'darkHorizontal',
        fgColor: { theme: 9, "tint":-0.249977111117893, rgb: 'F79646' },
        bgColor: { theme: 5, "tint":0.3999755851924192, rgb: 'C0504D' } },
      { patternType: 'darkUp',
        fgColor: { theme: 3, "tint":-0.249977111117893, rgb: 'EEECE1' },
        bgColor: { theme: 7, "tint":0.3999755851924192, rgb: '8064A2' } },
      { patternType: 'darkGray',
        fgColor: { theme: 3, rgb: 'EEECE1' },
        bgColor: { theme: 1, rgb: 'FFFFFF' } },
      { patternType: 'lightGray',
        fgColor: { theme: 6, "tint":0.3999755851924192, rgb: '9BBB59' },
        bgColor: { theme: 2, "tint":-0.499984740745262, rgb: '1F497D' } },
      { patternType: 'lightDown',
        fgColor: { theme: 4, rgb: '4F81BD' },
        bgColor: { theme: 7, rgb: '8064A2' } },
      { patternType: 'lightGrid',
        fgColor: { theme: 6, "tint":-0.249977111117893, rgb: '9BBB59' },
        bgColor: { theme: 9, "tint":-0.249977111117893, rgb: 'F79646' } },
      { patternType: 'lightGrid',
        fgColor: { theme: 4, rgb: '4F81BD' },
        bgColor: { theme: 2, "tint":-0.749992370372631, rgb: '1F497D' } },
      { patternType: 'lightVertical',
        fgColor: { theme: 3, "tint":0.3999755851924192, rgb: 'EEECE1' },
        bgColor: { theme: 7, "tint":0.3999755851924192, rgb: '8064A2' } }
    ];
 var map = [];
 exp.forEach(function(e) {
    e.fgColor.new =  rgb_tint( e.fgColor.rgb,  e.fgColor.tint || 0); 
console.log(e.fgColor.rgb, e.fgColor.new);
    e.bgColor.new =  rgb_tint( e.bgColor.rgb,  e.bgColor.tint || 0); 
console.log(e.bgColor.rgb, e.bgColor.new);
 });
