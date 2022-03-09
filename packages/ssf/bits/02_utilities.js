function _strrev(x/*:string*/)/*:string*/ { var o = "", i = x.length-1; while(i>=0) o += x.charAt(i--); return o; }
function fill(c/*:string*/,l/*:number*/)/*:string*/ { var o = ""; while(o.length < l) o+=c; return o; }
function pad0(v/*:any*/,d/*:number*/)/*:string*/{var t=""+v; return t.length>=d?t:fill('0',d-t.length)+t;}
function pad_(v/*:any*/,d/*:number*/)/*:string*/{var t=""+v;return t.length>=d?t:fill(' ',d-t.length)+t;}
function rpad_(v/*:any*/,d/*:number*/)/*:string*/{var t=""+v; return t.length>=d?t:t+fill(' ',d-t.length);}
function pad0r1(v/*:any*/,d/*:number*/)/*:string*/{var t=""+Math.round(v); return t.length>=d?t:fill('0',d-t.length)+t;}
function pad0r2(v/*:any*/,d/*:number*/)/*:string*/{var t=""+v; return t.length>=d?t:fill('0',d-t.length)+t;}
var p2_32 = Math.pow(2,32);
function pad0r(v/*:any*/,d/*:number*/)/*:string*/{if(v>p2_32||v<-p2_32) return pad0r1(v,d); var i = Math.round(v); return pad0r2(i,d); }
function isgeneral(s/*:string*/, i/*:?number*/)/*:boolean*/ { i = i || 0; return s.length >= 7 + i && (s.charCodeAt(i)|32) === 103 && (s.charCodeAt(i+1)|32) === 101 && (s.charCodeAt(i+2)|32) === 110 && (s.charCodeAt(i+3)|32) === 101 && (s.charCodeAt(i+4)|32) === 114 && (s.charCodeAt(i+5)|32) === 97 && (s.charCodeAt(i+6)|32) === 108; }
