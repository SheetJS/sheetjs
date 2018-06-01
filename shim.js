/*! shim.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* ES3/5 Compatibility shims and other utilities for older browsers. */

// From https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object/keys
if(!Object.keys) Object.keys = (function() {
  var hasOwnProperty = Object.prototype.hasOwnProperty,
      hasDontEnumBug = !({toString: null}).propertyIsEnumerable('toString'),
      dontEnums = [
        'toString',
        'toLocaleString',
        'valueOf',
        'hasOwnProperty',
        'isPrototypeOf',
        'propertyIsEnumerable',
        'constructor'
      ],
      dontEnumsLength = dontEnums.length;

  return function(obj) {
    if(typeof obj !== 'object' && typeof obj !== 'function' || obj === null) throw new TypeError('Object.keys called on non-object');

    var result = [];

    for(var prop in obj) if(hasOwnProperty.call(obj, prop)) result.push(prop);

    if(hasDontEnumBug)
      for(var i=0; i < dontEnumsLength; ++i)
        if(hasOwnProperty.call(obj, dontEnums[i])) result.push(dontEnums[i]);
    return result;
  };
})();

if(!String.prototype.trim) String.prototype.trim = function() {
  var s = this.replace(/^\s+/, '');
  for(var i = s.length - 1; i >=0 ; --i) if(!s.charAt(i).match(/^\s/)) return s.slice(0,i+1);
  return "";
};

if(!Array.prototype.forEach) Array.prototype.forEach = function(cb) {
  var len = (this.length>>>0), self = (arguments[1]||void 0);
  for(var i=0; i<len; ++i) if(i in this) self ? cb.call(self, this[i], i, this) : cb(this[i], i, this);
};

if(!Array.prototype.map) Array.prototype.map = function(cb) {
  var len = (this.length>>>0), self = (arguments[1]||void 0), A = new Array(len);
  for(var i=0; i<len; ++i) if(i in this) A[i] = self ? cb.call(self, this[i], i, this) : cb(this[i], i, this);
  return A;
};

if(!Array.prototype.indexOf) Array.prototype.indexOf = function(needle) {
  var len = (this.length>>>0), i = ((arguments[1]|0)||0);
  for(i<0 && (i+=len)<0 && (i=0); i<len; ++i) if(this[i] === needle) return i;
  return -1;
};

if(!Array.prototype.lastIndexOf) Array.prototype.lastIndexOf = function(needle) {
  var len = (this.length>>>0), i = len - 1;
  for(; i>=0; --i) if(this[i] === needle) return i;
  return -1;
};

if(!Array.isArray) Array.isArray = function(obj) { return Object.prototype.toString.call(obj) === "[object Array]"; };

if(!Date.prototype.toISOString) Date.prototype.toISOString = (function() {
  function p(n,i) { return ('0000000' + n).slice(-(i||2)); }

  return function _toISOString() {
    var y = this.getUTCFullYear(), yr = "";
    if(y>9999)   yr = '+' + p( y, 6);
    else if(y<0) yr = '-' + p(-y, 6);
    else         yr =       p( y, 4);

    return [
      yr, p(this.getUTCMonth()+1), p(this.getUTCDate())
    ].join('-') + 'T' + [
      p(this.getUTCHours()), p(this.getUTCMinutes()), p(this.getUTCSeconds())
    ].join(':') + '.' + p(this.getUTCMilliseconds(),3) + 'Z';
  };
}());

if(typeof ArrayBuffer !== 'undefined' && !ArrayBuffer.prototype.slice) ArrayBuffer.prototype.slice = function(start, end) {
  if(start == null) start = 0;
  if(start < 0) { start += this.byteLength; if(start < 0) start = 0; }
  if(start >= this.byteLength) return new Uint8Array(0);
  if(end == null) end = this.byteLength;
  if(end < 0) { end += this.byteLength; if(end < 0) end = 0; }
  if(end > this.byteLength) end = this.byteLength;
  if(start > end) return new Uint8Array(0);
  var out = new ArrayBuffer(end - start);
  var view = new Uint8Array(out);
  var data = new Uint8Array(this, start, end - start)
  /* IE10 should have Uint8Array#set */
  if(view.set) view.set(data); else while(start <= --end) view[end - start] = data[end];
  return out;
};
if(typeof Uint8Array !== 'undefined' && !Uint8Array.prototype.slice) Uint8Array.prototype.slice = function(start, end) {
  if(start == null) start = 0;
  if(start < 0) { start += this.length; if(start < 0) start = 0; }
  if(start >= this.length) return new Uint8Array(0);
  if(end == null) end = this.length;
  if(end < 0) { end += this.length; if(end < 0) end = 0; }
  if(end > this.length) end = this.length;
  if(start > end) return new Uint8Array(0);
  var out = new Uint8Array(end - start);
  while(start <= --end) out[end - start] = this[end];
  return out;
};

// VBScript + ActiveX fallback for IE5+
var IE_SaveFile = (function() { try {
  if(typeof IE_SaveFile_Impl == "undefined") document.write([
'<script type="text/vbscript" language="vbscript">',
'IE_GetProfileAndPath_Key = "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders\\"',
'Function IE_GetProfileAndPath(key): Set wshell = CreateObject("WScript.Shell"): IE_GetProfileAndPath = wshell.RegRead(IE_GetProfileAndPath_Key & key): IE_GetProfileAndPath = wshell.ExpandEnvironmentStrings("%USERPROFILE%") & "!" & IE_GetProfileAndPath: End Function',
'Function IE_SaveFile_Impl(FileName, payload): Dim data, plen, i, bit: data = CStr(payload): plen = Len(data): Set fso = CreateObject("Scripting.FileSystemObject"): fso.CreateTextFile FileName, True: Set f = fso.GetFile(FileName): Set stream = f.OpenAsTextStream(2, 0): For i = 1 To plen Step 3: bit = Mid(data, i, 2): stream.write Chr(CLng("&h" & bit)): Next: stream.Close: IE_SaveFile_Impl = True: End Function',
'|/script>'.replace("|","<")
  ].join("\r\n"));
  if(typeof IE_SaveFile_Impl == "undefined") return void 0;
  var IE_GetPath = (function() {
    var DDP1 = "";
    try { DDP1 = IE_GetProfileAndPath("{374DE290-123F-4565-9164-39C4925E467B}"); } catch(e) { try { DDP1 = IE_GetProfileAndPath("Personal"); } catch(e) { try { DDP1 = IE_GetProfileAndPath("Desktop"); } catch(e) { throw e; }}}
    var o = DDP1.split("!");
    DDP = o[1].replace("%USERPROFILE%", o[0]);
    return function(path) { return DDP + "\\" + path; };
  })();
  function fix_data(data) {
    var out = [];
    var T = typeof data == "string";
    for(var i = 0; i < data.length; ++i) out.push(("00"+(T ? data.charCodeAt(i) : data[i]).toString(16)).slice(-2));
    var o = out.join("|");
    return o;
  }
  return function(data, filename) { return IE_SaveFile_Impl(IE_GetPath(filename), fix_data(data)); };
} catch(e) { return void 0; }})();
var IE_LoadFile = (function() { try {
  if(typeof IE_LoadFile_Impl == "undefined") document.write([
'<script type="text/vbscript" language="vbscript">',
'Function IE_LoadFile_Impl(FileName): Dim out(), plen, i, cc: Set fso = CreateObject("Scripting.FileSystemObject"): Set f = fso.GetFile(FileName): Set stream = f.OpenAsTextStream(1, 0): plen = f.Size: ReDim out(plen): For i = 1 To plen Step 1: cc = Hex(Asc(stream.read(1))): If Len(cc) < 2 Then: cc = "0" & cc: End If: out(i) = cc: Next: IE_LoadFile_Impl = Join(out,""): End Function',
'|/script>'.replace("|","<")
  ].join("\r\n"));
  if(typeof IE_LoadFile_Impl == "undefined") return void 0;
  function fix_data(data) {
    var out = [];
    for(var i = 0; i < data.length; i+=2) out.push(String.fromCharCode(parseInt(data.slice(i, i+2), 16)));
    var o = out.join("");
    return o;
  }
  return function(filename) { return fix_data(IE_LoadFile_Impl(filename)); };
} catch(e) { return void 0; }})();

// getComputedStyle polyfill from https://gist.github.com/8HNHoFtE/5891086
if(typeof window !== 'undefined' && typeof window.getComputedStyle !== 'function') {
  window.getComputedStyle = function(e,t){return this.el=e,this.getPropertyValue=function(t){var n=/(\-([a-z]){1})/g;return t=="float"&&(t="styleFloat"),n.test(t)&&(t=t.replace(n,function(){return arguments[2].toUpperCase()})),e.currentStyle[t]?e.currentStyle[t]:null},this}
}
