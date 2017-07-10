var assert = function(bool) { if(!bool) { throw new Error("failed assert"); } };
assert.deepEqualArray = function(x,y) {
	if(x.length != y.length) throw new Error("Length mismatch: " + x.length + " != " + y.length);
	for(var i = 0; i < x.length; ++i) assert.deepEqual(x[i], y[i]);
};
assert.deepEqual = function(x,y) {
	if(x == y) return true;
	if(Array.isArray(x) && Array.isArray(y) && x.length > 5) return assert.deepEqualArray(x,y);
	if(typeof x != 'object' || typeof y != 'object') throw new Error(x + " !== " + y);
	Object.keys(x).forEach(function(k) { assert.deepEqual(x[k], y[k]); });
	Object.keys(y).forEach(function(k) { assert.deepEqual(x[k], y[k]); });
};
assert.notEqual = function(x,y) { if(x == y) throw new Error(x + " == " + y); };
assert.equal = function(x,y) { if(x != y) throw new Error(x + " !== " + y); };
assert.throws = function(cb) { var pass = true; try { cb(); pass = false; } catch(e) { } if(!pass) throw new Error("Function did not throw"); };
assert.doesNotThrow = function(cb) { var pass = true; try { cb(); } catch(e) { pass = false; } if(!pass) throw new Error("Function did throw"); };

function require(s) {
	switch(s) {
		case 'fs': return fs;
		case 'assert': return assert;
		case './': return XLSX;
	}
	if(s.slice(-5) == ".json") return JSON.parse(fs.readFileSync(s));
}

var fs = {};
fs.existsSync = function(p) { return !!fs[p]; };
fs.readdirSync = function(p) { return Object.keys(fs).filter(function(n) {
	return fs.hasOwnProperty(n) && n.slice(-4) != "Sync"; });
};
fs.readFileSync = function(f, enc) {
	if(!fs[f]) throw new Error("File not found: " + f);
	fs[f].length;
	switch(enc) {
		case 'base64': return fs[f];
		case 'buffer':
			var o = atob(fs[f]), oo = [];
			for(var i = 0; i < o.length; ++i) oo[i] = o.charCodeAt(i);
			return oo;
		default: return atob(fs[f]);
	}
};

fs.writeFileSync = function(f, d) { fs[f] = d; };

