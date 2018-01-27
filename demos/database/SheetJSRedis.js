/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
var XLSX = require("xlsx");

const pair = (arr) => arr.map((x,i)=>!(i%2)&&[x,+arr[i+1]]).filter(x=>x);
const keyify = (obj) => Object.keys(obj).map(x => [x, obj[x]]);

async function redis_to_wb(R) {
	var wb = XLSX.utils.book_new();
	var manifest = [], strs = [];

	/* store strings in strs and keep note of other objects in manifest */
	var keys = await R("keys")("*"), type = "";
	for(var i = 0; i < keys.length; ++i) {
		type = await R("type")(keys[i]);
		switch(type) {
			case "string": strs.push({key:keys[i], value: await R("get")(keys[i])}); break;
			case "list": case "zset": case "set": case "hash": manifest.push({key:keys[i], type:type}); break;
			default: throw new Error("bad type " + type);
		}
	}

	/* add worksheets if relevant */
	if(strs.length > 0) {
		var wss = XLSX.utils.json_to_sheet(strs, {header: ["key", "value"], skipHeader:1});
		XLSX.utils.book_append_sheet(wb, wss, "_strs");
	}
	if(manifest.length > 0) {
		var wsm = XLSX.utils.json_to_sheet(manifest, {header: ["key", "type"]});
		XLSX.utils.book_append_sheet(wb, wsm, "_manifest");
	}
	for(i = 0; i < manifest.length; ++i) {
		var sn = "obj" + i;
		var aoa, key = manifest[i].key;
		switch((type=manifest[i].type)) {
			case "list":
				aoa = (await R("lrange")(key, 0, -1)).map(x => [x]); break;
			case "set":
				aoa = (await R("smembers")(key)).map(x => [x]); break;
			case "zset":
				aoa = pair(await R("zrange")(key, 0, -1, "withscores")); break;
			case "hash":
				aoa = keyify(await R("hgetall")(key)); break;
			default: throw new Error("bad type " + type);
		}
		XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), sn);
	}
	return wb;
}

/* convert worksheet aoa to specific redis type */
const aoa_to_redis = {
	list: async (aoa, R, key) => await R("RPUSH")([key].concat(aoa.map(x=>x[0]))),
	zset: async (aoa, R, key) => await R("ZADD" )([key].concat(aoa.reduce((acc,x)=>acc.concat([+x[1], x[0]]), []))),
	hash: async (aoa, R, key) => await R("HMSET")([key].concat(aoa.reduce((acc,x)=>acc.concat(x), []))),
	set:  async (aoa, R, key) => await R("SADD" )([key].concat(aoa.map(x=>x[0])))
};
async function wb_to_redis(wb, R) {
	if(wb.Sheets._strs) {
		var strs = XLSX.utils.sheet_to_json(wb.Sheets._strs, {header:1});
		for(var i = 0; i < strs.length; ++i) await R("SET")(strs[i]);
	}
	if(!wb.Sheets._manifest) return;
	var M = XLSX.utils.sheet_to_json(wb.Sheets._manifest);
	for(i = 0; i < M.length; ++i) {
		var aoa = XLSX.utils.sheet_to_json(wb.Sheets["obj" + i], {header:1});
		await aoa_to_redis[M[i].type](aoa, R, M[i].key);
	}
}
module.exports = {
	redis_to_wb,
	wb_to_redis
};
