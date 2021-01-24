/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
var XLSX = require("xlsx");
var SheetJSRedis = require("./SheetJSRedis");
var assert = require('assert');
var redis = require("redis"), util = require("util");
var client = redis.createClient();


/* Sample data */
var init = [
  ["FLUSHALL", []],
  ["SADD",  ["birdpowers", "flight", "pecking"]],
  ["SET",   ["foo", "bar"]],
  ["SET",   ["baz", 0]],
  ["RPUSH", ["friends", "sam", "alice", "bob"]],
  ["ZADD",  ["hackers", 1906, 'Grace Hopper', 1912, 'Alan Turing', 1916, 'Claude Shannon', 1940, 'Alan Kay', 1953, 'Richard Stallman', 1957, 'Sophie Wilson', 1965, 'Yukihiro Matsumoto', 1969, 'Linus Torvalds']],
  ["SADD",  ["superpowers", "flight", 'x-ray vision']],
  ["HMSET", ["user:1000", "name", 'John Smith', "email", 'john.smith@example.com', "password", "s3cret", "visits", 1]],
  ["HMSET", ["user:1001", "name", 'Mary Jones', "email", 'mjones@example.com', "password", "hidden"]]
];

const R = (()=>{
  const Rcache = {};
  const R_ = (n) => Rcache[n] || (Rcache[n] = util.promisify(client[n]).bind(client));
  return (n) => R_(n.toLowerCase());
})();

(async () => {
  for(var i = 0; i < init.length; ++i) await R(init[i][0])(init[i][1]);

  /* Export database to XLSX */
  var wb = await SheetJSRedis.redis_to_wb(R);
  XLSX.writeFile(wb, "redis.xlsx");

  /* Import XLSX to database */
  await R("flushall")();
  var wb2 = XLSX.readFile("redis.xlsx");
  await SheetJSRedis.wb_to_redis(wb2, R);

  /* Verify */
  assert.equal(await R("get")("foo"), "bar");
  assert.equal(await R("lindex")("friends", 1), "alice");
  assert.equal(await R("zscore")("hackers", "Claude Shannon"), 1916);
  assert.equal(await R("hget")("user:1000", "name"), "John Smith");
  assert.equal(await R("sismember")("superpowers", "flight"), "1");
  assert.equal(await R("sismember")("birdpowers", "pecking"), "1");

  client.quit();
})();
