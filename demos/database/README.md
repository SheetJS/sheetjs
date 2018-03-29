# Databases

"Database" is a catch-all term referring to traditional RDBMS as well as K/V
stores, document databases, and other "NoSQL" storages. There are many external
database systems as well as browser APIs like WebSQL and `localStorage`

This demo discusses general strategies and provides examples for a variety of
database systems.  The examples are merely intended to demonstrate very basic
functionality.


## Structured Tables

Database tables are a common import and export target for spreadsheets.  One
common representation of a database table is an array of JS objects whose keys
are column headers and whose values are the underlying data values. For example,

| Name         | Index |
| :----------- | ----: |
| Barack Obama |    44 |
| Donald Trump |    45 |

is naturally represented as an array of objects

```js
[
  { Name: "Barack Obama", Index: 44 },
  { Name: "Donald Trump", Index: 45 }
]
```

The `sheet_to_json` and `json_to_sheet` helper functions work with objects of
similar shape, converting to and from worksheet objects.  The corresponding
worksheet would include a header row for the labels:

```
XXX|      A       |   B   |
---+--------------+-------+
 1 | Name         | Index |
 2 | Barack Obama |    44 |
 3 | Donald Trump |    45 |
```


## Building Schemas from Worksheets

The `sheet_to_json` helper function generates arrays of JS objects that can be
scanned to determine the column "types", and there are third-party connectors
that can push arrays of JS objects to database tables.

The [`sexql`](http://sheetjs.com/sexql) browser demo uses WebSQL, which is
limited to the SQLite fundamental types.

<details>
	<summary><b>Implementation details</b> (click to show)</summary>

The `sexql` schema builder scans the first row to find headers:

```js
  if(!ws || !ws['!ref']) return;
  var range = XLSX.utils.decode_range(ws['!ref']);
  if(!range || !range.s || !range.e || range.s > range.e) return;
  var R = range.s.r, C = range.s.c;

  var names = new Array(range.e.c-range.s.c+1);
  for(C = range.s.c; C<= range.e.c; ++C){
    var addr = XLSX.utils.encode_cell({c:C,r:R});
    names[C-range.s.c] = ws[addr] ? ws[addr].v : XLSX.utils.encode_col(C);
  }
```

After finding the headers, a deduplication step ensures that data is not lost.
Duplicate headers will be suffixed with `_1`, `_2`, etc.

```js
  for(var i = 0; i < names.length; ++i) if(names.indexOf(names[i]) < i)
    for(var j = 0; j < names.length; ++j) {
      var _name = names[i] + "_" + (j+1);
      if(names.indexOf(_name) > -1) continue;
      names[i] = _name;
    }
```

A column-major walk helps determine the data type.  For SQLite the only relevant
data types are `REAL` and `TEXT`.  If a string or date or error is seen in any
value of a column, the column is marked as `TEXT`:

```js
  var types = new Array(range.e.c-range.s.c+1);
  for(C = range.s.c; C<= range.e.c; ++C) {
    var seen = {}, _type = "";
    for(R = range.s.r+1; R<= range.e.r; ++R)
      seen[(ws[XLSX.utils.encode_cell({c:C,r:R})]||{t:"z"}).t] = true;
    if(seen.s || seen.str) _type = "TEXT";
    else if(seen.n + seen.b + seen.d + seen.e > 1) _type = "TEXT";
    else switch(true) {
      case seen.b:
      case seen.n: _type = "REAL"; break;
      case seen.e: _type = "TEXT"; break;
      case seen.d: _type = "TEXT"; break;
    }
    types[C-range.s.c] = _type || "TEXT";
  }
```

</details>

The included `SheetJSSQL.js` script demonstrates SQL statement generation.


## Objects, K/V and "Schema-less" Databases

So-called "Schema-less" databases allow for arbitrary keys and values within the
entries in the database.  K/V stores and Objects add additional restrictions.

There is no natural way to translate arbitrarily shaped schemas to worksheets
in a workbook.  One common trick is to dedicate one worksheet to holding named
keys.  For example, considering the JS object:

```json
{
  "title": "SheetDB",
  "metadata": {
    "author": "SheetJS",
    "code": 7262
  },
  "data": [
    { "Name": "Barack Obama", "Index": 44 },
    { "Name": "Donald Trump", "Index": 45 },
  ]
}
```

A dedicated worksheet should store the one-off named values:

```
XXX|        A        |    B    |
---+-----------------+---------+
 1 | Path            | Value   |
 2 | title           | SheetDB |
 3 | metadata.author | SheetJS |
 4 | metadata.code   |    7262 |
```

The included `ObjUtils.js` script demonstrates object-workbook conversion:

<details>
	<summary><b>Implementation details</b> (click to show)</summary>

```js
function deepset(obj, path, value) {
  if(path.indexOf(".") == -1) return obj[path] = value;
  var parts = path.split(".");
  if(!obj[parts[0]]) obj[parts[0]] = {};
  return deepset(obj[parts[0]], parts.slice(1).join("."), value);
}
function workbook_to_object(wb) {
  var out = {};

  /* assign one-off keys */
  var ws = wb.Sheets["_keys"]; if(ws) {
    var data = XLSX.utils.sheet_to_json(ws, {raw:true});
    data.forEach(function(r) { deepset(out, r.path, r.value); });
  }

  /* assign arrays from worksheet tables */
  wb.SheetNames.forEach(function(n) {
    if(n == "_keys") return;
    out[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], {raw:true});
  });

  return out;
}

function walk(obj, key, arr) {
  if(Array.isArray(obj)) return;
  if(typeof obj != "object") { arr.push({path:key, value:obj}); return; }
  Object.keys(obj).forEach(function(k) { walk(obj[k], key?key+"."+k:k, arr); });
}
function object_to_workbook(obj) {
  var wb = XLSX.utils.book_new();

  /* keyed entries */
  var base = []; walk(obj, "", base);
  var ws = XLSX.utils.json_to_sheet(base, {header:["path", "value"]});
  XLSX.utils.book_append_sheet(wb, ws, "_keys");

  /* arrays */
  Object.keys(obj).forEach(function(k) {
    if(!Array.isArray(obj[k])) return;
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(obj[k]), k);
  });

  return wb;
}
```

</details>


## Browser APIs

#### WebSQL

WebSQL is a popular SQL-based in-browser database available on Chrome / Safari.
In practice, it is powered by SQLite, and most simple SQLite-compatible queries
work as-is in WebSQL.

The public demo <http://sheetjs.com/sexql> generates a database from workbook.

#### LocalStorage and SessionStorage

The Storage API, encompassing `localStorage` and `sessionStorage`, describes
simple key-value stores that only support string values and keys. Objects can be
stored as JSON using `JSON.stringify` and `JSON.parse` to set and get keys.

`SheetJSStorage.js` extends the `Storage` prototype with a `load` function to
populate the db based on an object and a `dump` function to generate a workbook
from the data in the storage.  `LocalStorage.html` tests `localStorage`.

#### IndexedDB

IndexedDB is a more complex storage solution, but the `localForage` wrapper
supplies a Promise-based interface mimicking the `Storage` API.

`SheetJSForage.js` extends the `localforage` object with a `load` function to
populate the db based on an object and a `dump` function to generate a workbook
from the data in the storage.  `LocalForage.html` forces IndexedDB mode.


## External Database Demos

### SQL Databases

There are nodejs connector libraries for all of the popular RDBMS systems.  They
have facilities for connecting to a database, executing queries, and obtaining
results as arrays of JS objects that can be passed to `json_to_sheet`.  The main
differences surround API shape and supported data types.

#### SQLite

[The `better-sqlite3` module](https://www.npmjs.com/package/better-sqlite3)
provides a very simple API for working with SQLite databases.  `Statement#all`
runs a prepared statement and returns an array of JS objects.

`SQLiteTest.js` generates a simple two-table SQLite database (`SheetJS1.db`),
exports to XLSX (`sqlite.xlsx`), imports the new XLSX file to a new database
(`SheetJS2.db`) and verifies the tables are preserved.

#### MySQL / MariaDB

[The `mysql2` module](https://www.npmjs.com/package/mysql2) supplies a callback
API as well as a Promise wrapper.  `Connection#query` runs a statement and
returns an array whose first element is an array of JS objects.

`MySQLTest.js` connects to the MySQL instance running on `localhost`, builds two
tables in the `sheetjs` database, exports to XLSX, imports the new XLSX file to
the `sheetj5` database and verifies the tables are preserved.

#### PostgreSQL

[The `pg` module](https://www.npmjs.com/package/pg) supplies a Promise wrapper.
Like with `mysql2`, `Client#query` runs a statement and returns a result object.
The `rows` key of the object is an array of JS objects.

`PgSQLTest.js` connects to the PostgreSQL server on `localhost`, builds two
tables in the `sheetjs` database, exports to XLSX, imports the new XLSX file to
the `sheetj5` database and verifies the tables are preserved.

#### Knex Query Builder

[The `knex` module](https://www.npmjs.com/package/knex) builds SQL queries.  The
same exact code can be used against Oracle Database, MSSQL, and other engines.

`KnexTest.js` uses the `sqlite3` connector and follows the same procedure as the
SQLite test.  The included `SheetJSKnex.js` script converts between the query
builder and the common spreadsheet format.

### Key/Value Stores

#### Redis

Redis is a powerful data structure server that can store simple strings, sets,
sorted sets, hashes and lists.  One simple database representation stores the
strings in a special worksheet (`_strs`), the manifest in another worksheet
(`_manifest`), and each object in its own worksheet (`obj##`).

`RedisTest.js` connects to a local Redis server, populates data based on the
official Redis tutorial, exports to XLSX, flushes the server, imports the new
XLSX file and verifies the data round-tripped correctly.  `SheetJSRedis.js`
includes the implementation details.

#### LowDB

LowDB is a small schemaless database powered by `lodash`.  `_.get` and `_.set`
helper functions make storing metadata a breeze.  The included `SheetJSLowDB.js`
script demonstrates a simple adapter that can load and dump data.

### Document Databases

Since document databases are capable of holding more complex objects, they can
actually hold the underlying worksheet objects!  In some cases, where arrays are
supported, they can even hold the workbook object.

#### MongoDB

MongoDB is a popular document-oriented database engine.  `MongoDBTest.js` uses
MongoDB to hold a simple workbook and export to XLSX.

`MongoDBCRUD.js` follows the SQL examples using an idiomatic collection
structure.  Exporting and importing collections are straightforward:

<details>
	<summary><b>Example code</b> (click to show)</summary>

```js
/* generate a worksheet from a collection */
const aoa = await db.collection('coll').find({}).toArray();
aoa.forEach((x) => delete x._id);
const ws = XLSX.utils.json_to_sheet(aoa);


/* import data from a worksheet to a collection */
const aoa = XLSX.utils.sheet_to_json(ws);
await db.collection('coll').insertMany(aoa, {ordered: true});
```

#### Firebase

[`firebase-server`](https://www.npmjs.com/package/firebase-server) is a simple
mock Firebase server used in the tests, but the same code works in an external
Firebase deployment when plugging in the database connection info.

`FirebaseDemo.html` and `FirebaseTest.js` demonstrate a whole-workbook process.
The entire workbook object is persisted, a few cells are changed, and the stored
data is dumped and exported to XLSX.

</details>

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
