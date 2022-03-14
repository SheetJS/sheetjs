#### Cell Comments

Cell comments are objects stored in the `c` array of cell objects.  The actual
contents of the comment are split into blocks based on the comment author.  The
`a` field of each comment object is the author of the comment and the `t` field
is the plain text representation.

For example, the following snippet appends a cell comment into cell `A1`:

```js
if(!ws.A1.c) ws.A1.c = [];
ws.A1.c.push({a:"SheetJS", t:"I'm a little comment, short and stout!"});
```

Note: XLSB enforces a 54 character limit on the Author name.  Names longer than
54 characters may cause issues with other formats.

To mark a comment as normally hidden, set the `hidden` property:

```js
if(!ws.A1.c) ws.A1.c = [];
ws.A1.c.push({a:"SheetJS", t:"This comment is visible"});

if(!ws.A2.c) ws.A2.c = [];
ws.A2.c.hidden = true;
ws.A2.c.push({a:"SheetJS", t:"This comment will be hidden"});
```


_Threaded Comments_

Introduced in Excel 365, threaded comments are plain text comment snippets with
author metadata and parent references. They are supported in XLSX and XLSB.

To mark a comment as threaded, each comment part must have a true `T` property:

```js
if(!ws.A1.c) ws.A1.c = [];
ws.A1.c.push({a:"SheetJS", t:"This is not threaded"});

if(!ws.A2.c) ws.A2.c = [];
ws.A2.c.hidden = true;
ws.A2.c.push({a:"SheetJS", t:"This is threaded", T: true});
ws.A2.c.push({a:"JSSheet", t:"This is also threaded", T: true});
```

There is no Active Directory or Office 365 metadata associated with authors in a thread.

