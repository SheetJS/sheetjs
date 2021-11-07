# xlsx-cli

This is a standalone version of the CLI tool for [SheetJS](https://sheetjs.com).

For newer versions of node, the tool should be invoked with `npx`:

```bash
$ npx xlsx-cli --help             # help and usage info
$ npx xlsx-cli --xlsx test.csv    # generates test.csv.xlsx from test.csv
```

For older versions of node, the tool should be installed globally:

```bash
$ sudo npm install -g xlsx-cli    # install globally (once)

$ xlsx-cli --help                 # help and usage info
$ npx xlsx-cli --xlsx test.csv    # generates test.csv.xlsx from test.csv
```

