## Testing

### Node

<details>
	<summary>(click to show)</summary>

`make test` will run the node-based tests.  By default it runs tests on files in
every supported format.  To test a specific file type, set `FMTS` to the format
you want to test.  Feature-specific tests are avaialble with `make test_misc`

```bash
$ make test_misc   # run core tests
$ make test        # run full tests
$ make test_xls    # only use the XLS test files
$ make test_xlsx   # only use the XLSX test files
$ make test_xlsb   # only use the XLSB test files
$ make test_xml    # only use the XML test files
$ make test_ods    # only use the ODS test files
```

To enable all errors, set the environment variable `WTF=1`:

```bash
$ make test        # run full tests
$ WTF=1 make test  # enable all error messages
```

Flow and eslint checks are available:

```bash
$ make lint        # eslint checks
$ make flow        # make lint + Flow checking
```

</details>

### Browser

<details>
	<summary>(click to show)</summary>

The core in-browser tests are available at `tests/index.html` within this repo.
Start a local server and navigate to that directory to run the tests.
`make ctestserv` will start a server on port 8000.

`make ctest` will generate the browser fixtures.  To add more files, edit the
`tests/fixtures.lst` file and add the paths.

To run the full in-browser tests, clone the repo for
[oss.sheetjs.com](https://github.com/SheetJS/SheetJS.github.io) and replace
the xlsx.js file (then fire up the browser and go to `stress.html`):

```bash
$ cp xlsx.js ../SheetJS.github.io
$ cd ../SheetJS.github.io
$ simplehttpserver # or "python -mSimpleHTTPServer" or "serve"
$ open -a Chromium.app http://localhost:8000/stress.html
```
</details>

### Tested Environments

<details>
	<summary>(click to show)</summary>

 - NodeJS 0.8, 0.9, 0.10, 0.11, 0.12, 4.x, 5.x, 6.x, 7.x
 - IE 6/7/8/9/10/11 (IE6-9 browsers require shims for interacting with client)
 - Chrome 24+
 - Safari 6+
 - FF 18+

Tests utilize the mocha testing framework.  Travis-CI and Sauce Labs links:

 - <https://travis-ci.org/SheetJS/js-xlsx> for XLSX module in nodejs
 - <https://semaphoreci.com/sheetjs/js-xlsx> for XLSX module in nodejs
 - <https://travis-ci.org/SheetJS/SheetJS.github.io> for XLS\* modules
 - <https://saucelabs.com/u/sheetjs> for XLS\* modules using Sauce Labs

</details>

### Test Files

Test files are housed in [another repo](https://github.com/SheetJS/test_files).

Running `make init` will refresh the `test_files` submodule and get the files.



