## Contributing

Due to the precarious nature of the Open Specifications Promise, it is very
important to ensure code is cleanroom.  Consult CONTRIBUTING.md

### Tests

The `test_misc` target (`make test_misc` on Linux/OSX / `make misc` on Windows)
runs the targeted feature tests.  It should take 5-10 seconds to perform feature
tests without testing against the entire test battery.  New features should be
accompanied with tests for the relevant file formats and features.

For tests involving the read side, an appropriate feature test would involve
reading an existing file and checking the resulting workbook object.  If a
parameter is involved, files should be read with different values for the param
to verify that the feature is working as expected.

For tests involving a new write feature which can already be parsed, appropriate
feature tests would involve writing a workbook with the feature and then opening
and verifying that the feature is preserved.

For tests involving a new write feature without an existing read ability, please
add a feature test to the kitchen sink `tests/write.js`.

### OSX/Linux

The xlsx.js file is constructed from the files in the `bits` subdirectory. The
build script (run `make`) will concatenate the individual bits to produce the
script.  Before submitting a contribution, ensure that running make will produce
the xlsx.js file exactly.  The simplest way to test is to add the script:

```bash
$ git add xlsx.js
$ make clean
$ make
$ git diff xlsx.js
```

To produce the dist files, run `make dist`.  The dist files are updated in each
version release and *should not be committed between versions*.

### Windows

The included `make.cmd` script will build `xlsx.js` from the `bits` directory.
Building is as simple as:

```cmd
> make
```

To prepare dev environment:

```cmd
> make init
```

The full list of commands available in Windows are displayed in `make help`:

```
make init -- install deps and global modules
make lint -- run eslint linter
make test -- run mocha test suite
make misc -- run smaller test suite
make book -- rebuild README and summary
make help -- display this message
```

The normal approach uses a variety of command line tools to grab the test files.
For windows users, please download the latest version of the test files snapshot
from [github](https://github.com/SheetJS/test_files/releases)

Latest test files snapshot:
<https://github.com/SheetJS/test_files/releases/download/20170409/test_files.zip>

Download and unzip to the `test_files` subdirectory.

