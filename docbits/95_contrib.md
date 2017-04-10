## Contributing

Due to the precarious nature of the Open Specifications Promise, it is very
important to ensure code is cleanroom.  Consult CONTRIBUTING.md

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
> make.cmd
```

To prepare dev environment:

```cmd
> npm install -g mocha
> npm install
> mocha -t 30000
```

The normal approach uses a variety of command line tools to grab the test files.
For windows users, please download the latest version of the test files snapshot
from [github](https://github.com/SheetJS/test_files/releases)

Latest test files snapshot:
<https://github.com/SheetJS/test_files/releases/download/20170409/test_files.zip>

Download and unzip to the `test_files` subdirectory.

