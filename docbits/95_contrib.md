## Contributing

Due to the precarious nature of the Open Specifications Promise, it is very
important to ensure code is cleanroom.  [Contribution Notes](CONTRIBUTING.md)

<details>
  <summary><b>File organization</b> (click to show)</summary>

At a high level, the final script is a concatenation of the individual files in
the `bits` folder.  Running `make` should reproduce the final output on all
platforms.  The README is similarly split into bits in the `docbits` folder.

Folders:

| folder       | contents                                                      |
|:-------------|:--------------------------------------------------------------|
| `bits`       | raw source files that make up the final script                |
| `docbits`    | raw markdown files that make up `README.md`                   |
| `bin`        | server-side bin scripts (`xlsx.njs`)                          |
| `dist`       | dist files for web browsers and nonstandard JS environments   |
| `demos`      | demo projects for platforms like ExtendScript and Webpack     |
| `tests`      | browser tests (run `make ctest` to rebuild)                   |
| `types`      | typescript definitions and tests                              |
| `misc`       | miscellaneous supporting scripts                              |
| `test_files` | test files (pulled from the test files repository)            |

</details>

After cloning the repo, running `make help` will display a list of commands.

### OSX/Linux

<details>
  <summary>(click to show)</summary>

The `xlsx.js` file is constructed from the files in the `bits` subdirectory. The
build script (run `make`) will concatenate the individual bits to produce the
script.  Before submitting a contribution, ensure that running make will produce
the `xlsx.js` file exactly.  The simplest way to test is to add the script:

```bash
$ git add xlsx.js
$ make clean
$ make
$ git diff xlsx.js
```

To produce the dist files, run `make dist`.  The dist files are updated in each
version release and *should not be committed between versions*.
</details>

### Windows

<details>
  <summary>(click to show)</summary>

The included `make.cmd` script will build `xlsx.js` from the `bits` directory.
Building is as simple as:

```cmd
> make
```

To prepare development environment:

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

As explained in [Test Files](#test-files), on Windows the release ZIP file must
be downloaded and extracted.  If Bash on Windows is available, it is possible
to run the OSX/Linux workflow.  The following steps prepares the environment:

```bash
# Install support programs for the build and test commands
sudo apt-get install make git subversion mercurial

# Install nodejs and NPM within the WSL
wget -qO- https://deb.nodesource.com/setup_8.x | sudo bash
sudo apt-get install nodejs

# Install dev dependencies
sudo npm install -g mocha voc blanket xlsjs
```

</details>

### Tests

<details>
  <summary>(click to show)</summary>

The `test_misc` target (`make test_misc` on Linux/OSX / `make misc` on Windows)
runs the targeted feature tests.  It should take 5-10 seconds to perform feature
tests without testing against the entire test battery.  New features should be
accompanied with tests for the relevant file formats and features.

For tests involving the read side, an appropriate feature test would involve
reading an existing file and checking the resulting workbook object.  If a
parameter is involved, files should be read with different values to verify that
the feature is working as expected.

For tests involving a new write feature which can already be parsed, appropriate
feature tests would involve writing a workbook with the feature and then opening
and verifying that the feature is preserved.

For tests involving a new write feature without an existing read ability, please
add a feature test to the kitchen sink `tests/write.js`.
</details>

