@echo off
REM vim: set ts=2:

if "%1" == "help" (
	echo make init -- install deps and global modules
	echo make lint -- run eslint linter
	echo make test -- run mocha test suite
	echo              remember to download the test_files release!
	echo make misc -- run smaller test suite
	echo make book -- rebuild README and summary
	echo make help -- display this message
) else if "%1" == "init" (
	npm install
	npm install -g eslint eslint-plugin-html eslint-plugin-json
	npm install -g mocha markdown-toc
) else if "%1" == "lint" (
	eslint --ext .js,.njs,.json,.html,.htm xlsx.js xlsx.flow.js bin\xlsx.njs package.json bower.json
) else if "%1" == "test" (
	SET FMTS=
	mocha -R spec -t 30000
) else if "%1" == "misc" (
	SET FMTS=misc
	mocha -R spec -t 30000
) else if "%1" == "book" (
	type docbits\*.md > README.md
	markdown-toc -i README.md
) else (
	type bits\*.js > xlsx.flow.js
	node misc\strip_flow.js > xlsx.js
)
