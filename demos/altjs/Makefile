.PHONY: all
all: duktape nashorn

.PHONY: base
base:
	if [ ! -e sheetjs.xlsx ]; then node ../../tests/write.js; fi

.PHONY: duktape
duktape: base ## duktape / skookum demo
	sjs skookum.js

.PHONY: nashorn
nashorn: base ## nashorn demo
	if [ ! -e jvm-npm.js ]; then curl -O https://rawgit.com/nodyn/jvm-npm/master/src/main/javascript/jvm-npm.js; fi
	jjs nashorn.js
