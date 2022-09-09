SHELL=/bin/bash
LIB=xlsx
FMT=xlsx xlsm xlsb ods xls xml misc full
REQS=
ADDONS=dist/cpexcel.js
AUXTARGETS=
CMDS=bin/xlsx.njs
HTMLLINT=index.html

MINITGT=xlsx.mini.js
MINIFLOW=xlsx.mini.flow.js
MINIDEPS=$(shell cat misc/mini.lst)

ESMJSTGT=xlsx.mjs
ESMJSDEPS=$(shell cat misc/mjs.lst)


ULIB=$(shell echo $(LIB) | tr a-z A-Z)
DEPS=$(sort $(wildcard bits/*.js))
TSBITS=$(patsubst modules/%,bits/%,$(wildcard modules/[0-9][0-9]_*.js))
MTSBITS=$(patsubst modules/%,misc/%,$(wildcard modules/[0-9][0-9]_*.js))
TARGET=$(LIB).js
FLOWTARGET=$(LIB).flow.js
FLOWAUX=$(patsubst %.js,%.flow.js,$(AUXTARGETS))
AUXSCPTS=xlsxworker.js
FLOWTGTS=$(TARGET) $(AUXTARGETS) $(AUXSCPTS) $(MINITGT)
UGLIFYOPTS=--support-ie8 -m
CLOSURE=/usr/local/lib/node_modules/google-closure-compiler/compiler.jar

## Main Targets

.PHONY: all
all: $(TARGET) $(AUXTARGETS) $(AUXSCPTS) $(MINITGT) $(ESMJSTGT) ## Build library and auxiliary scripts

$(FLOWTGTS): %.js : %.flow.js
	node -e 'process.stdout.write(require("fs").readFileSync("$<","utf8").replace(/^[ \t]*\/\*[:#][^*]*\*\/\s*(\n)?/gm,"").replace(/\/\*[:#][^*]*\*\//gm,""))' > $@

$(FLOWTARGET): $(DEPS)
	cat $^ | tr -d '\15\32' > $@

$(MINIFLOW): $(MINIDEPS)
	cat $^ | tr -d '\15\32' > $@

$(ESMJSTGT): $(ESMJSDEPS)
	cat $^ | tr -d '\15\32' > $@

bits/01_version.js: package.json
	echo "$(ULIB).version = '"`grep version package.json | awk '{gsub(/[^0-9a-z\.-]/,"",$$2); print $$2}'`"';" > $@

#bits/18_cfb.js: node_modules/cfb/xlscfb.flow.js
#	cp $^ $@

$(TSBITS): bits/%: modules/%
	cp $^ $@

$(MTSBITS): misc/%: modules/%
	cp $^ $@


.PHONY: clean
clean: ## Remove targets and build artifacts
	rm -f $(TARGET) $(FLOWTARGET) $(ESMJSTGT) $(MINITGT) $(MINIFLOW)

.PHONY: clean-data
clean-data:
	rm -f *.xlsx *.xlsm *.xlsb *.xls *.xml

.PHONY: init
init: ## Initial setup for development
	git submodule init
	git submodule update
	#git submodule foreach git pull origin master
	git submodule foreach make all
	mkdir -p tmp

DISTHDR=misc/suppress_export.js
.PHONY: dist
dist: dist-deps $(TARGET) bower.json ## Prepare JS files for distribution
	mkdir -p dist
	cp LICENSE dist/
	uglifyjs shim.js $(UGLIFYOPTS) -o dist/shim.min.js --preamble "$$(head -n 1 bits/00_header.js)"
	@#
	<$(TARGET) sed "s/require('.*')/undefined/g;s/ process / undefined /g;s/process.versions/({})/g" > dist/$(TARGET)
	<$(MINITGT) sed "s/require('.*')/undefined/g;s/ process / undefined /g;s/process.versions/({})/g" > dist/$(MINITGT)
	@# core
	uglifyjs $(REQS) dist/$(TARGET) $(UGLIFYOPTS) -o dist/$(LIB).core.min.js --source-map dist/$(LIB).core.min.map --preamble "$$(head -n 1 bits/00_header.js)"
	misc/strip_sourcemap.sh dist/$(LIB).core.min.js
	@# full
	#cat <(head -n 1 bits/00_header.js) $(DISTHDR) $(REQS) $(ADDONS) dist/$(TARGET) $(AUXTARGETS) > dist/$(LIB).full.js
	uglifyjs $(DISTHDR) $(REQS) $(ADDONS) dist/$(TARGET) $(AUXTARGETS) $(UGLIFYOPTS) -o dist/$(LIB).full.min.js --source-map dist/$(LIB).full.min.map --preamble "$$(head -n 1 bits/00_header.js)"
	misc/strip_sourcemap.sh dist/$(LIB).full.min.js
	@# mini
	uglifyjs dist/$(MINITGT) $(UGLIFYOPTS) -o dist/$(LIB).mini.min.js --source-map dist/$(LIB).mini.min.map --preamble "$$(head -n 1 bits/00_header.js)"
	misc/strip_sourcemap.sh dist/$(LIB).mini.min.js
	@# extendscript
	cat <(printf '\xEF\xBB\xBF') <(head -n 1 bits/00_header.js) shim.js $(DISTHDR) $(REQS) dist/$(TARGET) > dist/$(LIB).extendscript.js
	@# zahl
	cp modules/xlsx.zahl.js modules/xlsx.zahl.mjs dist/
	@#
	rm dist/$(TARGET) dist/$(MINITGT)

.PHONY: dist-deps
dist-deps: ## Copy dependencies for distribution
	mkdir -p dist
	cp node_modules/codepage/dist/cpexcel.full.js dist/cpexcel.js

.PHONY: aux
aux: $(AUXTARGETS)

BYTEFILEC=dist/xlsx.{full,core,mini}.min.js
BYTEFILER=dist/xlsx.extendscript.js xlsx.mjs
.PHONY: bytes
bytes: ## Display minified and gzipped file sizes
	@for i in $(BYTEFILEC); do npx printj "%-30s %7d %10d" $$i $$(wc -c < $$i) $$(gzip --best --stdout $$i | wc -c); done
	@for i in $(BYTEFILER); do npx printj "%-30s %7d" $$i $$(wc -c < $$i); done


.PHONY: git
git: ## show version string
	@echo "$$(node -pe 'require("./package.json").version')"

.PHONY: nexe
nexe: xlsx.exe ## Build nexe standalone executable

xlsx.exe: bin/xlsx.njs xlsx.js
	tail -n+2 $< | sed 's#\.\./#./xlsx#g' > nexe.js
	nexe -i nexe.js -o $@
	rm nexe.js

.PHONY: pkg
pkg: bin/xlsx.njs xlsx.js ## Build pkg standalone executable
	pkg $<

## Testing

.PHONY: test mocha
test mocha: test.js ## Run test suite
	mocha -R spec -t 30000

#*                      To run tests for one format, make test_<fmt>
#*                      To run the core test suite, make test_misc

.PHONY: test-esm
test-esm: test.mjs ## Run Node ESM test suite
	npx -y mocha@9 -R spec -t 30000 $<

test.ts: test.mts
	node -pe 'var data = fs.readFileSync("'$<'", "utf8"); data.split("\n").map(function(l) { return l.replace(/^describe\((.*?)function\(\)/, "Deno.test($$1async function(t)").replace(/\b(?:it|describe)\((.*?)function\(\)/g, "await t.step($$1async function(t)").replace("assert.ok", "assert.assert"); }).join("\n")' > $@

.PHONY: test-bun
test-bun: testbun.mjs ## Run Bun test suite
	bun $<

.PHONY: test-deno
test-deno: test.ts ## Run Deno test suite
	deno test --allow-env --allow-read --allow-write --config misc/test.deno.jsonc $<

.PHONY: test-denocp
test-denocp: testnocp.ts ## Run Deno test suite (without codepage)
	deno test --allow-env --allow-read --allow-write --config misc/test.deno.jsonc $<

TESTFMT=$(patsubst %,test_%,$(FMT))
.PHONY: $(TESTFMT)
$(TESTFMT): test_%:
	FMTS=$* make test

TESTESMFMT=$(patsubst %,test-esm_%,$(FMT))
.PHONY: $(TESTESMFMT)
$(TESTESMFMT): test-esm_%:
	FMTS=$* make test-esm

TESTDENOFMT=$(patsubst %,test-deno_%,$(FMT))
.PHONY: $(TESTDENOFMT)
$(TESTDENOFMT): test-deno_%:
	FMTS=$* make test-deno

TESTDENOCPFMT=$(patsubst %,test-denocp_%,$(FMT))
.PHONY: $(TESTDENOCPFMT)
$(TESTDENOCPFMT): test-denocp_%:
	FMTS=$* make test-denocp

TESTBUNFMT=$(patsubst %,test-bun_%,$(FMT))
.PHONY: $(TESTBUNFMT)
$(TESTBUNFMT): test-bun_%:
	FMTS=$* make test-bun

.PHONY: travis
travis: ## Run test suite with minimal output
	mocha -R dot -t 30000

.PHONY: ctest
ctest: ## Build browser test fixtures
	node tests/make_fixtures.js

.PHONY: ctestserv
ctestserv: ## Start a test server on port 8000
	@cd tests && python -mSimpleHTTPServer || python3 -mhttp.server || npx -y http-server -p 8000 .

## Code Checking

.PHONY: fullint
fullint: lint mdlint ## Run all checks (removed: old-lint, tslint, flow)

.PHONY: lint
lint: $(TARGET) $(AUXTARGETS) ## Run eslint checks
	@./node_modules/.bin/eslint --ext .js,.njs,.json,.html,.htm $(FLOWTARGET) $(AUXTARGETS) $(CMDS) $(HTMLLINT) package.json bower.json
	@if [ -x "$(CLOSURE)" ]; then java -jar $(CLOSURE) $(REQS) $(FLOWTARGET) --jscomp_warning=reportUnknownTypes >/dev/null; fi

.PHONY: old-lint
old-lint: $(TARGET) $(AUXTARGETS) ## Run jshint and jscs checks
	@./node_modules/.bin/jscs $(TARGET) $(AUXTARGETS) test.js
	@./node_modules/.bin/jshint --show-non-errors $(TARGET) $(AUXTARGETS)
	@./node_modules/.bin/jshint --show-non-errors $(CMDS)
	@./node_modules/.bin/jshint --show-non-errors package.json bower.json test.js
	@./node_modules/.bin/jshint --show-non-errors --extract=always $(HTMLLINT)
	@if [ -x "$(CLOSURE)" ]; then java -jar $(CLOSURE) $(REQS) $(FLOWTARGET) --jscomp_warning=reportUnknownTypes >/dev/null; fi

.PHONY: tslint
tslint: $(TARGET) ## Run typescript checks
	#@npm install dtslint typescript
	#@npm run-script dtslint
	./node_modules/.bin/dtslint types

.PHONY: flow
flow: lint ## Run flow checker
	@./node_modules/.bin/flow check --all --show-all-errors --include-warnings

.PHONY: mjslint
mjslint: $(ESMJSTGT) ## Lint the ESM build
	@npx eslint -c .eslintmjs $<

.PHONY: cov
cov: misc/coverage.html ## Run coverage test

#*                      To run coverage tests for one format, make cov_<fmt>
COVFMT=$(patsubst %,cov_%,$(FMT))
.PHONY: $(COVFMT)
$(COVFMT): cov_%:
	FMTS=$* make cov

misc/coverage.html: $(TARGET) test.js
	mocha --require blanket -R html-cov -t 30000 > $@

.PHONY: coveralls
coveralls: ## Coverage Test + Send to coveralls.io
	mocha --require blanket --reporter mocha-lcov-reporter -t 30000 | node ./node_modules/coveralls/bin/coveralls.js

DEMOMDS=$(sort $(wildcard demos/*/README.md))
MDLINT=$(DEMOMDS) README.md demos/README.md
.PHONY: mdlint
mdlint: $(MDLINT) ## Check markdown documents
	./node_modules/.bin/alex $^
	./node_modules/.bin/mdspell -a -n -x -r --en-us $^

.PHONY: help
help:
	@grep -hE '(^[a-zA-Z_-][ a-zA-Z_-]*:.*?|^#[#*])' $(MAKEFILE_LIST) | bash misc/help.sh

#* To show a spinner, append "-spin" to any target e.g. cov-spin
%-spin:
	@make $* & bash misc/spin.sh $$!
