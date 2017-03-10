SHELL=/bin/bash
LIB=xlsx
FMT=xlsx xlsm xlsb ods xls xml misc full
REQS=jszip.js
ADDONS=dist/cpexcel.js
AUXTARGETS=
CMDS=bin/xlsx.njs
HTMLLINT=index.html

ULIB=$(shell echo $(LIB) | tr a-z A-Z)
DEPS=$(sort $(wildcard bits/*.js))
TARGET=$(LIB).js
FLOWTARGET=$(LIB).flow.js
FLOWAUX=$(patsubst %.js,%.flow.js,$(AUXTARGETS))
AUXSCPTS=xlsxworker1.js xlsxworker2.js xlsxworker.js
FLOWTGTS=$(TARGET) $(AUXTARGETS) $(AUXSCPTS)
UGLIFYOPTS=--support-ie8

## Main Targets

.PHONY: all
all: $(TARGET) $(AUXTARGETS) $(AUXSCPTS) ## Build library and auxiliary scripts

$(FLOWTGTS): %.js : %.flow.js
	node -e 'process.stdout.write(require("fs").readFileSync("$<","utf8").replace(/^[ \t]*\/\*[:#][^*]*\*\/\s*(\n)?/gm,"").replace(/\/\*[:#][^*]*\*\//gm,""))' > $@

$(FLOWTARGET): $(DEPS)
	cat $^ | tr -d '\15\32' > $@

bits/01_version.js: package.json
	echo "$(ULIB).version = '"`grep version package.json | awk '{gsub(/[^0-9a-z\.-]/,"",$$2); print $$2}'`"';" > $@

bits/18_cfb.js: node_modules/cfb/xlscfb.flow.js
	cp $^ $@

.PHONY: clean
clean: ## Remove targets and build artifacts
	rm -f $(TARGET) $(FLOWTARGET)

.PHONY: clean-data
clean-data:
	rm -f *.xlsx *.xlsm *.xlsb *.xls *.xml

.PHONY: init
init: ## Initial setup for development
	git submodule init
	git submodule update
	git submodule foreach git pull origin master
	git submodule foreach make
	mkdir -p tmp

.PHONY: dist
dist: dist-deps $(TARGET) bower.json ## Prepare JS files for distribution
	cp $(TARGET) dist/
	cp LICENSE dist/
	uglifyjs $(UGLIFYOPTS) $(TARGET) -o dist/$(LIB).min.js --source-map dist/$(LIB).min.map --preamble "$$(head -n 1 bits/00_header.js)"
	misc/strip_sourcemap.sh dist/$(LIB).min.js
	uglifyjs $(UGLIFYOPTS) $(REQS) $(TARGET) -o dist/$(LIB).core.min.js --source-map dist/$(LIB).core.min.map --preamble "$$(head -n 1 bits/00_header.js)"
	misc/strip_sourcemap.sh dist/$(LIB).core.min.js
	uglifyjs $(UGLIFYOPTS) $(REQS) $(ADDONS) $(TARGET) $(AUXTARGETS) -o dist/$(LIB).full.min.js --source-map dist/$(LIB).full.min.map --preamble "$$(head -n 1 bits/00_header.js)"
	misc/strip_sourcemap.sh dist/$(LIB).full.min.js
	cat <(head -n 1 bits/00_header.js) $(REQS) $(ADDONS) $(TARGET) $(AUXTARGETS) > demos/requirejs/$(LIB).full.js

.PHONY: dist-deps
dist-deps: ## Copy dependencies for distribution
	cp node_modules/codepage/dist/cpexcel.full.js dist/cpexcel.js
	cp jszip.js dist/jszip.js

.PHONY: aux
aux: $(AUXTARGETS)

.PHONY: nexe
nexe: xlsx.exe

xlsx.exe: bin/xlsx.js xlsx.js
	nexe -i $< -o $@ --flags

## Testing

.PHONY: test mocha
test mocha: test.js ## Run test suite
	mocha -R spec -t 20000

#*                      To run tests for one format, make test_<fmt>
TESTFMT=$(patsubst %,test_%,$(FMT))
.PHONY: $(TESTFMT)
$(TESTFMT): test_%:
	FMTS=$* make test

.PHONY: demos
demos: demo-browserify demo-webpack demo-requirejs

.PHONY: demo-browserify
demo-browserify: ## Run browserify demo build
	make -C demos/browserify
	@echo "start a local server and go to demos/browserify/browserify.html"

.PHONY: demo-webpack
demo-webpack: ## Run webpack demo build
	make -C demos/webpack
	@echo "start a local server and go to demos/webpack/webpack.html"

.PHONY: demo-requirejs
demo-requirejs: ## Run requirejs demo build
	make -C demos/requirejs
	@echo "start a local server and go to demos/requirejs/requirejs.html"

## Code Checking

.PHONY: lint
lint: $(TARGET) $(AUXTARGETS) ## Run jshint and jscs checks
	@jshint --show-non-errors $(TARGET) $(AUXTARGETS)
	@jshint --show-non-errors $(CMDS)
	@jshint --show-non-errors package.json bower.json
	@jshint --show-non-errors --extract=always $(HTMLLINT)
	@jscs $(TARGET) $(AUXTARGETS)

.PHONY: flow
flow: lint ## Run flow checker
	@flow check --all --show-all-errors

.PHONY: cov
cov: misc/coverage.html ## Run coverage test

#*                      To run coverage tests for one format, make cov_<fmt>
COVFMT=$(patsubst %,cov_%,$(FMT))
.PHONY: $(COVFMT)
$(COVFMT): cov_%:
	FMTS=$* make cov

misc/coverage.html: $(TARGET) test.js
	mocha --require blanket -R html-cov -t 20000 > $@

.PHONY: coveralls
coveralls: ## Coverage Test + Send to coveralls.io
	mocha --require blanket --reporter mocha-lcov-reporter -t 20000 | node ./node_modules/coveralls/bin/coveralls.js


.PHONY: help
help:
	@grep -hE '(^[a-zA-Z_-][ a-zA-Z_-]*:.*?|^#[#*])' $(MAKEFILE_LIST) | bash misc/help.sh

#* To show a spinner, append "-spin" to any target e.g. cov-spin
%-spin:
	@make $* & bash misc/spin.sh $$!
