SHELL=/bin/bash
LIB=ssf
CMDS=bin/ssf.njs
HTMLLINT=

ULIB=$(shell echo $(LIB) | tr a-z A-Z)
DEPS=$(sort $(wildcard bits/*.js))
TARGET=$(LIB).js
FLOWTARGET=$(LIB).flow.js
FLOWAUX=$(patsubst %.js,%.flow.js,$(AUXTARGETS))
AUXSCPTS=
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

.PHONY: clean
clean: ## Remove targets and build artifacts
	rm -f $(TARGET) $(FLOWTARGET)


## Testing

.PHONY: test mocha
test mocha: ## Run test suite
	mocha -R spec -t 30000

.PHONY: test_min
test_min:
	MINTEST=1 npm test

.PHONY: travis
travis: ## Run test suite with minimal output
	mocha -R dot -t 30000

## Code Checking

.PHONY: lint
lint: $(TARGET) $(AUXTARGETS) ## Run jshint and jscs checks
	@jshint --show-non-errors $(TARGET) $(AUXTARGETS)
	@jshint --show-non-errors test/
	@jshint --show-non-errors $(CMDS)
	@jshint --show-non-errors package.json
	@jshint --show-non-errors --extract=always $(HTMLLINT)
	@jscs $(TARGET) $(AUXTARGETS)

.PHONY: flow
flow: lint ## Run flow checker
	@flow check --all --show-all-errors

.PHONY: cov
cov: tmp/coverage.html ## Run coverage test

.PHONY: cov_min
cov_min:
	MINTEST=1 make cov

misc/coverage.html: $(TARGET)
	mocha --require blanket -R html-cov -t 20000 > $@

.PHONY: full_coveralls
full_coveralls:
	mocha --require blanket --reporter mocha-lcov-reporter | ./node_modules/coveralls/bin/coveralls.js

.PHONY: coveralls
coveralls: ## Coverage Test + Send to coveralls.io
	MINTEST=1 make full_coveralls


.PHONY: help
help:
	@grep -hE '(^[a-zA-Z_-][ a-zA-Z_-]*:.*?|^#[#*])' $(MAKEFILE_LIST) | bash misc/help.sh

#* To show a spinner, append "-spin" to any target e.g. cov-spin
%-spin:
	@make $* & bash misc/spin.sh $$!
