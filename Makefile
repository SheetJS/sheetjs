SHELL=/bin/bash
LIB=ssf
CMDS=bin/ssf.njs
HTMLLINT=

ULIB=$(shell echo $(LIB) | tr a-z A-Z)
TARGET=$(LIB).js

## Main Targets


.PHONY: ssf
ssf: ssf.md
	voc ssf.md

## Testing

.PHONY: test mocha
test mocha: ## Run test suite
	npm test

test_min:
	MINTEST=1 npm test

## Code Checking

.PHONY: lint
lint: ## Run jshint and jscs checks
	@jshint --show-non-errors $(TARGET) test/
	@jshint --show-non-errors $(CMDS)
	@jshint --show-non-errors package.json
	@jshint --show-non-errors --extract=always $(HTMLLINT)
	@jscs $(TARGET)

.PHONY: flow
flow: lint ## Run flow checker
	@flow check --all --show-all-errors

.PHONY: cov
cov: tmp/coverage.html ## Run coverage test

.PHONY: cov_min
cov_min:
	MINTEST=1 make cov

tmp/coverage.html: ssf
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
