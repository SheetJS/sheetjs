LIB=xlsx
DEPS=$(wildcard bits/*.js)
TARGET=$(LIB).js
FMT=xlsx xlsm xlsb

$(TARGET): $(DEPS)
	cat $^ > $@

bits/31_version.js: package.json
	echo "XLSX.version = '"`grep version package.json | awk '{gsub(/[^0-9\.]/,"",$$2); print $$2}'`"';" > $@

.PHONY: clean
clean:
	rm $(TARGET)

.PHONY: init
init:
	git submodule init
	git submodule update
	git submodule foreach git pull origin master
	git submodule foreach make


.PHONY: test mocha
test mocha: test.js
	mocha -R spec

TESTFMT=$(patsubst %,test_%,$(FMT))
.PHONY: $(TESTFMT)
$(TESTFMT): test_%:
	FMTS=$* make test

.PHONY: jasmine
jasmine:
	npm run-script test-jasmine

.PHONY: lint
lint: $(TARGET)
	jshint --show-non-errors $(TARGET)

.PHONY: cov
cov: misc/coverage.html

misc/coverage.html: $(TARGET) test.js
	mocha --require blanket -R html-cov > misc/coverage.html

.PHONY: coveralls
coveralls:
	mocha --require blanket --reporter mocha-lcov-reporter | ./node_modules/coveralls/bin/coveralls.js

.PHONY: dist
dist: $(TARGET)
	cp $(TARGET) dist/
	cp LICENSE dist/
	uglifyjs $(TARGET) -o dist/$(LIB).min.js --source-map dist/$(LIB).min.map --preamble "$$(head -n 1 bits/00_header.js)"
