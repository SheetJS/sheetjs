LIB=xlsx
DEPS=$(wildcard bits/*.js)
TARGET=$(LIB).js
FMT=xlsx xlsm xlsb misc full
REQS=jszip.js
ADDONS=dist/cpexcel.js

$(TARGET): $(DEPS)
	cat $^ > $@

bits/01_version.js: package.json
	echo "XLSX.version = '"`grep version package.json | awk '{gsub(/[^0-9a-z\.-]/,"",$$2); print $$2}'`"';" > $@

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
	mkdir -p tmp
	mocha -R spec

TESTFMT=$(patsubst %,test_%,$(FMT))
.PHONY: $(TESTFMT)
$(TESTFMT): test_%:
	FMTS=$* make test


.PHONY: lint
lint: $(TARGET)
	jshint --show-non-errors $(TARGET)

.PHONY: cov cov-spin
cov: misc/coverage.html
cov-spin:
	make cov & bash misc/spin.sh $$!

COVFMT=$(patsubst %,cov_%,$(FMT))
.PHONY: $(COVFMT)
$(COVFMT): cov_%:
	FMTS=$* make cov

misc/coverage.html: $(TARGET) test.js
	mocha --require blanket -R html-cov > $@

.PHONY: coveralls coveralls-spin
coveralls:
	mocha --require blanket --reporter mocha-lcov-reporter | ./node_modules/coveralls/bin/coveralls.js

coveralls-spin:
	make coveralls & bash misc/spin.sh $$!

.PHONY: dist
dist: dist-deps $(TARGET)
	cp $(TARGET) dist/
	cp LICENSE dist/
	uglifyjs $(TARGET) -o dist/$(LIB).min.js --source-map dist/$(LIB).min.map --preamble "$$(head -n 1 bits/00_header.js)"
	uglifyjs $(REQS) $(TARGET) -o dist/$(LIB).core.min.js --source-map dist/$(LIB).core.min.map --preamble "$$(head -n 1 bits/00_header.js)"
	uglifyjs $(REQS) $(ADDONS) $(TARGET) -o dist/$(LIB).full.min.js --source-map dist/$(LIB).full.min.map --preamble "$$(head -n 1 bits/00_header.js)"

.PHONY: dist-deps
dist-deps:
	cp node_modules/codepage/dist/cpexcel.full.js dist/cpexcel.js
	cp jszip.js dist/jszip.js
