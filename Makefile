LIB=xlsx
FMT=xlsx xlsm xlsb ods misc full
REQS=jszip.js
ADDONS=dist/cpexcel.js
AUXTARGETS=ods.js

ULIB=$(shell echo $(LIB) | tr a-z A-Z)
DEPS=$(sort $(wildcard bits/*.js))
TARGET=$(LIB).js

.PHONY: all
all: $(TARGET) $(AUXTARGETS)

$(TARGET): $(DEPS)
	cat $^ | tr -d '\15\32' > $@

bits/01_version.js: package.json
	echo "$(ULIB).version = '"`grep version package.json | awk '{gsub(/[^0-9a-z\.-]/,"",$$2); print $$2}'`"';" > $@

.PHONY: clean
clean:
	rm -f $(TARGET)

.PHONY: clean-data
clean-data:
	rm -f *.xlsx *.xlsm *.xlsb *.xls *.xml

.PHONY: init
init:
	git submodule init
	git submodule update
	git submodule foreach git pull origin master
	git submodule foreach make


.PHONY: test mocha
test mocha: test.js
	mkdir -p tmp
	mocha -R spec -t 20000

.PHONY: prof
prof:
	cat misc/prof.js test.js > prof.js
	node --prof prof.js

TESTFMT=$(patsubst %,test_%,$(FMT))
.PHONY: $(TESTFMT)
$(TESTFMT): test_%:
	FMTS=$* make test


.PHONY: lint
lint: $(TARGET)
	jshint --show-non-errors $(TARGET) $(AUXTARGETS)
	jscs $(TARGET) $(AUXTARGETS)

.PHONY: test-osx
test-osx:
	node tests/write.js
	open -a Numbers sheetjs.xlsx
	open -a "Microsoft Excel" sheetjs.xlsx

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

bower.json: misc/_bower.json package.json
	cat $< | sed 's/_VERSION_/'`grep version package.json | awk '{gsub(/[^0-9a-z\.-]/,"",$$2); print $$2}'`'/' > $@

.PHONY: dist
dist: dist-deps $(TARGET) bower.json
	cp $(TARGET) dist/
	cp LICENSE dist/
	uglifyjs $(TARGET) -o dist/$(LIB).min.js --source-map dist/$(LIB).min.map --preamble "$$(head -n 1 bits/00_header.js)"
	uglifyjs $(REQS) $(TARGET) -o dist/$(LIB).core.min.js --source-map dist/$(LIB).core.min.map --preamble "$$(head -n 1 bits/00_header.js)"
	uglifyjs $(REQS) $(ADDONS) $(TARGET) -o dist/$(LIB).full.min.js --source-map dist/$(LIB).full.min.map --preamble "$$(head -n 1 bits/00_header.js)"

.PHONY: aux
aux: $(AUXTARGETS)

.PHONY: ods
ods: ods.js

ODSDEPS=$(sort $(wildcard odsbits/*.js))
ods.js: $(ODSDEPS)
	cat $(ODSDEPS) | tr -d '\15\32' > $@
	cp ods.js dist/ods.js

.PHONY: dist-deps
dist-deps: ods.js
	cp node_modules/codepage/dist/cpexcel.full.js dist/cpexcel.js
	cp jszip.js dist/jszip.js
	cp ods.js dist/ods.js
