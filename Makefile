DEPS=$(wildcard bits/*.js)
TARGET=xlsx.js

$(TARGET): $(DEPS)
	cat $^ > $@

bits/51_version.js: package.json
	echo "XLSX.version = '"`grep version package.json | awk '{gsub(/[^0-9\.]/,"",$$2); print $$2}'`"';" > bits/51_version.js

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
test mocha:
	mocha -R spec

.PHONY: jasmine
jasmine:
	npm run-script test-jasmine

.PHONY: lint
lint: $(TARGET)
	jshint --show-non-errors $(TARGET)

.PHONY: cov
cov: misc/coverage.html

misc/coverage.html: xlsx.js 
	mocha --require blanket -R html-cov > misc/coverage.html

.PHONY: coveralls
coveralls:
	mocha --require blanket --reporter mocha-lcov-reporter | ./node_modules/coveralls/bin/coveralls.js
