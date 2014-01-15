DEPS=$(wildcard bits/*.js)
TARGET=xlsx.js
$(TARGET): $(DEPS)
	cat $^ > $@

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
