.PHONY: test ssf
ssf: ssf.md
	voc ssf.md

test:
	npm test

.PHONY: lint
lint:
	jshint ssf.js test/

.PHONY: cov
cov: tmp/coverage.html

tmp/coverage.html: ssf.md
	mocha --require blanket -R html-cov > tmp/coverage.html

.PHONY: coveralls
coveralls:
	mocha --require blanket --reporter mocha-lcov-reporter | ./node_modules/coveralls/bin/coveralls.js
