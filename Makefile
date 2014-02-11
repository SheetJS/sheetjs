.PHONY: test ssf
ssf: ssf.md
	voc ssf.md

test:
	npm test

test_min:
	MINTEST=1 npm test

.PHONY: lint
lint:
	jshint ssf.js test/

.PHONY: cov
cov: tmp/coverage.html

tmp/coverage.html: ssf
	mocha --require blanket -R html-cov > tmp/coverage.html

.PHONY: cov_min
cov_min:
	MINTEST=1 make cov

.PHONY: coveralls
coveralls:
	mocha --require blanket --reporter mocha-lcov-reporter | ./node_modules/coveralls/bin/coveralls.js
