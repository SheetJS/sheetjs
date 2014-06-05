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
	jscs ssf.js

.PHONY: cov
cov: tmp/coverage.html

tmp/coverage.html: ssf
	mocha --require blanket -R html-cov > tmp/coverage.html

.PHONY: cov_min
cov_min:
	MINTEST=1 make cov

.PHONY: coveralls full_coveralls
full_coveralls:
	mocha --require blanket --reporter mocha-lcov-reporter | ./node_modules/coveralls/bin/coveralls.js

coveralls:
	MINTEST=1 make full_coveralls
