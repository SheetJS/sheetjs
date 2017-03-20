.PHONY: update
update:
	git show master:dist/cpexcel.js > dist/cpexcel.js
	git show master:dist/xlsx.core.min.js > xlsx.core.min.js
	git show master:dist/xlsx.full.min.js > xlsx.full.min.js
	git show master:xlsx.js > xlsx.js
	git show master:tests/core.js > tests/core.js
	git show master:tests/fixtures.js > tests/fixtures.js
	git show master:tests/fs_.js > tests/fs_.js
	git show master:tests/mocha.css > tests/mocha.css
	git show master:tests/mocha.js > tests/mocha.js
	git show master:tests/xhr-hack.js > tests/xhr-hack.js
