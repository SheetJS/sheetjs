.PHONY: update
update:
	git show master:dist/cpexcel.js > dist/cpexcel.js
	git show master:dist/xlsx.core.min.js > xlsx.core.min.js
	git show master:dist/xlsx.full.min.js > xlsx.full.min.js
	git show master:xlsx.js > xlsx.js	
