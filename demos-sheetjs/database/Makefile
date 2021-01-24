.PHONY: init
init:
	rm -f node_modules/xlsx
	mkdir -p node_modules
	cd node_modules; ln -s ../../../ xlsx; cd -
	rm -f xlsx.full.min.js
	ln -s ../../dist/xlsx.full.min.js

FILES=$(filter-out xlsx.full.min.js,$(wildcard *.js)) $(wildcard *.html)
.PHONY: lint
lint: $(FILES)
	eslint $(FILES)

.PHONY: clean
clean:
	rm -f *.db *.xlsx *.csv
