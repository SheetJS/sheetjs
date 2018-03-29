.PHONY: init
init:
	mkdir -p node_modules
	cd node_modules; if [ ! -e xlsx ]; then ln -s ../../../ xlsx ; fi; cd -

.PHONY: lint
lint:
	eslint *.js
.PHONY: run
run:
	electron .
