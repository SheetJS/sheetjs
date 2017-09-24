.PHONY: init
init:
	if [ ! -e sheetjs.xlsx ]; then ln -s ../../sheetjs.xlsx; fi
	mkdir -p node_modules
	cd node_modules; if [ ! -e xlsx ]; then ln -s ../../../ xlsx; fi; cd -

.PHONY: request
request: init ## request demo
	node _request.js

.PHONY: express
express: init ## express demo
	node express.js

.PHONY: micro
micro: init ## micro demo
	micro -p 7262 micro.js

.PHONY: koa
koa: init ## koa demo
	node koa.js

.PHONY: hapi
hapi: init ## hapi demo
	cp ../../dist/xlsx.full.min.js .
	node hapi.js
