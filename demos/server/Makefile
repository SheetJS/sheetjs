.PHONY: init
init:
	mkdir -p node_modules
	cd node_modules; if [ ! -e xlsx ]; then ln -s ../../../ xlsx; fi; cd -

.PHONY: request
request: init ## request demo
	node _request.js

.PHONY: express
express: init ## express demo
	node express.js

.PHONY: koa
koa: init ## koa demo
	node koa.js

.PHONY: hapi
hapi: init ## hapi demo
	cp ../../dist/xlsx.full.min.js .
	node hapi.js

.PHONY: nest
nest: init ## nest demo
	bash -c ./nest.sh

.PHONY: drash
drash: ## drash demo
	deno run --allow-net drash.ts
