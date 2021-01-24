.PHONY: all
all:
	@npm run build

.PHONY: lint
lint:
	@npm run lint

.PHONY: init
init:
	mkdir -p node_modules
	npm install typescript
	cd node_modules; ln -s ../../../ xlsx; cd -
