.PHONY: start
start:
	@meteor

.PHONY: init
init:
	if [ ! -e .meteor ]; then meteor create .; fi;
	@npm install babel-runtime meteor-node-stubs
	@meteor add pfafman:filesaver check
	@mkdir -p node_modules; cd node_modules; ln -s ../../../ xlsx; cd -

.PHONY: lint
lint:
	@meteor npm run lint