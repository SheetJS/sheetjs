.PHONY: start
start:
	@meteor

.PHONY: init
init:
	@npm install babel-runtime meteor-node-stubs
	@meteor add pfafman:filesaver
	@mkdir -p node_modules; cd node_modules; ln -s ../../../ xlsx; cd -
