TOOL=rollup
.PHONY: all
all: $(TOOL).min.js worker.min.js

$(TOOL).min.js: $(TOOL).js
	uglifyjs $< > $@

.PHONY: $(TOOL).js
$(TOOL).js:
	# node
	rollup -c rollup.config.node.js
	node -e 'require("./rollup.node")'
	# browser
	rollup -c
	rollup -c rollup.config.worker.js

worker.min.js: worker.js
	uglifyjs $< > $@

.PHONY: init
init:
	@npm install rollup-plugin-node-resolve rollup-plugin-commonjs
	@mkdir -p node_modules; cd node_modules; ln -s ../../../ xlsx; cd -
