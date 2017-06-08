TOOL=rollup
.PHONY: all
all: $(TOOL).min.js

$(TOOL).min.js: $(TOOL).js
	uglifyjs $< > $@

.PHONY: $(TOOL).js
$(TOOL).js:
	# node
	rollup -c rollup.config.node.js
	node -e 'require("./rollup.node")'
	# browser
	rollup -c

.PHONY: init
init:
	@npm install rollup-plugin-node-resolve rollup-plugin-commonjs
	@mkdir -p node_modules; cd node_modules; ln -s ../../../ xlsx; cd -
