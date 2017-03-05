TOOL=requirejs
.PHONY: all
all: $(TOOL).js

.PHONY: $(TOOL).js
$(TOOL).js:
	if [ ! -e require.js ]; then curl -O http://requirejs.org/docs/release/2.3.3/comments/require.js; fi
	if [ ! -e r.js ]; then curl -O http://requirejs.org/docs/release/2.3.3/r.js; fi
	node r.js -o build.js
