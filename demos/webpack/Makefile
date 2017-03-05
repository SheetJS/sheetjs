TOOL=webpack
.PHONY: all
all: $(TOOL).min.js

$(TOOL).min.js: $(TOOL).js
	uglifyjs $< > $@

.PHONY: $(TOOL).js
$(TOOL).js:
	webpack main.js --output-filename $@ --display-modules
