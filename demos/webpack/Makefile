TOOL=webpack
WPOPTS=--display-modules --display-reasons --profile
.PHONY: all
all: $(TOOL).min.js core.out.js full.out.js

$(TOOL).min.js: $(TOOL).js
	uglifyjs $< > $@

.PHONY: $(TOOL).js
$(TOOL).js:
	webpack main.js --output-filename $@ $(WPOPTS)

.PHONY: core.out.js full.out.js
core.out.js full.out.js: %.out.js: %.js
	webpack $< --output-filename $@ $(WPOPTS)
