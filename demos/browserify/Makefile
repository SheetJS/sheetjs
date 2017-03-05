TOOL=browserify
.PHONY: all
all: $(TOOL).min.js

$(TOOL).min.js: $(TOOL).js
	uglifyjs $< > $@

.PHONY: $(TOOL).js
$(TOOL).js:
	browserify -r './main.js:xlsx' > $@
