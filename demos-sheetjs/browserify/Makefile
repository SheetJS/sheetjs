TOOL=browserify
.PHONY: all
all: $(TOOL).min.js worker.min.js

$(TOOL).min.js: $(TOOL).js
	uglifyjs $< > $@

.PHONY: $(TOOL).js
$(TOOL).js: app.js
	browserify $< > $@

worker.min.js: worker.js
	uglifyjs $< > $@

.PHONY: worker.js
worker.js: xlsxworker.js
	browserify $< > $@
