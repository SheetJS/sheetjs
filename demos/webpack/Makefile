TOOL=webpack
WPOPTS=--display-modules --display-reasons --profile
.PHONY: all
all: main.min.js core.out.js full.out.js app.out.js

main.min.js: main.out.js
	uglifyjs $< > $@

.PHONY: main.out.js core.out.js full.out.js
main.out.js core.out.js full.out.js: %.out.js: %.js
	webpack $< --output-filename $@ $(WPOPTS)

.PHONY: app.out.js
app.out.js: webpack.app.js app.js appworker.js
	webpack --config $< $(WPOPTS)
