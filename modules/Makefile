LIBFILES=$(wildcard src/*.ts)
TSFILES=$(wildcard *.ts)
ENTRIES=$(subst .ts,.js,$(TSFILES))

BAREJS=04_base64.js 51_xlsxmeta.js 51_xlsbmeta.js 59_vba.js 64_ftab.js

.PHONY: all
all: $(ENTRIES) 51_xlmeta.js

51_xlmeta.js: 51_xlsxmeta.js 51_xlsbmeta.js
	cat $^ > $@

$(BAREJS): %.js: %.ts $(LIBFILES)
	npx esbuild $< --outfile=$@ --platform=browser --target=es5

83_numbers.js: 83_numbers.ts $(LIBFILES)
	npx esbuild $< --bundle --outfile=$@ --platform=browser --format=iife --global-name=NUMBERS --target=es5
	sed -i '' 's/NUMBERS = /NUMBERS = !Object.defineProperty ? (void 0) :/g' $@

%.node.js: %.node.ts $(LIBFILES)
	npx esbuild $< --bundle --external:xlsx --outfile=$@ --minify --platform=node
	sed -i '' 's/ts-node/node/g' $@

%.js: %.ts $(LIBFILES)
	npx esbuild $< --bundle --outfile=$@ --platform=browser --format=iife --global-name=$* --target=es5

.PHONY: clean
clean:
	rm $(ENTRIES)
