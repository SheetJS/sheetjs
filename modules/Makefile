LIBFILES=src/types.ts
#LIBFILES=$(wildcard src/*.ts)
TSFILES=$(wildcard *.ts)
ENTRIES=$(subst .ts,.js,$(TSFILES))

BAREJS=04_base64.js 51_xlsxmeta.js 51_xlsbmeta.js 59_vba.js 64_ftab.js 83_numbers.js

.PHONY: all
all: $(ENTRIES) xlsx.zahl.js

xlsx.zahl.js: test.numbers reframe.node.js
	bash -c ./reframe.sh

$(BAREJS): %.js: %.ts $(LIBFILES)
	npx esbuild $< --outfile=$@ --platform=browser --target=es5

%.node.js: %.node.ts $(LIBFILES) src/numbers.ts
	npx esbuild $< --bundle --external:xlsx --outfile=$@ --platform=node
	sed -i '' 's/ts-node/node/g' $@

%.js: %.ts $(LIBFILES)
	npx esbuild $< --bundle --outfile=$@ --platform=browser --format=iife --global-name=$* --target=es5

src/numbers.ts: 83_numbers.ts
	cat $< | sed 's/^\/\/<<//g' > $@

.PHONY: clean
clean:
	rm $(ENTRIES)
