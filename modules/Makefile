LIBFILES=src/types.ts
#LIBFILES=$(wildcard src/*.ts)
TSFILES=$(wildcard *.ts)
ENTRIES=$(subst .ts,.js,$(TSFILES))
CC=esbuild@0.14.14

BAREJS=04_base64.js 45_rtf.js 51_xlsxmeta.js 51_xlsbmeta.js 59_vba.js 64_ftab.js 83_numbers.js

.PHONY: all
all: bits xlsx.zahl.js

.PHONY: bits
bits: $(ENTRIES)

xlsx.zahl.js: test.numbers
	bash -c ./reframe.sh

$(BAREJS): %.js: %.ts $(LIBFILES)
	npx $(CC) $< --outfile=$@ --platform=browser --target=es5

%.node.js: %.node.ts $(LIBFILES) src/numbers.ts
	npx $(CC) $< --bundle --external:xlsx --outfile=$@ --platform=node
	sed -i '' 's/ts-node/node/g' $@ || sed -i'' 's/ts-node/node/g' $@ || { echo "sed failed"; exit 1; }

%.js: %.ts $(LIBFILES)
	npx $(CC) $< --bundle --outfile=$@ --platform=browser --format=iife --global-name=$* --target=es5

src/numbers.ts: 83_numbers.ts
	cat $< | sed 's/^\/\/<<//g' > $@

.PHONY: clean
clean:
	rm $(ENTRIES)
