LIBFILES=$(wildcard src/*.ts)
TSFILES=$(wildcard *.ts)
ENTRIES=$(subst .ts,.js,$(TSFILES))

.PHONY: all
all: $(ENTRIES)

83_numbers.js: 83_numbers.ts $(LIBFILES)
	npx esbuild $< --bundle --outfile=$@ --platform=browser --format=iife --global-name=NUMBERS --target=es5

%.node.js: %.node.ts $(LIBFILES)
	npx esbuild $< --bundle --external:xlsx --outfile=$@ --minify --platform=node
	sed -i '' 's/ts-node/node/g' $@

%.js: %.ts $(LIBFILES)
	npx esbuild $< --bundle --outfile=$@ --platform=browser --format=iife --global-name=$* --target=es5

.PHONY: clean
clean:
	rm $(ENTRIES)
