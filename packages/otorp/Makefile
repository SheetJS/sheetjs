LIBFILES=$(wildcard src/*.ts)
TSFILES=$(wildcard *.ts)
ENTRIES=$(subst .ts,.js,$(TSFILES))

.PHONY: all
all: $(ENTRIES)

index.node.js: index.node.ts $(LIBFILES)
	npx esbuild@0.14.14 $< --bundle --outfile=$@ --platform=node --format=cjs

%.node.js: %.node.ts $(LIBFILES)
	npx esbuild@0.14.14 $< --bundle --external:./ --outfile=$@ --platform=node
	sed -i '' 's/ts-node/node/g' $@

%.js: %.ts $(LIBFILES)
	npx esbuild@0.14.14 $< --bundle --outfile=$@ --platform=browser --format=iife --global-name=$* --target=es5

.PHONY: clean
clean:
	rm $(ENTRIES)
