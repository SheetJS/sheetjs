.PHONY: all
all: cjs esm mjs S types

.PHONY: cjs esm types
cjs esm types: %: ./tsconfig/tsconfig.%.json ./src/
	@tsc -p ./tsconfig/tsconfig.$@.json

.PHONY: mjs
mjs: esm
	@./misc/make_mjs.sh

.PHONY: S
S: cjs
	@webpack-cli --config misc/webpack.config.js misc/standalone.js --output-filename S.js --display-modules --display-reasons --profile

.PHONY: test
test: mjs cjs S
	@node misc/require_test.njs
	@node --experimental-modules --es-module-specifier-resolution=node misc/import_test.mjs
	@cp dist/S.js misc && node misc/puppetsta.js
	@concurrently --kill-others -s first "http-server -s -p 7262" "node misc/puppetmjs.js"
	@parcel build misc/parcel.html -d misc/parcel -o parcelout.html --public-url "./" && node misc/puppetpcl.js
	@mocha -r ts-node/register test.ts