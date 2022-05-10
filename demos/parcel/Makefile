.PHONY: app
app: init
	npx parcel build index.html

.PHONY: init
init:
	mkdir -p node_modules
	npm i

.PHONY: ctest 
ctest: init
	npx parcel index.html
