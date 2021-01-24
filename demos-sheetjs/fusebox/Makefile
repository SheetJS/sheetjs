.PHONY: all
all server.js client.js: fuse.js
	node fuse.js

.PHONY: test ctest
test: server.js
	@node server.js
	xlsx --dev sheetjsfbox.xlsb

ctest: client.js
	python -mSimpleHTTPServer

.PHONY: init
init:
	mkdir -p node_modules
	cd node_modules; ln -s ../../../ xlsx; cd -

