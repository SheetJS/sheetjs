.PHONY: angular
angular:
	# Test Angular2 build
	cp package.json-angular2 package.json
	npm install
	if [ ! -e node_modules ]; then mkdir node_modules; fi
	if [ ! -e node_modules/xlsx ]; then cd node_modules; ln -s ../../../ xlsx; cd -; fi
	ng build

	# Test Angular4 build
	cp package.json-angular4 package.json
	npm install
	if [ ! -e node_modules ]; then mkdir node_modules; fi
	if [ ! -e node_modules/xlsx ]; then cd node_modules; ln -s ../../../ xlsx; cd -; fi
	ng build

	# Test Angular5 build
	cp package.json-angular5 package.json
	npm install
	if [ ! -e node_modules ]; then mkdir node_modules; fi
	if [ ! -e node_modules/xlsx ]; then cd node_modules; ln -s ../../../ xlsx; cd -; fi
	ng build

