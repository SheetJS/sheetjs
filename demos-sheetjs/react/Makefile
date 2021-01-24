.PHONY: react
react: init ## Simple server for react and clones
	python -mSimpleHTTPServer

.PHONY: next
next: init ## next.js demo
	mkdir -p pages static
	cat nexthdr.js sheetjs.jsx > pages/sheetjs.js
	cp ../../shim.js static/shim.js
	next

.PHONY: native
native: ## Build react-native project
	bash ./native.sh

.PHONY: ios
ios: native ## react-native ios sim
	cd SheetJS; react-native run-ios --simulator="iPhone X"; cd -

.PHONY: android
android: native ## react-native android sim
	cd SheetJS; react-native run-android; cd -

.PHONY: init
init: ## set up node_modules and symlink
	mkdir -p node_modules
	cd node_modules; if [ ! -e xlsx ]; then ln -s ../../../ xlsx; fi; cd -
