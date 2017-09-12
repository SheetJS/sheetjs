.PHONY: react
react: ## Simple server for react and clones
	python -mSimpleHTTPServer

.PHONY: next
next: ## next.js demo
	# next doesn't support jsx extension
	mkdir -p pages
	cp sheetjs.jsx pages/sheetjs.js
	next

.PHONY: native
native: ## Build react-native project
	bash ./native.sh

.PHONY: ios
ios: native ## react-native ios sim
	cd SheetJS; react-native run-ios; cd -

.PHONY: android
android: native ## react-native android sim
	cd SheetJS; react-native run-android; cd -
