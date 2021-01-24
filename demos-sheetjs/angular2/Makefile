.PHONY: all
all: angular5

.PHONY: angular2 angular4 angular5
angular2 angular4 angular5:
	cp package.json-$@ package.json
	rm -rf node_modules
	npm install
	if [ ! -e node_modules ]; then mkdir node_modules; fi
	if [ ! -e node_modules/xlsx ]; then cd node_modules; ln -s ../../../ xlsx; cd -; fi
	npm run build

.PHONY: ionic
ionic:
	bash ./ionic.sh

.PHONY: ios android browser
ios android browser: ionic
	cd SheetJSIonic; ionic cordova emulate $@ </dev/null; cd -

.PHONY: nativescript
nativescript:
	bash ./nscript.sh

.PHONY: ns-ios ns-android
ns-ios: nativescript
	cd SheetJSNS; tns run ios; cd -
ns-android: nativescript
	cd SheetJSNS; tns run android; cd -

