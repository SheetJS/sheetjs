.PHONY: ng2 ng4 ng5 ng6 ng7 ng8 ng9 ng10 ng11 ng12 ng13
ng2 ng4 ng5 ng6 ng7 ng8 ng9 ng10 ng11 ng12 ng13:
	rm -f angular.json tsconfig.app.json src/polyfills.ts
	cp versions/package.json-$@ package.json
	if [ -e versions/angular.json-$@ ]; then cp versions/angular.json-$@ angular.json; fi
	if [ -e versions/tsconfig.app.json-$@ ]; then cp versions/tsconfig.app.json-$@ tsconfig.app.json; fi
	if [ -e versions/polyfills.ts-$@ ]; then cp versions/polyfills.ts-$@ src/polyfills.ts; fi
	rm -rf node_modules
	if [ ! -e node_modules ]; then mkdir node_modules; fi
	npm install
	if [ ! -e node_modules/xlsx ]; then cd node_modules; ln -s ../../../ xlsx; cd -; fi
	npm run build

.PHONY: refresh
refresh: ## refresh the `xlsx` symlink to force angular to rebuild
	rm -rf .angular/
	rm -f node_modules/xlsx
	cd node_modules; ln -s ../../../ xlsx; cd -
	touch node_modules/xlsx

.PHONY: all
all:
	for i in 2 4 5 6 7 8 9 10 11 12 13; do make ng$$i; done

.PHONY: ionic
ionic:
	bash ./ionic.sh

.PHONY: ios android browser
ios browser: ionic
	cd SheetJSIonic; ionic cordova emulate $@ </dev/null; cd -
android: ionic
	cd SheetJSIonic; ionic cordova prepare $@ </dev/null; ionic cordova emulate $@ </dev/null; cd -


.PHONY: nativescript
nativescript:
	bash ./nscript.sh

.PHONY: ns-ios ns-android
ns-ios: nativescript
	cd SheetJSNS; ns run ios; cd -
ns-android: nativescript
	cd SheetJSNS; ns run android; cd -

