#!/bin/bash
if [ ! -e SheetJSIonic ]; then
	ionic start SheetJSIonic blank --type angular --cordova --quiet --no-git --no-link --confirm </dev/null
	cd SheetJSIonic
	ionic cordova platform add browser --confirm </dev/null
	ionic cordova platform add ios --confirm </dev/null
	ionic cordova platform add android --confirm </dev/null
	ionic cordova plugin add cordova-plugin-file </dev/null
	npm install --save @ionic-native/core
	npm install --save @ionic-native/file
	npm install --save xlsx
	cp ../ionic-app.module.ts src/app/app.module.ts
	cd -
fi

cp ionic.ts SheetJSIonic/src/app/home/home.page.ts
