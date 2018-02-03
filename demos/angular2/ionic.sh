#!/bin/bash
if [ ! -e SheetJSIonic ]; then
	ionic start SheetJSIonic blank --cordova --no-git --no-link </dev/null
	cd SheetJSIonic
	ionic cordova platform add browser </dev/null
	ionic cordova platform add ios </dev/null
	ionic cordova plugin add cordova-plugin-file </dev/null
	npm install --save @ionic-native/file
	npm install --save xlsx

	cp src/app/app.module.ts{,.bak}
	cat src/app/app.module.ts.bak | awk 'BEGIN{p=0} !/import/ && !p { ++p; print "import { File } from '"'"'@ionic-native/file'"'"';"; } 1; /providers: \[/ {print "    File,"}' > src/app/app.module.ts
	cd -
fi

cp ionic.ts SheetJSIonic/src/pages/home/home.ts
rm -f SheetJSIonic/src/pages/home/home.html
