#!/bin/bash
# xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */

# Create starter project
if [ ! -e SheetJS ]; then react-native init SheetJS --version="0.67.2"; fi

# Install dependencies
cd SheetJS; npm i -S xlsx react-native-table-component; cd -

cd SheetJS; npm i -S react-native-file-access@2.x; cd -
# cd SheetJS; npm i -S react-native-fs; cd -
# cd SheetJS; npm i -S react-native-fetch-blob; cd -

# Copy demo assets
if [ ! -e SheetJS/logo.png ]; then
	curl -O http://oss.sheetjs.com/assets/img/logo.png
	mv logo.png SheetJS/logo.png
fi
if [ -e SheetJS/index.ios.js ]; then
	cp react-native.js SheetJS/index.ios.js
	cp react-native.js SheetJS/index.android.js
else
	cp react-native.js SheetJS/index.js
fi

# Link
cd SheetJS; RNFB_ANDROID_PERMISSIONS=true react-native link; cd -
