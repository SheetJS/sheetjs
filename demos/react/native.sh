#!/bin/bash
# xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
if [ ! -e SheetJS ]; then
	react-native init SheetJS
	cd SheetJS
	npm i -S xlsx react react-native react-native-table-component react-native-fs
	cd -
fi
if [ ! -e SheetJS/logo.png ]; then
	curl -O http://oss.sheetjs.com/assets/img/logo.png
	mv logo.png SheetJS/logo.png
fi
cp react-native.js SheetJS/index.ios.js
cp react-native.js SheetJS/index.android.js
cd SheetJS;
RNFB_ANDROID_PERMISSIONS=true react-native link
cd -;
