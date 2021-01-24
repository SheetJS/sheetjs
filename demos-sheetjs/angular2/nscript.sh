#!/bin/bash
if [ ! -e SheetJSNS ]; then
	tns create SheetJSNS --template nativescript-template-ng-tutorial
	cd SheetJSNS
	tns plugin add nativescript-nodeify
	npm install xlsx
	cd app
	npm install xlsx
	cd ../..
fi

cp ../../dist/xlsx.full.min.js SheetJSNS/
cp ../../dist/xlsx.full.min.js SheetJSNS/app/
cp nsmain.ts  SheetJSNS/app/main.ts
cp nscript.ts SheetJSNS/app/app.component.ts
