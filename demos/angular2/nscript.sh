#!/bin/bash
if [ ! -e SheetJSNS ]; then
	ns create SheetJSNS --ng
	cd SheetJSNS
	npm install xlsx
	cd ..
fi

<../../dist/xlsx.full.min.js sed 's/require("fs")/null/g' > SheetJSNS/src/app/xlsx.full.min.js
cp nscript.ts SheetJSNS/src/app/app.component.ts
