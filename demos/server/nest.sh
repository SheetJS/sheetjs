#!/bin/bash

# it is assumed that @nestjs/cli is installed globally

if [ ! -e xlsx-demo ]; then
	nest new -p npm xlsx-demo
fi

cd xlsx-demo
npm i --save xlsx
npm i --save-dev @types/multer

if [ ! -e src/sheetjs/sheetjs.module.ts ]; then
	nest generate module sheetjs
fi

if [ ! -e src/sheetjs/sheetjs.controller.ts ]; then
	nest generate controller sheetjs
fi

cp ../sheetjs.module.ts src/sheetjs/
cp ../sheetjs.controller.ts src/sheetjs/
mkdir -p upload
npm run start
