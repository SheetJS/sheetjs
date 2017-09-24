#!/bin/bash

if [ ! -e SheetJS ]; then
	weexpack create SheetJS
	cd SheetJS
	npm install
	weexpack platform add ios
	# see https://github.com/weexteam/weex-pack/issues/133#issuecomment-295806132
	sed -i.bak 's/ATSDK-Weex/ATSDK/g' platforms/ios/Podfile
	cd -
fi
cp native.vue SheetJS/src/index.vue
if [ ! -e SheetJS/web/bootstrap.min.css ]; then
	curl -O https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css
	mv bootstrap.min.css SheetJS/web/
fi
