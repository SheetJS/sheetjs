#!/bin/bash
DUKTAPE_VER=2.2.0
if [ ! -e duktape-$DUKTAPE_VER ]; then
	if [ ! -e duktape-$DUKTAPE_VER.tar ]; then
		if [ ! -e duktape-$DUKTAPE_VER.tar.xz ]; then
			curl -O http://duktape.org/duktape-$DUKTAPE_VER.tar.xz
		fi
		xz -d duktape-$DUKTAPE_VER.tar.xz
	fi
	tar -xf duktape-$DUKTAPE_VER.tar
fi

for f in duktape.{c,h} duk_config.h; do
	cp duktape-$DUKTAPE_VER/src/$f .
done

