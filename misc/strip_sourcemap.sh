#!/bin/bash
# strip_sourcemap.sh -- strip sourcemaps from a JS file (missing from uglifyjs)
# Copyright (C) 2014-present  SheetJS

if [ $# -gt 0 ]; then
	if [ -e "$1" ]; then
		sed -i.sheetjs '/sourceMappingURL/d' "$1"
	fi
else
	cat - | sed '/sourceMappingURL/d'
fi
