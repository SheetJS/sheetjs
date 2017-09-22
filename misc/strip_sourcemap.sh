#!/bin/bash
# strip_sourcemap.sh -- strip sourcemaps from a JS file (missing from uglifyjs)
# Copyright (C) 2014-present  SheetJS
# note: this version also renames write_shift / read_shift to _W / _R

if [ $# -gt 0 ]; then
	if [ -e "$1" ]; then
		sed -i.sheetjs '/sourceMappingURL/d' "$1"
		sed -i.sheetjs 's/write_shift/_W/g; s/read_shift/_R/g' "$1"
	fi
else
	cat - | sed '/sourceMappingURL/d' | sed 's/write_shift/_W/g; s/read_shift/_R/g'
fi
