#!/bin/bash
rm -rf dist/mjs/
cp -r dist/esm dist/mjs
find dist/mjs -name '*.js' | while read x; do
	<"$x" awk '/(im|ex)port / { gsub(/";/, ".mjs\";"); } 1' > "${x%.js}.mjs"
	rm -f "$x"
done

