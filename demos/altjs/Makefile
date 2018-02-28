.PHONY: all
all: duktape nashorn rhinojs swift

.PHONY: base
base:
	if [ ! -e sheetjs.xlsx ]; then node ../../tests/write.js; fi
	if [ ! -e xlsx.full.min.js ]; then cp ../../dist/xlsx.full.min.js .; fi

.PHONY: duktape
duktape: base ## duktape demo
	bash ./duktape.sh
	gcc -std=c99 -Wall -osheetjs.duk sheetjs.duk.c duktape.c -lm
	./sheetjs.duk

.PHONY: nashorn
nashorn: base ## nashorn demo
	jjs nashorn.js

.PHONY: swift
swift: base ## swift demo
	swiftc SheetJSCore.swift main.swift -o SheetJSSwift
	./SheetJSSwift

.PHONY: chakra
chakra: base ## Chakra demo
	node -pe "fs.writeFileSync('payload.js', 'var payload = \"' + fs.readFileSync('sheetjs.xlsx').toString('base64') + '\";')"
	cat global.js xlsx.full.min.js payload.js chakra.js > xlsx.chakra.js
	chakra ./xlsx.chakra.js

.PHONY: rhinojs ## rhino demo
rhinojs: base SheetJSRhino.class
	for ext in xlsx xlsb biff8.xls xml.xls; do java -cp .:SheetJS.jar:rhino.jar SheetJSRhino sheetjs.$$ext; done

RHDEPS=$(filter-out SheetJSRhino.class,$(patsubst %.java,%.class,$(wildcard com/sheetjs/*.java)))
$(RHDEPS): %.class: %.java rhino.jar
	javac -cp .:SheetJS.jar:rhino.jar $*.java

SheetJSRhino.class: $(RHDEPS)
	jar -cf SheetJS.jar $^ xlsx.full.min.js
	javac -cp .:SheetJS.jar:rhino.jar SheetJSRhino.java

rhino.jar:
	if [ ! -e rhino ]; then git clone --depth=1 https://github.com/mozilla/rhino; fi
	#if [ ! -e rhino/build/rhino*/js.jar ]; then cd rhino; ant jar; fi
	#cp rhino/build/rhino*/js.jar rhino.jar
	if [ ! -e rhino/buildGradle/libs/rhino*.jar ]; then cd rhino; ./gradlew jar; fi
	cp rhino/buildGradle/libs/rhino*.jar rhino.jar

.PHONY: clean
clean:
	rm SheetJS.jar *.class com/sheetjs/*.class
