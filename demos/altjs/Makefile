.PHONY: all
all: duktape nashorn rhinojs swift

.PHONY: base
base:
	if [ ! -e sheetjs.xlsx ]; then node ../../tests/write.js; fi

.PHONY: duktape
duktape: base ## duktape / skookum demo
	sjs skookum.js

.PHONY: nashorn
nashorn: base ## nashorn demo
	if [ ! -e jvm-npm.js ]; then curl -O https://rawgit.com/nodyn/jvm-npm/master/src/main/javascript/jvm-npm.js; fi
	jjs nashorn.js

.PHONY: swift
swift: base ## swift demo
	if [ ! -e xlsx.swift.js ]; then cp ../../dist/xlsx.full.min.js xlsx.swift.js; fi
	./SheetJSCore.swift

.PHONY: rhinojs ## rhino demo
rhinojs: base SheetJSRhino.class
	java -cp .:SheetJS.jar:rhino.jar SheetJSRhino sheetjs.xlsx
	java -cp .:SheetJS.jar:rhino.jar SheetJSRhino sheetjs.xlsb
	java -cp .:SheetJS.jar:rhino.jar SheetJSRhino sheetjs.xls
	java -cp .:SheetJS.jar:rhino.jar SheetJSRhino sheetjs.xml.xls

RHDEPS=$(filter-out SheetJSRhino.class,$(patsubst %.java,%.class,$(wildcard com/sheetjs/*.java)))
$(RHDEPS): %.class: %.java rhino.jar
	javac -cp .:SheetJS.jar:rhino.jar $*.java

SheetJSRhino.class: $(RHDEPS)
	jar -cf SheetJS.jar $^ ../../dist/xlsx.full.min.js
	javac -cp .:SheetJS.jar:rhino.jar SheetJSRhino.java

rhino.jar:
	if [ ! -e rhino ]; then git clone https://github.com/mozilla/rhino; fi
	if [ ! -e rhino/build/rhino*/js.jar ]; then cd rhino; ant jar; fi
	cp rhino/build/rhino*/js.jar rhino.jar

.PHONY: clean
clean:
	rm SheetJS.jar *.class com/sheetjs/*.class
