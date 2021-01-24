APPS= aftereffects estoolkit illustrator indesign photoshop
TARGETS=$(patsubst %,%.jsx,$(APPS))

.PHONY: all
all: deps $(TARGETS)

.PHONY: deps
deps:
	cp ../../dist/xlsx.extendscript.js .

%.base:
	echo "#target $*" > $@

.PHONY: $(TARGETS)
$(TARGETS):%.jsx:%.base test.jsx
	cat $^ > $@
