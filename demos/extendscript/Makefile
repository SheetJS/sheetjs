APPS= aftereffects estoolkit illustrator indesign photoshop
TARGETS=$(patsubst %,%.jsx,$(APPS))

.PHONY: all
all: deps $(TARGETS)

.PHONY: deps
deps:
	cp ../../shim.js .
	cp ../../jszip.js .
	cp ../../xlsx.flow.js .

%.base:
	echo "#target $*" > $@

.PHONY: $(TARGETS)
$(TARGETS):%.jsx:%.base test.jsx
	cat $^ > $@
