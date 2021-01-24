.PHONY: init
init:
	cp ../../dist/xlsx.full.min.js .
	if [ ! -e logo.png ]; then curl -O http://sheetjs.com/logo.png; fi

.PHONY: lint
lint:
	eslint content.js popup.js table.js
