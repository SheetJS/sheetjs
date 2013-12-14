.PHONY: test ssf
ssf: ssf.md
	voc ssf.md

test:
	npm test
