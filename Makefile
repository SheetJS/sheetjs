DEPS=$(wildcard bits/*.js)

xlsx.js: $(DEPS)
	cat $^ > $@

.PHONY: clean
clean:
	rm xlsx.js
