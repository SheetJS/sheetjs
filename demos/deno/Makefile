TESTS= x mjs
UNSTABLE= node
TEST_FILES=number_format_greek.xls

.PHONY: test
test: $(UNSTABLE) $(TESTS)

.PHONY: sheet2csv
sheet2csv: sheet2csv.ts
	deno compile -r --allow-read $<

$(TESTS): %: %.ts doit.ts $(TEST_FILES)
	deno run --allow-read --allow-write $<

# --unstable is required, see https://github.com/denoland/deno_std/issues/1900
$(UNSTABLE): %: %.ts doit.ts $(TEST_FILES)
	deno run --allow-read --allow-write --unstable $<

.PHONY: stream
stream: stream.ts
	deno run $<

number_format_greek.xls:
	cp ../../test_files/biff5/number_format_greek.xls .
