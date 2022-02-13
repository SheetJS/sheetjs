TESTS= x mjs jspm
UNSTABLE= node

.PHONY: test
test: $(UNSTABLE) $(TESTS)

.PHONY: sheet2csv
sheet2csv: sheet2csv.ts
	deno compile -r --allow-read $<

$(TESTS): %: %.ts doit.ts
	deno run --allow-read --allow-write $<

# --unstable is required, see https://github.com/denoland/deno_std/issues/1900
$(UNSTABLE): %: %.ts doit.ts
	deno run --allow-read --allow-write --unstable $<
