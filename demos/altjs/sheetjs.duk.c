/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "duktape.h"

#define FAIL_LOAD { \
	duk_push_undefined(ctx); \
	perror("Error in load_file"); \
	return 1; \
}

static char *read_file(const char *filename, size_t *sz) {
	FILE *f = fopen(filename, "rb");
	if(!f) return NULL;
	long fsize; { fseek(f, 0, SEEK_END); fsize = ftell(f); fseek(f, 0, SEEK_SET); }
	char *buf = (char *)malloc(fsize * sizeof(char));
	*sz = fread((void *) buf, 1, fsize, f);
	fclose(f);
	return buf;
}

static duk_int_t eval_file(duk_context *ctx, const char *filename) {
	size_t len; char *buf = read_file(filename, &len);
	if(!buf) FAIL_LOAD

	duk_push_lstring(ctx, (const char *)buf, (duk_size_t)len);
	duk_int_t retval = duk_peval(ctx);
	duk_pop(ctx);
	return retval;
}

static duk_int_t load_file(duk_context *ctx, const char *filename, const char *var) {
	size_t len; char *buf = read_file(filename, &len);
	if(!buf) FAIL_LOAD

	duk_push_external_buffer(ctx);
	duk_config_buffer(ctx, -1, buf, len);
	duk_put_global_string(ctx, var);
	return 0;
}

static duk_int_t save_file(duk_context *ctx, const char *filename, const char *var) {
	duk_get_global_string(ctx, var);
	duk_size_t sz;
	char *buf = (char *)duk_get_buffer_data(ctx, -1, &sz);

	if(!buf) return 1;
	FILE *f = fopen(filename, "wb"); fwrite(buf, 1, sz, f); fclose(f);
	return 0;
}

#define FAIL(cmd) { \
	printf("error in %s: %s\n", cmd, duk_safe_to_string(ctx, -1)); \
	duk_destroy_heap(ctx); \
	return res; \
}

#define DOIT(cmd) duk_eval_string_noresult(ctx, cmd);
int main(int argc, char *argv[]) {
	duk_int_t res = 0;

	/* initialize */
	duk_context *ctx = duk_create_heap_default();
	/* duktape does not expose a standard "global" by default */
	DOIT("var global = (function(){ return this; }).call(null);");

	/* load library */
	res = eval_file(ctx, "xlsx.full.min.js");
	if(res != 0) FAIL("library load")

	/* get version string */
	duk_eval_string(ctx, "XLSX.version");
	printf("SheetJS library version %s\n", duk_get_string(ctx, -1));
	duk_pop(ctx);

	/* read file */
#define INFILE "sheetjs.xlsx"
	res = load_file(ctx, INFILE, "buf");
	if(res != 0) FAIL("load " INFILE)

	/* parse workbook */
	DOIT("wb = XLSX.read(buf, {type:'buffer', cellNF:true});");
	DOIT("ws = wb.Sheets[wb.SheetNames[0]]");

	/* print CSV */
	duk_eval_string(ctx, "XLSX.utils.sheet_to_csv(ws)");
	printf("%s\n", duk_get_string(ctx, -1));
	duk_pop(ctx);

	/* change cell A1 to 3 */
	DOIT("ws['A1'].v = 3; delete ws['A1'].w;");

	/* write file */
#define WRITE_TYPE(BOOKTYPE) \
	DOIT("newbuf = (XLSX.write(wb, {type:'array', bookType:'" BOOKTYPE "'}));");\
	res = save_file(ctx, "sheetjsw." BOOKTYPE, "newbuf");\
	if(res != 0) FAIL("save sheetjsw." BOOKTYPE)

	WRITE_TYPE("xlsb")
	WRITE_TYPE("xlsx")
	WRITE_TYPE("xls")
	WRITE_TYPE("csv")

	/* cleanup */
	duk_destroy_heap(ctx);
	return res;
}
