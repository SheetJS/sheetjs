package main

import (
	b64 "encoding/base64"
	"fmt"
	"os"
	"io/ioutil"
	"github.com/dop251/goja"
)

func safe_run_file(vm *goja.Runtime, file string) {
	data, err := ioutil.ReadFile(file)
	if err != nil { panic(err) }
	src := string(data)
	_, err = vm.RunString(src)
	if err != nil { panic(err) }
}

func eval_string(vm *goja.Runtime, cmd string) goja.Value {
	v, err := vm.RunString(cmd)
	if err != nil { panic(err) }
	return v
}

func write_type(vm *goja.Runtime, t string) {
	/* due to some wonkiness with array passing, use base64 */
	b64str := eval_string(vm, "XLSX.write(wb, {type:'base64', bookType:'" + t + "'})")
	buf, err := b64.StdEncoding.DecodeString(b64str.String());
	if err != nil { panic(err) }
	err = ioutil.WriteFile("sheetjsg." + t, buf, 0644)
	if err != nil { panic(err) }
}

func main() {
	vm := goja.New()

	/* initialize */
	eval_string(vm, "if(typeof global == 'undefined') global = (function(){ return this; }).call(null);")

	/* load library */
	safe_run_file(vm, "shim.min.js")
	safe_run_file(vm, "xlsx.core.min.js")

	/* get version string */
	v := eval_string(vm, "XLSX.version")
	fmt.Printf("SheetJS library version %s\n", v)

	/* read file */
	data, err := ioutil.ReadFile(os.Args[1])
	if err != nil { panic(err) }
	vm.Set("buf", data)
	fmt.Printf("Loaded file %s\n", os.Args[1])

	/* parse workbook */
	eval_string(vm, "var bstr = ''; for(var i = 0; i < buf.length; ++i) bstr += String.fromCharCode(buf[i]);")
	eval_string(vm, "wb = XLSX.read(bstr, {type:'binary', cellNF:true});")
	eval_string(vm, "ws = wb.Sheets[wb.SheetNames[0]]")

	/* print CSV */
	csv := eval_string(vm, "XLSX.utils.sheet_to_csv(ws)")
	fmt.Printf("%s\n", csv)

	/* change cell A1 to 3 */
	eval_string(vm, "ws['A1'].v = 3; delete ws['A1'].w;")

	/* write file */
	//write_type(vm, "xlsb")
	//write_type(vm, "xlsx")
	write_type(vm, "xls")
	write_type(vm, "csv")
}
