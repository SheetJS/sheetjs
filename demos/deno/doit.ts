const bts = [
	"xlsx",
	"xlsb",
	"xls",
	"csv",
	"fods",
	"xlml",
	"slk"
];
export default function doit(XLSX: any, tag: string) {
	const path = "number_format_greek.xls";
	let workbook: any;

	/* read file */
	try {
		workbook = XLSX.readFile(path);
	} catch(e) {
		console.log(e);
		console.error("Cannot use readFile, falling back to read");
		const rawdata = Deno.readFileSync(path);
		workbook = XLSX.read(rawdata, {type: "buffer"});
	}

	/* write file */
	try {
		bts.forEach(bt => {
			console.log(bt);
			XLSX.writeFile(workbook, `${tag}.${bt}`);
		});
	} catch(e) {
		console.log(e);
		console.error("Cannot use writeFile, falling back to write");
		bts.forEach(bt => {
			console.log(bt);
			const buf = XLSX.write(workbook, {type: "buffer", bookType: bt});
			if(typeof buf == "string") {
				const nbuf = new Uint8Array(buf.length);
				for(let i = 0; i < buf.length; ++i) nbuf[i] = buf.charCodeAt(i);
				Deno.writeFileSync(`${tag}.${bt}`, nbuf);
			} else Deno.writeFileSync(`${tag}.${bt}`, new Uint8Array(buf));
		});
	}
}
