export const version = XLSX.version;
export {
	parse_xlscfb,
	parse_zip,
	readSync as read,
	readFileSync as readFile,
	readFileSync,
	writeSync as write,
	writeFileSync as writeFile,
	writeFileSync,
	writeFileAsync,
	writeSyncXLSX as writeXLSX,
	writeFileSyncXLSX as writeFileXLSX,
	utils,
	set_fs,
	set_cptable,
	__stream as stream,
	SSF,
	CFB
};
export default {
	parse_xlscfb,
	parse_zip,
	read: readSync,
	readFile: readFileSync,
	readFileSync,
	write: writeSync,
	writeFile: writeFileSync,
	writeFileSync,
	writeFileAsync,
	writeXLSX: writeSyncXLSX,
	writeFileXLSX: writeFileSyncXLSX,
	utils,
	set_fs,
	set_cptable,
	stream: __stream,
	SSF,
	CFB
}
