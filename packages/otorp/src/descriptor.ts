/*! otorp (C) 2013-present SheetJS -- http://sheetjs.com */

import { parse_shallow, varint_to_i32, mappa } from "./proto";
import { u8str, indent } from "./util";

var TYPES = [
	"error",
	"double",
	"float",
	"int64",
	"uint64",
	"int32",
	"fixed64",
	"fixed32",
	"bool",
	"string",
	"group",
	"message",
	"bytes",
	"uint32",
	"enum",
	"sfixed32",
	"sfixed64",
	"sint32",
	"sint64"
];
export { TYPES };


interface FileOptions {
	javaPackage?: string;
	javaOuterClassname?: string;
	javaMultipleFiles?: string;
	goPackage?: string;
}
function parse_FileOptions(buf: Uint8Array): FileOptions {
	var data = parse_shallow(buf);
	var out: FileOptions = {};
	if(data[1]?.[0]) out.javaPackage = u8str(data[1][0].data);
	if(data[8]?.[0]) out.javaOuterClassname = u8str(data[8][0].data);
	if(data[11]?.[0]) out.goPackage = u8str(data[11][0].data);
	return out;
}


interface EnumValue {
	name?: string;
	number?: number;
}
function parse_EnumValue(buf: Uint8Array): EnumValue {
	var data = parse_shallow(buf);
	var out: EnumValue = {};
	if(data[1]?.[0]) out.name = u8str(data[1][0].data);
	if(data[2]?.[0]) out.number = varint_to_i32(data[2][0].data);
	return out;
}


interface Enum {
	name?: string;
	value?: EnumValue[];
}
function parse_Enum(buf: Uint8Array): Enum {
	var data = parse_shallow(buf);
	var out: Enum = {};
	if(data[1]?.[0]) out.name = u8str(data[1][0].data);
	out.value = mappa(data[2], parse_EnumValue);
	return out;
}
var write_Enum = (en: Enum): string => {
	var out = [`enum ${en.name} {`];
	en.value?.forEach(({name, number}) => out.push(`  ${name} = ${number};`));
	return out.concat(`}`).join("\n");
};
export { Enum, parse_Enum, write_Enum };


interface FieldOptions {
	packed?: boolean;
	deprecated?: boolean;
}
function parse_FieldOptions(buf: Uint8Array): FieldOptions {
	var data = parse_shallow(buf);
	var out: FieldOptions = {};
	if(data[2]?.[0]) out.packed = !!data[2][0].data;
	if(data[3]?.[0]) out.deprecated = !!data[3][0].data;
	return out;
}


interface Field {
	name?: string;
	extendee?: string;
	number?: number;
	label?: number;
	type?: number;
	typeName?: string;
	defaultValue?: string;
	options?: FieldOptions;
}
function parse_Field(buf: Uint8Array): Field {
	var data = parse_shallow(buf);
	var out: Field = {};
	if(data[1]?.[0]) out.name = u8str(data[1][0].data);
	if(data[2]?.[0]) out.extendee = u8str(data[2][0].data);
	if(data[3]?.[0]) out.number = varint_to_i32(data[3][0].data);
	if(data[4]?.[0]) out.label = varint_to_i32(data[4][0].data);
	if(data[5]?.[0]) out.type = varint_to_i32(data[5][0].data);
	if(data[6]?.[0]) out.typeName = u8str(data[6][0].data);
	if(data[7]?.[0]) out.defaultValue = u8str(data[7][0].data);
	if(data[8]?.[0]) out.options = parse_FieldOptions(data[8][0].data);
	return out;
}
function write_Field(field: Field): string {
	var out = [];
	var label = ["", "optional ", "required ", "repeated "][field.label] || "";
	var type = field.typeName || TYPES[field.type] || "s5s";
	var opts = [];
	if(field.defaultValue) opts.push(`default = ${field.defaultValue}`);
	if(field.options?.packed) opts.push(`packed = true`);
	if(field.options?.deprecated) opts.push(`deprecated = true`);
	var os = opts.length ? ` [${opts.join(", ")}]`: "";
	out.push(`${label}${type} ${field.name} = ${field.number}${os};`);
	return out.length ? indent(out.join("\n"), 1) : "";
}
export { Field, parse_Field, write_Field };


function write_extensions(ext: Field[], xtra = false, coalesce = true): string {
	var res = [];
	var xt: Array<[string, Array<Field>]> = [];
	ext.forEach(ext => {
		if(!ext.extendee) return;
		var row = coalesce ?
			xt.find(x => x[0] == ext.extendee) :
			(xt[xt.length - 1]?.[0] == ext.extendee ? xt[xt.length - 1]: null);
		if(row) row[1].push(ext);
		else xt.push([ext.extendee, [ext]]);
	});
	xt.forEach(extrow => {
		var out = [`extend ${extrow[0]} {`];
		extrow[1].forEach(ext => out.push(write_Field(ext)));
		res.push(out.concat(`}`).join("\n") + (xtra ? "\n" : ""));
	});
	return res.join("\n");
}
export { write_extensions };


interface ExtensionRange { start?: number; end?: number; }
interface MessageType {
	name?: string;
	nestedType?: MessageType[];
	enumType?: Enum[];
	field?: Field[];
	extension?: Field[];
	extensionRange?: ExtensionRange[];
}
function parse_mtype(buf: Uint8Array): MessageType {
	var data = parse_shallow(buf);
	var out: MessageType = {};
	if(data[1]?.[0]) out.name = u8str(data[1][0].data);
	if(data[2]?.length >= 1) out.field = mappa(data[2], parse_Field);
	if(data[3]?.length >= 1) out.nestedType = mappa(data[3], parse_mtype);
	if(data[4]?.length >= 1) out.enumType = mappa(data[4], parse_Enum);
	if(data[6]?.length >= 1) out.extension = mappa(data[6], parse_Field);
	if(data[5]?.length >= 1) out.extensionRange = data[5].map(d => {
		var data = parse_shallow(d.data);
		var out: ExtensionRange = {};
		if(data[1]?.[0]) out.start = varint_to_i32(data[1][0].data);
		if(data[2]?.[0]) out.end	 = varint_to_i32(data[2][0].data);
		return out;
	});
	return out;
}
var write_mtype = (message: MessageType): string => {
	var out = [ `message ${message.name} {` ];
	message.nestedType?.forEach(m => out.push(indent(write_mtype(m), 1)));
	message.enumType?.forEach(en => out.push(indent(write_Enum(en), 1)));
	message.field?.forEach(field => out.push(write_Field(field)));
	if(message.extensionRange) message.extensionRange.forEach(er => out.push(`  extensions ${er.start} to ${er.end - 1};`));
	if(message.extension?.length) out.push(indent(write_extensions(message.extension), 1));
	return out.concat(`}`).join("\n");
};


interface Descriptor {
	name?: string;
	package?: string;
	dependency?: string[];
	messageType?: MessageType[];
	enumType?: Enum[];
	extension?: Field[];
	options?: FileOptions;
}
function parse_FileDescriptor(buf: Uint8Array): Descriptor {
	var data = parse_shallow(buf);
	var out: Descriptor = {};
	if(data[1]?.[0]) out.name = u8str(data[1][0].data);
	if(data[2]?.[0]) out.package = u8str(data[2][0].data);
	if(data[3]?.[0]) out.dependency = data[3].map(x => u8str(x.data));

	if(data[4]?.length >= 1) out.messageType = mappa(data[4], parse_mtype);
	if(data[5]?.length >= 1) out.enumType = mappa(data[5], parse_Enum);
	if(data[7]?.length >= 1) out.extension = mappa(data[7], parse_Field);

	if(data[8]?.[0]) out.options = parse_FileOptions(data[8][0].data);

	return out;
}
var write_FileDescriptor = (pb: Descriptor): string => {
	var out = [
		'syntax = "proto2";',
		''
	];
	if(pb.dependency) pb.dependency.forEach((n: string) => { if(n) out.push(`import "${n}";`); });
	if(pb.package) out.push(`package ${pb.package};\n`);
	if(pb.options) {
		var o = out.length;

		if(pb.options.javaPackage) out.push(`option java_package = "${pb.options.javaPackage}";`);
		if(pb.options.javaOuterClassname?.replace(/\W/g, "")) out.push(`option java_outer_classname = "${pb.options.javaOuterClassname}";`);
		if(pb.options.javaMultipleFiles) out.push(`option java_multiple_files = true;`);
		if(pb.options.goPackage) out.push(`option go_package = "${pb.options.goPackage}";`);

		if(out.length > o) out.push('');
	}

	pb.enumType?.forEach(en => { if(en.name) out.push(write_Enum(en) + "\n"); });
	pb.messageType?.forEach(m => { if(m.name) { var o = write_mtype(m); if(o) out.push(o + "\n"); }});

	if(pb.extension?.length) {
		var e = write_extensions(pb.extension, true, false);
		if(e) out.push(e);
	}
	return out.join("\n") + "\n";
};
export { Descriptor, parse_FileDescriptor, write_FileDescriptor };
