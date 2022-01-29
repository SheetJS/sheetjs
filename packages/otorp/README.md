# otorp

Recover [Protocol Buffer](https://en.wikipedia.org/wiki/Protocol_Buffers) v2
definitions from a Mach-O binary.


## Usage

```bash
$ npx otorp /path/to/macho/binary        # print all discovered defs to stdout
$ npx otorp /path/to/macho/binary out/   # write each discovered def to a file
```

This library and the embedded `otorp` CLI tool make the following assumptions:

- In a serialized `FileDescriptorProto`, the `name` field appears first.

- The name does not exceed 127 bytes in length.

- The name always ends in ".proto".

- There is at least one simple reference to the start of the definition.


## License

Please consult the attached LICENSE file for details.  All rights not explicitly
granted by the Apache 2.0 license are reserved by the Original Author.

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/sheetjs?pixel)](https://github.com/SheetJS/sheetjs)

