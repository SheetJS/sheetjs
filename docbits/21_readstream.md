### Note on Streaming Read

The most common and interesting formats (XLS, XLSX/M, XLSB, ODS) are ultimately
ZIP or CFB containers of files.  Neither format puts the directory structure at
the beginning of the file: ZIP files place the Central Directory records at the
end of the logical file, while CFB files can place the FAT structure anywhere in
the file! As a result, to properly handle these formats, a streaming function
would have to buffer the entire file before commencing.  That belies the
expectations of streaming, so we do not provide any streaming read API.  If you
really want to stream, there are node modules like `concat-stream` that will do
the buffering for you.

