## Philosophy

<details>
  <summary><b>Philosophy</b> (click to show)</summary>

Prior to SheetJS, APIs for processing spreadsheet files were format-specific.
Third-party libraries either supported one format, or they involved a separate
set of classes for each supported file type.  Even though XLSB was introduced in
Excel 2007, nothing outside of SheetJS or Excel supported the format.

To promote a format-agnostic view, js-xlsx starts from a pure-JS representation
that we call the ["Common Spreadsheet Format"](#common-spreadsheet-format).
Emphasizing a uniform object representation enables new features like format
conversion (reading an XLSX template and saving as XLS) and circumvents the
"class trap".  By abstracting the complexities of the various formats, tools
need not worry about the specific file type!

A simple object representation combined with careful coding practices enables
use cases in older browsers and in alternative environments like ExtendScript
and Web Workers. It is always tempting to use the latest and greatest features,
but they tend to require the latest versions of browsers, limiting usability.

Utility functions capture common use cases like generating JS objects or HTML.
Most simple operations should only require a few lines of code.  More complex
operations generally should be straightforward to implement.

Excel pushes the XLSX format as default starting in Excel 2007.  However, there
are other formats with more appealing properties.  For example, the XLSB format
is spiritually similar to XLSX but files often tend up taking less than half the
space and open much faster!  Even though an XLSX writer is available, other
format writers are available so users can take advantage of the unique
characteristics of each format.

The primary focus of the Community Edition is correct data interchange, focused
on extracting data from any compatible data representation and exporting data in
various formats suitable for any third party interface.

</details>

