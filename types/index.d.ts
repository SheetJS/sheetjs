/* index.d.ts (C) 2015-present SheetJS and contributors */
// TypeScript Version: 2.2

/** Version string */
export const version: string;

/** SSF Formatter Library */
export const SSF: any;

/** Attempts to read filename and parse */
export function readFile(filename: string, opts?: ParsingOptions): WorkBook;
/** Attempts to parse data */
export function read(data: any, opts?: ParsingOptions): WorkBook;
/** NODE ONLY! Attempts to write workbook data to filename */
export function writeFile(data: WorkBook, filename: string, opts?: WritingOptions): any;
/** Attempts to write the workbook data */
export function write(data: WorkBook, opts?: WritingOptions): any;

/** Utility Functions */
export const utils: XLSX$Utils;
/** Stream Utility Functions */
export const stream: StreamUtils;

/** Number Format (either a string or an index to the format table) */
export type NumberFormat = string | number;

/** Basic File Properties */
export interface Properties {
    /** Summary tab "Title" */
    Title?: string;
    /** Summary tab "Subject" */
    Subject?: string;
    /** Summary tab "Author" */
    Author?: string;
    /** Summary tab "Manager" */
    Manager?: string;
    /** Summary tab "Company" */
    Company?: string;
    /** Summary tab "Category" */
    Category?: string;
    /** Summary tab "Keywords" */
    Keywords?: string;
    /** Summary tab "Comments" */
    Comments?: string;
    /** Statistics tab "Last saved by" */
    LastAuthor?: string;
    /** Statistics tab "Created" */
    CreatedDate?: Date;
}

/** Other supported properties */
export interface FullProperties extends Properties {
    ModifiedDate?: Date;
    Application?: string;
    AppVersion?: string;
    DocSecurity?: string;
    HyperlinksChanged?: boolean;
    SharedDoc?: boolean;
    LinksUpToDate?: boolean;
    ScaleCrop?: boolean;
    Worksheets?: number;
    SheetNames?: string[];
    ContentStatus?: string;
    LastPrinted?: string;
    Revision?: string | number;
    Version?: string;
    Identifier?: string;
    Language?: string;
}

export interface CommonOptions {
    /**
     * If true, throw errors when features are not understood
     * @default false
     */
    WTF?: boolean;

    /**
     * When reading a file, store dates as type d (default is n)
     * When writing XLSX/XLSM file, use native date (default uses date codes)
     * @default false
     */
    cellDates?: boolean;
}

/** Options for read and readFile */
export interface ParsingOptions extends CommonOptions {
    /** Input data encoding */
    type?: 'base64' | 'binary' | 'buffer' | 'file' | 'array';

    /**
     * Save formulae to the .f field
     * @default true
     */
    cellFormula?: boolean;

    /**
     * Parse rich text and save HTML to the .h field
     * @default true
     */
    cellHTML?: boolean;

    /**
     * Save number format string to the .z field
     * @default false
     */
    cellNF?: boolean;

    /**
     * Save style/theme info to the .s field
     * @default false
     */
    cellStyles?: boolean;

    /**
     * Generate formatted text to the .w field
     * @default true
     */
    cellText?: boolean;

    /** Override default date format (code 14) */
    dateNF?: string;

    /**
     * Create cell objects for stub cells
     * @default false
     */
    sheetStubs?: boolean;

    /**
     * If >0, read the first sheetRows rows
     * @default 0
     */
    sheetRows?: number;

    /**
     * If true, parse calculation chains
     * @default false
     */
    bookDeps?: boolean;

    /**
     * If true, add raw files to book object
     * @default false
     */
    bookFiles?: boolean;

    /**
     * If true, only parse enough to get book metadata
     * @default false
     */
    bookProps?: boolean;

    /**
     * If true, only parse enough to get the sheet names
     * @default false
     */
    bookSheets?: boolean;

    /**
     * If true, expose vbaProject.bin to vbaraw field
     * @default false
     */
    bookVBA?: boolean;

    /**
     * If defined and file is encrypted, use password
     * @default ''
     */
    password?: string;
}

/** Options for write and writeFile */
export interface WritingOptions extends CommonOptions {
    /** Output data encoding */
    type?: 'base64' | 'binary' | 'buffer' | 'file';

    /**
     * Generate Shared String Table
     * @default false
     */
    bookSST?: boolean;

    /**
     * File format of generated workbook
     * @default 'xlsx'
     */
    bookType?: BookType;

    /**
     * Name of Worksheet (for single-sheet formats)
     * @default ''
     */
    sheet?: string;

    /**
     * Use ZIP compression for ZIP-based formats
     * @default false
     */
    compression?: boolean;

    /** Override workbook properties on save */
    Props?: Properties;
}

/** Workbook Object */
export interface WorkBook {
    /**
     * A dictionary of the worksheets in the workbook.
     * Use SheetNames to reference these.
     */
    Sheets: { [sheet: string]: WorkSheet };

    /** Ordered list of the sheet names in the workbook */
    SheetNames: string[];

    /**
     * an object storing the standard properties. wb.Custprops stores custom properties.
     * Since the XLS standard properties deviate from the XLSX standard, XLS parsing stores core properties in both places.
     */
    Props?: FullProperties;

    Workbook?: WBProps;
}

export interface SheetProps {
    /** Sheet Visibility (0=Visible 1=Hidden 2=VeryHidden) */
    Hidden?: 0 | 1 | 2;
}

/** Defined Name Object */
export interface DefinedName {
    /** Name */
    Name: string;

    /** Reference */
    Ref: string;

    /** Scope (undefined for workbook scope) */
    Sheet?: number;

    /** Name comment */
    Comment?: string;
}

/** Workbook-Level Attributes */
export interface WBProps {
    /** Sheet Properties */
    Sheets?: SheetProps[];

    /** Defined Names */
    Names?: DefinedName[];

    /** Other Workbook Properties */
    WBProps?: WorkbookProperties;
}

/** Other Workbook Properties */
export interface WorkbookProperties {
    /** Worksheet Epoch (1904 if true, 1900 if false) */
    date1904?: boolean;

    /** Warn or strip personally identifying info on save */
    filterPrivacy?: boolean;
}

/** Column Properties Object */
export interface ColInfo {
    /* --- visibility --- */

    /** if true, the column is hidden */
    hidden?: boolean;

    /* --- column width --- */

    /** width in Excel's "Max Digit Width", width*256 is integral */
    width?: number;

    /** width in screen pixels */
    wpx?: number;

    /** width in "characters" */
    wch?: number;

    /** Excel's "Max Digit Width" unit, always integral */
    MDW?: number;
}

/** Row Properties Object */
export interface RowInfo {
    /* --- visibility --- */

    /** if true, the column is hidden */
    hidden?: boolean;

    /* --- row height --- */

    /** height in screen pixels */
    hpx?: number;

    /** height in points */
    hpt?: number;
}

/**
 * Write sheet protection properties.
 */
export interface ProtectInfo {
    /**
     * The password for formats that support password-protected sheets
     * (XLSX/XLSB/XLS). The writer uses the XOR obfuscation method.
     */
    password?: string;
    /**
     * Select locked cells
     * @default: true
     */
    selectLockedCells?: boolean;
    /**
     * Select unlocked cells
     * @default: true
     */
    selectUnlockedCells?: boolean;
    /**
     * Format cells
     * @default: false
     */
    formatCells?: boolean;
    /**
     * Format columns
     * @default: false
     */
    formatColumns?: boolean;
    /**
     * Format rows
     * @default: false
     */
    formatRows?: boolean;
    /**
     * Insert columns
     * @default: false
     */
    insertColumns?: boolean;
    /**
     * Insert rows
     * @default: false
     */
    insertRows?: boolean;
    /**
     * Insert hyperlinks
     * @default: false
     */
    insertHyperlinks?: boolean;
    /**
     * Delete columns
     * @default: false
     */
    deleteColumns?: boolean;
    /**
     * Delete rows
     * @default: false
     */
    deleteRows?: boolean;
    /**
     * Sort
     * @default: false
     */
    sort?: boolean;
    /**
     * Filter
     * @default: false
     */
    autoFilter?: boolean;
    /**
     * Use PivotTable reports
     * @default: false
     */
    pivotTables?: boolean;
    /**
     * Edit objects
     * @default: true
     */
    objects?: boolean;
    /**
     * Edit scenarios
     * @default: true
     */
    scenarios?: boolean;
}

/** Page Margins -- see Excel Page Setup .. Margins diagram for explanation */
export interface MarginInfo {
    /** Left side margin (inches) */
    left?: number;
    /** Right side margin (inches) */
    right?: number;
    /** Top side margin (inches) */
    top?: number;
    /** Bottom side margin (inches) */
    bottom?: number;
    /** Header top margin (inches) */
    header?: number;
    /** Footer bottom height (inches) */
    footer?: number;
}
export type SheetType = 'sheet' | 'chart';
export type SheetKeys = string | MarginInfo | SheetType;
/** General object representing a Sheet (worksheet or chartsheet) */
export interface Sheet {
    /**
     * Indexing with a cell address string maps to a cell object
     * Special keys start with '!'
     */
    [cell: string]: CellObject | SheetKeys | any;

    /** Sheet type */
    '!type'?: SheetType;

    /** Sheet Range */
    '!ref'?: string;

    /** Page Margins */
    '!margins'?: MarginInfo;
}

/** AutoFilter properties */
export interface AutoFilterInfo {
    /** Range of the AutoFilter table */
    ref: string;
}
export type WSKeys = SheetKeys | ColInfo[] | RowInfo[] | Range[] | ProtectInfo | AutoFilterInfo;

/** Worksheet Object */
export interface WorkSheet extends Sheet {
    /**
     * Indexing with a cell address string maps to a cell object
     * Special keys start with '!'
     */
    [cell: string]: CellObject | WSKeys | any;

    /** Column Info */
    '!cols'?: ColInfo[];

    /** Row Info */
    '!rows'?: RowInfo[];

    /** Merge Ranges */
    '!merges'?: Range[];

    /** Worksheet Protection info */
    '!protect'?: ProtectInfo;

    /** AutoFilter info */
    '!autofilter'?: AutoFilterInfo;
}

/**
 * The Excel data type for a cell.
 * b Boolean, n Number, e error, s String, d Date, z Stub
 */
export type ExcelDataType = 'b' | 'n' | 'e' | 's' | 'd' | 'z';

/**
 * Type of generated workbook
 * @default 'xlsx'
 */
export type BookType = 'xlsx' | 'xlsm' | 'xlsb' | 'biff2' | 'xlml' | 'ods' | 'fods' | 'csv' | 'txt' | 'sylk' | 'html' | 'dif' | 'prn';

/** Comment element */
export interface Comment {
    /** Author of the comment block */
    a?: string;

    /** Plaintext of the comment */
    t: string;
}

/** Link object */
export interface Hyperlink {
    /** Target of the link (HREF) */
    Target: string;

    /** Plaintext tooltip to display when mouse is over cell */
    Tooltip?: string;
}

/** Worksheet Cell Object */
export interface CellObject {
    /** The raw value of the cell.  Can be omitted if a formula is specified */
    v?: string | number | boolean | Date;

    /** Formatted text (if applicable) */
    w?: string;

    /**
     * The Excel Data Type of the cell.
     * b Boolean, n Number, e Error, s String, d Date, z Empty
     */
    t: ExcelDataType;

    /** Cell formula (if applicable) */
    f?: string;

    /** Range of enclosing array if formula is array formula (if applicable) */
    F?: string;

    /** Rich text encoding (if applicable) */
    r?: any;

    /** HTML rendering of the rich text (if applicable) */
    h?: string;

    /** Comments associated with the cell */
    c?: Comment[];

    /** Number format string associated with the cell (if requested) */
    z?: NumberFormat;

    /** Cell hyperlink object (.Target holds link, .tooltip is tooltip) */
    l?: Hyperlink;

    /** The style/theme of the cell (if applicable) */
    s?: any;
}

/** Simple Cell Address */
export interface CellAddress {
    /** Column number */
    c: number;
    /** Row number */
    r: number;
}

/**
 * Range object (representing ranges like "A1:B2")
 */
export interface Range {
    /** Starting cell */
    s: CellAddress;
    /** Ending cell */
    e: CellAddress;
}

export interface Sheet2CSVOpts {
    /** Field Separator ("delimiter") */
    FS?: string;

    /** Record Separator ("row separator") */
    RS?: string;

    /** Use specified date format */
    dateNF?: NumberFormat;

    /** Remove trailing field separators in each record */
    strip?: boolean;

    /** Include blank lines in the CSV output */
    blankrows?: boolean;

    /** Skip hidden rows and columns in the CSV output */
    skipHidden?: boolean;
}

export interface Sheet2HTMLOpts {
    /** Add contenteditable to every cell */
    editable?: boolean;

    /** Header HTML */
    header?: string;

    /** Footer HTML */
    footer?: string;
}

export interface Sheet2JSONOpts {
    /** Use specified format for date cells */
    dateNF?: NumberFormat;

    /** Output format */
    header?: "A"|number|string[];

    /** Override worksheet range */
    range?: any;

    /** Include or omit blank lines in the output */
    blankrows?: boolean;

    /** Default value for null/undefined values */
    defval?: any;

    /** if true, return raw data; if false, return formatted text */
    raw?: boolean;
}

export interface AOA2SheetOpts {
    /** Use specified format for date cells */
    dateNF?: NumberFormat;

    /**
     * Store dates as type d (default is n)
     * @default false
     */
    cellDates?: boolean;

    /**
     * Create cell objects for stub cells
     * @default false
     */
    sheetStubs?: boolean;
}

export interface JSON2SheetOpts {
    /** Use specified date format */
    dateNF?: NumberFormat;
}

export interface Table2SheetOpts {
    /** Use specified date format */
    dateNF?: NumberFormat;
}

/** General utilities */
export interface XLSX$Utils {
    /* --- Import Functions --- */

    /** Converts an array of arrays of JS data to a worksheet. */
    aoa_to_sheet<T>(data: T[][], opts?: AOA2SheetOpts): WorkSheet;
    aoa_to_sheet(data: any[][], opts?: AOA2SheetOpts): WorkSheet;

    /** Converts an array of JS objects to a worksheet. */
    json_to_sheet<T>(data: T[], opts?: JSON2SheetOpts): WorkSheet;
    json_to_sheet(data: any[], opts?: JSON2SheetOpts): WorkSheet;

    /** BROWSER ONLY! Converts a TABLE DOM element to a worksheet. */
    table_to_sheet(data: any,  opts?: Table2SheetOpts): WorkSheet;
    table_to_book(data: any,  opts?: Table2SheetOpts): WorkBook;

    /* --- Export Functions --- */

    /** Converts a worksheet object to an array of JSON objects */
    sheet_to_json<T>(worksheet: WorkSheet, opts?: Sheet2JSONOpts): T[];
    sheet_to_json(worksheet: WorkSheet, opts?: Sheet2JSONOpts): any[][];
    sheet_to_json(worksheet: WorkSheet, opts?: Sheet2JSONOpts): any[];

    /** Generates delimiter-separated-values output */
    sheet_to_csv(worksheet: WorkSheet, options?: Sheet2CSVOpts): string;

    /** Generates HTML */
    sheet_to_html(worksheet: WorkSheet, options?: Sheet2HTMLOpts): string;

    /** Generates a list of the formulae (with value fallbacks) */
    sheet_to_formulae(worksheet: WorkSheet): string[];

    /* --- Cell Address Utilities --- */

    /** Converts 0-indexed cell address to A1 form */
    encode_cell(cell: CellAddress): string;

    /** Converts 0-indexed row to A1 form */
    encode_row(row: number): string;

    /** Converts 0-indexed column to A1 form */
    encode_col(col: number): string;

    /** Converts 0-indexed range to A1 form */
    encode_range(s: CellAddress, e: CellAddress): string;
    encode_range(r: Range): string;

    /** Converts A1 cell address to 0-indexed form */
    decode_cell(address: string): CellAddress;

    /** Converts A1 row to 0-indexed form */
    decode_row(row: string): number;

    /** Converts A1 column to 0-indexed form */
    decode_col(col: string): number;

    /** Converts A1 range to 0-indexed form */
    decode_range(range: string): Range;

    /* --- General Utilities --- */

    /** Creates a new workbook */
    book_new(): WorkBook;

    /** Append a worksheet to a workbook */
    book_append_sheet(workbook: WorkBook, worksheet: WorkSheet, name?: string): void;

    /** Set sheet visibility (visible/hidden/very hidden) */
    book_set_sheet_visibility(workbook: WorkBook, sheet: number|string, visibility: number): void;

    /** Set number format for a cell */
    cell_set_number_format(cell: CellObject, fmt: string|number): CellObject;

    /** Set hyperlink for a cell */
    cell_set_hyperlink(cell: CellObject, target: string, tooltip?: string): CellObject;

    /** Add comment to a cell */
    cell_add_comment(cell: CellObject, text: string, author?: string): void;

    /** Assign an Array Formula to a range */
    sheet_set_array_formula(ws: WorkSheet, range: Range|string, formula: string): WorkSheet;

    consts: XLSX$Consts;
}

export interface XLSX$Consts {
    /* --- Sheet Visibility --- */

    /** Visibility: Visible */
    SHEET_VISIBLE: 0;

    /** Visibility: Hidden */
    SHEET_HIDDEN: 1;

    /** Visibility: Very Hidden */
    SHEET_VERYHIDDEN: 2;
}

/** NODE ONLY! these return Readable Streams */
export interface StreamUtils {
    /** CSV output stream, generate one line at a time */
    to_csv(sheet: WorkSheet, opts?: Sheet2CSVOpts): any;
    /** HTML output stream, generate one line at a time */
    to_html(sheet: WorkSheet, opts?: Sheet2HTMLOpts): any;
}
