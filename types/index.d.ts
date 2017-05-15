/* index.d.ts (C) 2015-present SheetJS and contributors */
// TypeScript Version: 2.2

/** Attempts to read filename and parse */
export function readFile(filename: string, opts?: ParsingOptions): WorkBook;
/** Attempts to parse data */
export function read(data: any, opts?: ParsingOptions): WorkBook;
/** Attempts to write workbook data to filename */
export function writeFile(data: WorkBook, filename: string, opts?: WritingOptions): any;
/** Attempts to write the workbook data */
export function write(data: WorkBook, opts?: WritingOptions): any;

export const utils: Utils;

export interface Properties {
    Title?: string;
    Subject?: string;
    Author?: string;
    Manager?: string;
    Company?: string;
    Category?: string;
    Keywords?: string;
    Comments?: string;
    LastAuthor?: string;
    CreatedDate?: Date;
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
}

export interface ParsingOptions {
    /**
     * Input data encoding
     */
    type?: 'base64' | 'binary' | 'buffer' | 'array' | 'file';

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
     * Store dates as type d (default is n)
     * @default false
     */
    cellDates?: boolean;

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

export interface WritingOptions {
    /**
     * Output data encoding
     */
    type?: 'base64' | 'binary' | 'buffer' | 'file';

    /**
     * Store dates as type d (default is n)
     * @default false
     */
    cellDates?: boolean;

    /**
     * Generate Shared String Table
     * @default false
     */
    bookSST?: boolean;

    /**
     * Type of Workbook
     * @default 'xlsx'
     */
    bookType?: 'xlsx' | 'xlsm' | 'xlsb' | 'biff2' | 'xlml' | 'ods' | 'fods' | 'csv' | 'txt' | 'sylk' | 'html' | 'dif' | 'prn';

    /**
     * Name of Worksheet for single-sheet formats
     * @default ''
     */
    sheet?: string;

    /**
     * Use ZIP compression for ZIP-based formats
     * @default false
     */
    compression?: boolean;
}

export interface WorkBook {
    /**
     * A dictionary of the worksheets in the workbook.
     * Use SheetNames to reference these.
     */
    Sheets: { [sheet: string]: WorkSheet };

    /**
     * ordered list of the sheet names in the workbook
     */
    SheetNames: string[];

    /**
     * an object storing the standard properties. wb.Custprops stores custom properties.
     * Since the XLS standard properties deviate from the XLSX standard, XLS parsing stores core properties in both places.
     */
    Props?: Properties;

    Workbook?: WBProps;
}

export interface WBProps {
    Sheets?: any[];
}

export interface ColInfo {
    /**
     * Excel's "Max Digit Width" unit, always integral
     */
    MDW?: number;
    /**
     * width in Excel's "Max Digit Width", width*256 is integral
     */
    width?: number;
    /**
     * width in screen pixels
     */
    wpx?: number;
    /**
     * intermediate character calculation
     */
    wch?: number;
    /**
     * if true, the column is hidden
     */
    hidden?: boolean;
}
export interface RowInfo {
    /**
     * height in screen pixels
     */
    hpx?: number;
    /**
     * height in points
     */
    hpt?: number;
    /**
     * if true, the column is hidden
     */
    hidden?: boolean;
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

/**
 * object representing any sheet (worksheet or chartsheet)
 */
export interface Sheet {
    '!ref'?: string;
    '!margins'?: {
        left: number,
        right: number,
        top: number,
        bottom: number,
        header: number,
        footer: number,
    };
}

/**
 * object representing the worksheet
 */
export interface WorkSheet extends Sheet {
    [cell: string]: CellObject | any;
    '!cols'?: ColInfo[];
    '!rows'?: RowInfo[];
    '!merges'?: Range[];
    '!protect'?: ProtectInfo;
    '!autofilter'?: {ref: string};
}

/**
 * The Excel data type for a cell.
 * b Boolean, n Number, e error, s String, d Date
 */
export type ExcelDataType = 'b' | 'n' | 'e' | 's' | 'd' | 'z';

export interface CellObject {
    /**
     * The raw value of the cell.
     */
    v: string | number | boolean | Date;

    /**
     * Formatted text (if applicable)
     */
    w?: string;

    /**
     * The Excel Data Type of the cell.
     * b Boolean, n Number, e error, s String, d Date
     */
    t: ExcelDataType;

    /**
     * Cell formula (if applicable)
     */
    f?: string;

    /**
     * Range of enclosing array if formula is array formula (if applicable)
     */
    F?: string;

    /**
     * Rich text encoding (if applicable)
     */
    r?: string;

    /**
     * HTML rendering of the rich text (if applicable)
     */
    h?: string;

    /**
     * Comments associated with the cell **
     */
    c?: string;

    /**
     * Number format string associated with the cell (if requested)
     */
    z?: string;

    /**
     * Cell hyperlink object (.Target holds link, .tooltip is tooltip)
     */
    l?: object;

    /**
     * The style/theme of the cell (if applicable)
     */
    s?: object;
}

export interface CellAddress {
    /** Column number */
    c: number;
    /** Row number */
    r: number;
}

export interface Range {
    /** Starting cell */
    s: CellAddress;
    /** Ending cell */
    e: CellAddress;
}

export interface Utils {
    /* --- Cell Address Utilities --- */

    /** converts an array of arrays of JS data to a worksheet. */
    aoa_to_sheet<T>(data: T[], opts?: any): WorkSheet;

    /** Converts a worksheet object to an array of JSON objects */
    sheet_to_json<T>(worksheet: WorkSheet, opts?: {
        raw?: boolean;
        range?: any;
        header?: "A"|number|string[];
    }): T[];

    /** Generates delimiter-separated-values output */
    sheet_to_csv(worksheet: WorkSheet, options?: { FS: string, RS: string }): string;

    /** Generates a list of the formulae (with value fallbacks) */
    sheet_to_formulae(worksheet: WorkSheet): any;

    /* --- Cell Address Utilities --- */

    /** Converts 0-indexed cell address to A1 form */
    encode_cell(cell: CellAddress): string;

    /** Converts 0-indexed row to A1 form */
    encode_row(row: number): string;

    /** Converts 0-indexed column to A1 form */
    encode_col(col: number): string;

    /** Converts 0-indexed range to A1 form */
    encode_range(s: CellAddress, e: CellAddress): string;

    /** Converts A1 cell address to 0-indexed form */
    decode_cell(address: string): CellAddress;

    /** Converts A1 row to 0-indexed form */
    decode_row(row: string): number;

    /** Converts A1 column to 0-indexed form */
    decode_col(col: string): number;

    /** Converts A1 range to 0-indexed form */
    decode_range(range: string): Range;
}
