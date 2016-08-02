declare enum CFBEntryType { unknown, storage, stream, lockbytes, property, root }
declare enum CFBStorageType { fat, minifat }

/* CFB Entry Object demanded by write functions */
interface CFBEntryMin {
  
  /* Raw Content (Buffer when available, Array of bytes otherwise) */
  content:any;
}

/* CFB Entry Object returned by parse functions */
interface CFBEntry extends CFBEntryMin {
  
  /* Case-sensitive internal name */
  name:string;
  
  /* CFB type (salient types: stream, storage) -- see CFBEntryType */
  type:string;
  
  /* Creation Time */
  ct:Date;
  /* Modification Time */
  mt:Date;


  /* Raw creation time -- see [MS-DTYP] 2.3.3 FILETIME */
  mtime:string;
  /* Raw modification time -- see [MS-DTYP] 2.3.3 FILETIME */
  ctime:string;

  /* RBT color: 0 = red, 1 = black */
  color:number;

  /* Class ID represented as hex string */
  clsid:string;

  /* User-Defined State Bits */
  state:number;

  /* Starting Sector */
  start:number;

  /* Data Size */
  size:number;
  
  /* Storage location -- see CFBStorageType */
  storage:string;
}


/* cfb.FullPathDir as demanded by write functions */
interface CFBDirectoryMin {

  /* keys are unix-style paths */
  [key:string]: CFBEntryMin;
}

/* cfb.FullPathDir Directory object */
interface CFBDirectory extends CFBDirectoryMin {

  /* cfb.FullPathDir keys are paths; cfb.Directory keys are file names */
  [key:string]: CFBEntry;
}


/* cfb object demanded by write functions */
interface CFBContainerMin {

  /* Path -> CFB object mapping */
  FullPathDir:CFBDirectoryMin;
}

/* cfb object returned by read and parse functions */
interface CFBContainer extends CFBContainerMin {

  /* search by path or file name */
  find(string):CFBEntry;

  /* list of streams and storages */
  FullPaths:string[];

  /* Path -> CFB object mapping */
  FullPathDir:CFBDirectory;

  /* Array of entries in the same order as FullPaths */
  FileIndex:CFBEntry[];

  /* Raw Content, in chunks (Buffer when available, Array of bytes otherwise) */
  raw:any[];
}


interface CFB {
  read(f:any, options:any):CFBContainer;
  parse(f:any):CFBContainer;
  utils: {
    ReadShift(size:any,t?:any):any;
    WarnField(hexstr:string,fld?:string);
    CheckField(hexstr:string,fld?:string);
    prep_blob(blob:any, pos?:number):any;
    bconcat(bufs:any[]):any;
  };
  main;
}
