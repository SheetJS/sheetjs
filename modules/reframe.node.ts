/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
import { decompress_iwa_file, compress_iwa_file, u8concat, parse_iwa_file, write_iwa_file, varint_to_i32, parse_shallow, write_shallow, u8str, stru8, IWAArchiveInfo, parse_TSP_Reference, write_varint49, u8contains, u8_to_dataview, write_new_storage } from "./src/numbers";
import { read, writeFile, utils, find, CFB$Entry, CFB$Container } from 'cfb';

var f = process.argv[2]; var o = process.argv[3];
var cfb = read(f, {type: "file"});

var nuevo = utils.cfb_new();

interface DependentInfo {
  deps: number[];
  location: string;
  type: number;
}
var dependents: {[x:number]: DependentInfo} = {};
var indices: number[] = [];

/* First Pass: reframe, clean up junk, collect message space */
cfb.FileIndex.map((fi, idx): [CFB$Entry, string] => ([fi, cfb.FullPaths[idx]])).forEach(row => {
  var fi = row[0], fp = row[1];
  /* blank all plist files */
  if(fi.name.match(/\.plist/)) {
    console.error(`Blanking plist ${fi.name}`);
    fi.content = new Uint8Array([
      0x62, 0x70, 0x6c, 0x69, 0x73, 0x74, 0x30, 0x30, 0xd0, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
      0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x09, 0x0a
    ]);
    utils.cfb_add(nuevo, row[1], fi.content);
    return;
  }

  if(fi.type != 2) return;

  /* Remove other metadata */
  if(!fi.name.match(/\.iwa/)) {
    if(fi.name.match(/Sh33tJ5/)) return;
    console.error(`Removing file ${fi.name}`);
    return;
  }

  /* Reframe .iwa files */
  var old_content = fi.content;
  var raw1 = decompress_iwa_file(old_content as Uint8Array);
  var new_content = compress_iwa_file(raw1);
  var raw2 = decompress_iwa_file(new_content);
  for(var i = 0; i < raw1.length; ++i) if(raw1[i] != raw2[i]) throw new Error(`${fi.name} did not properly roundtrip`);

  var x = parse_iwa_file(raw2);
  x.forEach(ia => {
    ia.messages.forEach(m => {
      delete m.meta[2];
      delete m.meta[4];
      delete m.meta[5]; // extremely slow open if deleted
      for(var j = 6; j < m.meta.length; ++j) delete m.meta[j];
    });
  });

  x.forEach(packet => {
    indices.push(packet.id);
    dependents[packet.id] = { deps: [], location: fp, type: varint_to_i32(packet.messages[0].meta[1][0].data) };
  });

  var y = write_iwa_file(x);
  for(var i = 0; i < raw1.length; ++i) if(y[i] != raw1[i]) { console.log(fi.name, i, raw1[i], y[i]); break; }

  var raw3 = compress_iwa_file(y);
  fi.content = raw3; fi.size = fi.content.length;
  // utils.cfb_add(nuevo, row[1], fi.content);
});

indices.sort((x,y) => x-y);
var indices_varint: Array<[number, Uint8Array]> = indices.filter(x => x > 1).map(x => [x, write_varint49(x)] );

/* Second pass: build dependency map */
cfb.FileIndex.map((fi, idx): [CFB$Entry, string] => ([fi, cfb.FullPaths[idx]])).forEach(row => {
  var fi = row[0], fp = row[1];
  if(!fi?.name?.match(/\.iwa/)) return;
  var x = parse_iwa_file(decompress_iwa_file(fi.content as Uint8Array));

  x.forEach(ia => {
    ia.messages.forEach(m => {
      indices_varint.forEach(([i, vi]) => {
        if(ia.messages.some(mess => varint_to_i32(mess.meta[1][0].data) != 11006 && u8contains(mess.data, vi))) {
          dependents[i].deps.push(ia.id);
        }
      })
    });
  });
});

var deletables = [];
function delete_message(id: number, cfb: CFB$Container) {
  var dep = dependents[id];
  if(dep?.deps?.length > 0) return console.error(`Message ${id} has dependents ${dep.deps}`);

  //delete dependents[id];
  indices = indices.filter(x => x != id);

  /* TODO: this really should be a forward map */
  indices.map(i => dependents[i]).filter(x => x).forEach(dep => {
    dep.deps = dep.deps.filter(x => x != id);
  });
  if(deletables.indexOf(id) == -1) deletables.push(id);
}

function gc(cfb: CFB$Container) {
  deletables.forEach(id => {
    var dep = dependents[id];
    var entry = find(cfb, dep.location);
    var x = parse_iwa_file(decompress_iwa_file(entry.content as Uint8Array));

    /* remove packet */
    console.log(`Removing packet ${id} ${dependents[id]?.type || "NaN"}`);
    for(var xi = 0; xi < x.length; ++xi) {
      var packet = x[xi];
      if(packet.id == id) { x.splice(xi,1); --xi; }
    }

    delete dependents[id];

    entry.content = compress_iwa_file(write_iwa_file(x));
    entry.size = entry.content.length;
  });
  deletables = [];

}

function delete_ref(dmeta: ReturnType<typeof parse_shallow>, j: number, me: number) {
  if(!dmeta) return;
  if(dmeta[j]) dmeta[j].forEach(pi => {
    var target = parse_TSP_Reference(pi.data);
    dependents[target].deps = dependents[target].deps.filter(x => x != me);
    if(dependents[target].deps.length == 0) delete_message(target, cfb);
  });
  delete dmeta[j];
}

function shake() {
  var done = false;
  while(!done) {
    done = true;
    cfb.FileIndex.map((fi, idx): [CFB$Entry, string] => ([fi, cfb.FullPaths[idx]])).forEach(row => {
      var fi = row[0], fp = row[1];
      if(fi.name.match(/\.plist/)) return;
      if(!fi.name.match(/\.iwa/)) return;
      var old_content = fi.content;
      var raw1 = decompress_iwa_file(old_content as Uint8Array);
      var x = parse_iwa_file(raw1);

      (()=>{
        for(var i = 0; i < x.length; ++i) {
          var w = x[i];
          var type = varint_to_i32(w.messages[0].meta[1][0].data);
          if(w.id > 10000 && (!dependents[w.id] || !dependents[w.id]?.deps?.length)) {
            console.log(`Deleting orphan ${fp} ${w.id}`);
            delete_message(w.id, cfb);
            done = false;
          }
        }
      })();

      var y = write_iwa_file(x);
      var raw3 = compress_iwa_file(y);
      //fi.content = raw3;
    });
    gc(cfb);
  }
}

/* Third pass: trim messages */
cfb.FileIndex.map((fi, idx): [CFB$Entry, string] => ([fi, cfb.FullPaths[idx]])).forEach(row => {
  var fi = row[0], fp = row[1];
  if(!fi?.name?.match(/\.iwa/)) return;
  var x = parse_iwa_file(decompress_iwa_file(fi.content as Uint8Array));

  trim(x, fi);

  var y = write_iwa_file(x);
  var raw3 = compress_iwa_file(y);
  fi.content = raw3;
});
gc(cfb);

shake();

function mutate_row(tri: ReturnType<typeof parse_shallow>) {
  if(!tri[6]?.[0] || !tri[7]?.[0]) throw "Mutation only works on post-BNC storages!";
	var wide_offsets = tri[8]?.[0]?.data && varint_to_i32(tri[8][0].data) > 0 || false;
  if(wide_offsets) throw "Math only works with normal offsets";
  var dv = u8_to_dataview(tri[7][0].data);
  var old_sz = 0, sz = 0;
  for(var i = 0; i < tri[7][0].data.length / 2; ++i) {
    sz = dv.getUint16(i*2, true);
    if(sz < 65535) old_sz = sz;
    if(old_sz) break;
  }
  if(!old_sz) old_sz = sz = tri[6][0].data.length;
  var start = 0,
    preamble = tri[6][0].data.slice(0, start),
    intramble = tri[6][0].data.slice(start, old_sz),
    postamble = tri[6][0].data.slice(old_sz);
  var sst = []; sst[69] = "SheetJS";
  //intramble = write_new_storage({t:"n", v:12345}, sst);
  //intramble = write_new_storage({t:"b", v:false}, sst);
  intramble = write_new_storage({t:"s", v:"SheetJS"}, sst);
  tri[6][0].data = u8concat([preamble, intramble, postamble]);
  var delta = intramble.length - old_sz;
  for(var i = 0; i < tri[7][0].data.length / 2; ++i) {
    sz = dv.getUint16(i*2, true);
    if(sz < 65535 && sz > start) dv.setUint16(i*2, sz + delta, true);
  }
}

/* Find first sheet -> first table -> add "SheetJS" to data store and set cell A1 */
(function() {
  var entry = find(cfb, dependents[1].location);
  var x = parse_iwa_file(decompress_iwa_file(entry.content as Uint8Array));
  var docroot: IWAArchiveInfo;
  for(var xi = 0; xi < x.length; ++xi) {
    var packet = x[xi];
    if(packet.id == 1) docroot = packet;
  }
  var sheetrootref = parse_TSP_Reference(parse_shallow(docroot.messages[0].data)[1][0].data);

  entry = find(cfb, dependents[sheetrootref].location);
  x = parse_iwa_file(decompress_iwa_file(entry.content as Uint8Array));
  for(xi = 0; xi < x.length; ++xi) {
    var packet = x[xi];
    if(packet.id == sheetrootref) docroot = packet;
  }
  sheetrootref = parse_TSP_Reference(parse_shallow(docroot.messages[0].data)[2][0].data);

  entry = find(cfb, dependents[sheetrootref].location);
  x = parse_iwa_file(decompress_iwa_file(entry.content as Uint8Array));
  for(xi = 0; xi < x.length; ++xi) {
    var packet = x[xi];
    if(packet.id == sheetrootref) docroot = packet;
  }
  sheetrootref = parse_TSP_Reference(parse_shallow(docroot.messages[0].data)[2][0].data);

  entry = find(cfb, dependents[sheetrootref].location);
  x = parse_iwa_file(decompress_iwa_file(entry.content as Uint8Array));
  for(xi = 0; xi < x.length; ++xi) {
    var packet = x[xi];
    if(packet.id == sheetrootref) docroot = packet;
  }

  var pb = parse_shallow(docroot.messages[0].data);
  {
    var store = parse_shallow(pb[4][0].data);
    {
      /* ref to string table */
      var sstref = parse_TSP_Reference(store[4][0].data);
      (() => {
        var sentry = find(cfb, dependents[sstref].location);
        var sx = parse_iwa_file(decompress_iwa_file(sentry.content as Uint8Array));
        var sstroot: IWAArchiveInfo;
        for(var sxi = 0; sxi < sx.length; ++sxi) {
          var packet = sx[sxi];
          if(packet.id == sstref) sstroot = packet;
        }

        var sstdata = parse_shallow(sstroot.messages[0].data);
        {
          if(!sstdata[3]) sstdata[3] = [];
          var newsst: ReturnType<typeof parse_shallow> = [];
          newsst[1] = [ { type: 0, data: write_varint49(69) } ];
          newsst[2] = [ { type: 0, data: write_varint49(1) } ];
          newsst[3] = [ { type: 2, data: stru8("SheetJS") } ];
          sstdata[3].push({type: 2, data: write_shallow(newsst)});
        }
        sstroot.messages[0].data = write_shallow(sstdata);

        var sy = write_iwa_file(sx);
        var raw3 = compress_iwa_file(sy);
        sentry.content = raw3; sentry.size = sentry.content.length;
      })();

      var tile = parse_shallow(store[3][0].data); // TileStorage
      {
        var t = tile[1][0];
        delete tile[2];
        var tl = parse_shallow(t.data); // first Tile
        {
          var tileref = parse_TSP_Reference(tl[2][0].data);
          (() => {
            var tentry = find(cfb, dependents[tileref].location);
            var tx = parse_iwa_file(decompress_iwa_file(tentry.content as Uint8Array));
            var tileroot: IWAArchiveInfo;
            for(var sxi = 0; sxi < tx.length; ++sxi) {
              var packet = tx[sxi];
              if(packet.id == tileref) tileroot = packet;
            }

            // .TST.Tile
            var tiledata = parse_shallow(tileroot.messages[0].data);
            {
              tiledata[5].forEach((row, R) => {
                var tilerow = parse_shallow(row.data);
                if(R == 0) mutate_row(tilerow);
                row.data = write_shallow(tilerow);
              })
            }
            tileroot.messages[0].data = write_shallow(tiledata);

            var ty = write_iwa_file(tx);
            var raw3 = compress_iwa_file(ty);
            tentry.content = raw3; tentry.size = tentry.content.length;
            //throw dependents[tileref];
          })();
        }
        t.data = write_shallow(tl);
      }
      store[3][0].data = write_shallow(tile);
    }
    pb[4][0].data = write_shallow(store);
  }
  docroot.messages[0].data = write_shallow(pb);

  var y = write_iwa_file(x);
  var raw3 = compress_iwa_file(y);
  entry.content = raw3; entry.size = entry.content.length;
})();

//console.log(indices.map(i => [i, dependents[i]]).filter(([i, dep]) => dep?.deps && (dep.deps.length == 0 || dep.deps.length == 1 && dep.deps[0] == 1)));

/* Last pass: fix metadata and commit to new file */
cfb.FileIndex.map((fi, idx): [CFB$Entry, string] => ([fi, cfb.FullPaths[idx]])).forEach(row => {
  var fi = row[0], fp = row[1];
  if(!fi?.name?.match(/\.iwa/)) return;
  var x = parse_iwa_file(decompress_iwa_file(fi.content as Uint8Array));

  if(fi.name.match(/^Metadata.iwa$/)) {
		x.forEach((w: IWAArchiveInfo) => {
			var type = varint_to_i32(w.messages[0].meta[1][0].data);
			if(type != 11006) return;
      var package_metadata = parse_shallow(w.messages[0].data);
      [3,11].forEach(x => {
        if(!package_metadata[x]) return;
        for(var pmj = 0; pmj < package_metadata[x].length; ++pmj) {
          var ci = package_metadata[x][pmj];
          var comp = parse_shallow(ci.data);
          var compid = varint_to_i32(comp[1][0].data);
          if(!dependents[compid]) {
            console.log(`Removing ${compid} (${u8str(comp[2][0].data)} -> ${comp[3]?.[0] && u8str(comp[3][0].data) || ""}) from metadata`);
            package_metadata[x].splice(pmj, 1);
            --pmj; continue;
          }
          [13, 14, 15].forEach(j => delete comp[j]);
          ci.data = write_shallow(comp);
        }
      });
      [2, 4, 6, 8, 9].forEach(j => delete package_metadata[j]);
      w.messages[0].data = write_shallow(package_metadata);
		});
	}

  var y = write_iwa_file(x);
  var raw3 = compress_iwa_file(y);
  fi.content = raw3;
  if(x.length == 0) {
    console.log(`Deleting ${fi.name} (no messages)`);
    return;
  }
  if(fi.name.match(/^Metadata.iwa$/)) {
    console.error(`Reframing iwa ${fi.name} (${fi.size} -> ${fi.content.length})`);
    fi.size = fi.content.length;
    utils.cfb_add(nuevo, row[1], fi.content);
  } else {
    console.error(`Reframing iwa ${fi.name} (${fi.size} -> ${fi.content.length})`);
    fi.size = fi.content.length;
    utils.cfb_add(nuevo, row[1], fi.content);
  }
});

writeFile(nuevo, o, { fileType: "zip", compression: true });

function process_root(dmeta: ReturnType<typeof parse_shallow>, id: number) {
  var tsa = parse_shallow(dmeta[8][0].data);
  {
    var tsk = parse_shallow(tsa[1][0].data);
    {
      [7, 8, 14].forEach(j => delete_ref(tsk, j, id)); // annotation_author_storage
      [4, 9, 10, 11, 12, 15, 16, 17].forEach(j => delete tsk[j]);
    }
    tsa[1][0].data = write_shallow(tsk);
    [5, 6, 7, 10, 11, 12, 13].forEach(j => delete_ref(tsa, j, 1));
    [3, 8, 9, 14, 15, 16].forEach(j => delete tsa[j]);
  }
  dmeta[8][0].data = write_shallow(tsa);
}

function process_sheet(dmeta: ReturnType<typeof parse_shallow>, id: number) {
  [3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14, 20, 21, 23, 24].forEach(j => delete dmeta[j]);
  [15, 16, 17, 18, 19, 22].forEach(j => delete_ref(dmeta, j, id));
}

function process_TST_TableInfoArchive(dmeta: ReturnType<typeof parse_shallow>, id: number) {
  [7, 8, 9, 10, 14, 16].forEach(j => delete dmeta[j]);
  [3, 4, 5, 6, 15, 17].forEach(j => delete_ref(dmeta, j, id));
}

function process_TST_TableModelArchive(dmeta: ReturnType<typeof parse_shallow>, id: number) {
  [90, 91, 92].forEach(j => delete dmeta[j]);
  [34, 35, 38, 49].forEach(j => delete_ref(dmeta, j, id));
  [60, 61, 62, 63, 64, 65, 66, 67, 68, 69].forEach(j => delete_ref(dmeta, j, id));
  [70].forEach(j => delete dmeta[j]);
  [71, 72, 73, 74, 75, 76, 77, 78, 79, 80].forEach(j => delete_ref(dmeta, j, id));
  [85, 86, 87, 88, 89].forEach(j => delete_ref(dmeta, j, id));
  var store = parse_shallow(dmeta[4][0].data);
  { // .TST.DataStore
    [12, 13, 15, 16, 17, 18, 19, 20, 21, 22].forEach(j => delete_ref(store, j, id));
    var tiles = parse_shallow(store[3][0].data);
    { // .TST.TileStorage
      //console.log(parse_TSP_Reference(parse_shallow(tiles[1][0].data)[2][0].data));
    }
    store[3][0].data = write_shallow(tiles);
  //   4  -> sst
  //   17 -> rsst
  }
  dmeta[4][0].data = write_shallow(store);
}

function process_TN_ThemeArchive(dmeta: ReturnType<typeof parse_shallow>, id: number) {
  [2].forEach(j => delete_ref(dmeta, j, id));
  var tssta = parse_shallow(dmeta[1][0].data);
  {
    //Object.keys(tssta).filter(x => +x >= 100).forEach(j => delete tssta[j]);
  }
  dmeta[1][0].data = write_shallow(tssta);
}

function process_TSS_StylesheetArchive(dmeta: ReturnType<typeof parse_shallow>, id: number) {
  [7, 8, 9, 10, 11, 12].forEach(j => {
    if(!dmeta[j]) return;
    var vstyle = parse_shallow(dmeta[j][0].data);
    delete_ref(vstyle, 1, id);
    if(vstyle[2]) vstyle[2].forEach(ised => {
      var ise = parse_shallow(ised.data);
      delete_ref(ise, 2, id);
    })
    if(vstyle[3]) vstyle[3].forEach(sced => {
      var sce = parse_shallow(sced.data);
      delete_ref(sce, 2, id);
      delete_ref(sce, 1, id);
    })
    delete dmeta[j];
  });
  dmeta[4] = [{data: write_varint49(0), type: 0}];
  [3].forEach(j => delete_ref(dmeta, j, id));
  dmeta[5]?.forEach(pi => {
    var sce = parse_shallow(pi.data);
    delete_ref(sce, 1, id);
    delete_ref(sce, 2, id);
  }); delete dmeta[5];
  var deleted_styles = [];
  if(dmeta[2]) for(var pij = 0; pij < dmeta[2].length; ++pij) {
    var ise = parse_shallow(dmeta[2][pij].data);
    var sname = u8str(ise[1][0].data);
    if(sname.match(/_\d*$/) || sname.match(/^(captions|chart|image|movie|stickyComment)/) || sname.match(/evel[2-5]|pivot|liststyle/) ) {
      var deletedid = parse_TSP_Reference(ise[2][0].data);
      console.log(`Deleting style ${sname} [${deletedid}]`);
      deleted_styles.push(deletedid);
      delete_ref(ise, 2, id);
      dmeta[2].splice(pij,1);
      --pij; continue;
    } else console.log(`Keeping style ${sname}`)
  }
  if(dmeta[1]) for(var sj = 0; sj < dmeta[1].length; ++sj) {
    var dsid = parse_TSP_Reference(dmeta[1][sj].data);
    if(deleted_styles.indexOf(dsid) > -1) {
      //console.log(`Really deleting style ${dsid}`);
      dependents[dsid].deps = dependents[dsid].deps.filter(x => x != id);
      if(dependents[dsid].deps.length == 0) delete_message(dsid, cfb);
      dmeta[1].splice(sj, 1); --sj; continue;
    }
  }
  [6].forEach(j => delete dmeta[j]);
}

function process_TSCE_CalculationEngineArchive(dmeta: ReturnType<typeof parse_shallow>, id: number) {
  [3, 12, 14, 15].forEach(j => delete_ref(dmeta, j, id));
  [1, 4, 5, 6, 7, 9, 10, 11, 13, 16, 17, 18].forEach(i => delete dmeta[i]);
  var dta = parse_shallow(dmeta[2][0].data);
  delete dta[5]; delete dta[2]; delete dta[4];
  delete dta[1]; delete dta[3]; delete_ref(dta, 6, id);
  dmeta[2][0].data = write_shallow(dta);
}

function trim(x: IWAArchiveInfo[], fi: CFB$Entry) {
  for(var i = 0; i < x.length; ++i) {
    var w = x[i];
    var type = varint_to_i32(w.messages[0].meta[1][0].data);
    var dmeta = parse_shallow(w.messages[0].data);

    switch(type) {
      case 222: // .TSK.CustomFormatListArchive
      case 601: // .TSA.FunctionBrowserStateArchive
        x.splice(i, 1); --i; continue;
      case 3047: break; // .TSD.GuideStorageArchive
      case 205: {
      } break; // .TSK.TreeNode

      case 1: { // .TN.DocumentArchive
        process_root(dmeta, w.id);
      } break;
      case 2: { // .TN.SheetArchive
        process_sheet(dmeta, w.id);
      } break;
      case 6000: { // .TST.TableInfoArchive
        process_TST_TableInfoArchive(dmeta, w.id);
      } break;
      case 6001: { // .TST.TableModelArchive
        process_TST_TableModelArchive(dmeta, w.id)
      } break;
      case 401: { // .TSS.StylesheetArchive
        process_TSS_StylesheetArchive(dmeta, w.id);
        break;
      }
      case 11011: { // .TSP.DocumentMetadata
        [1, 3].forEach(i => delete dmeta[i]);
      } break;
      case 6002: {
        //process_TST_Tile(dmeta, w.id);
      } break; // .TST.Tile
      case 6005: break; // .TST.TableDataList
      case 6006: break; // .TST.HeaderStorageBucket
      case 2001: break; // .TSWP.StorageArchive
      case 12009: process_TN_ThemeArchive(dmeta, w.id); break; // .TN.ThemeArchive

      case 3097: break; // .TSD.StandinCaptionArchive
      case 4000: process_TSCE_CalculationEngineArchive(dmeta, w.id); break; // .TSCE.CalculationEngineArchive
      case 4003: break; // .TSCE.NamedReferenceManagerArchive
      case 4004: break; // .TSCE.TrackedReferenceStoreArchive
      case 4008: { // .TSCE.FormulaOwnerDependenciesArchive
        [11].forEach(j => delete_ref(dmeta, j, w.id));
        [3].forEach(i => delete dmeta[i]);
      } break;
      case 4009: { // .TSCE.CellRecordTileArchive
        delete dmeta[4];
      } break;
      case 6267: { // .TST.ColumnRowUIDMapArchive
        // console.log("CRUIDMA", dmeta[1]?.[0], dmeta[2]?.[0], dmeta[3]?.[0]);
        //[1,2,3,4,5,6].forEach(j => delete dmeta[j]); break;
      } break;
      case 6366: break; // .TST.HeaderNameMgrArchive
      case 6204: break; // .TST.HiddenStateFormulaOwnerArchive
      case 6220: break; // .TST.FilterSetArchive
      case 6305: break; // .TST.StrokeSidecarArchive
      case 6316: break; // .TST.SummaryModelArchive
      case 6317: break; // .TST.SummaryCellVendorArchive
      case 6318: break; // .TST.CategoryOrderArchive
      case 6372: break; // .TST.CategoryOwnerRefArchive
      case 6373: break; // .TST.GroupByArchive
      case 2043: [2,3,4].forEach(j => delete dmeta[j]); break; // .TSWP.NumberAttachmentArchive

      case 5020: // .TSCH.ChartStylePreset
        //[1,2,3,4,5,6].forEach(j => delete_ref(dmeta, j, w.id)); // some part of this is required to auto-save
        Object.keys(dmeta).filter(x => +x >= 10000).forEach(j => delete dmeta[j]);
        break;
      case 5022: // .TSCH.ChartStyleArchive
      case 5024: // .TSCH.LegendStyleArchive
      case 5026: // .TSCH.ChartAxisStyleArchive
      case 5028: // .TSCH.ChartSeriesStyleArchive
      case 5030: // .TSCH.ReferenceLineStyleArchive
        Object.keys(dmeta).filter(x => +x >= 10000).forEach(j => delete dmeta[j]);
        break;
      case 2021: break;
      case 2022: [10,11,12].forEach(j => delete dmeta[j]); break; // .TSWP.ParagraphStyleArchive
      case 2023: [10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25].forEach(j => delete dmeta[j]); break; // .TSWP.ListStyleArchive
      case 3016: [10,11].forEach(j => delete dmeta[j]); break; // .TSD.MediaStyleArchive
      case 2025: /* [10,11].forEach(j => delete dmeta[j]); */ break; // .TSWP.ShapeStyleArchive
      case 6003: /* [10,11].forEach(j => delete dmeta[j]); */ break; // .TST.TableStyleArchive
      case 6004: /* [10,11].forEach(j => delete dmeta[j]); */ break; // .TST.CellStyleArchive
      case 10024: [10,11,12].forEach(j => delete dmeta[j]); break; // .TN.DropCapStyleArchive
      case 12050: /* [2,3].forEach(j => delete dmeta[j]); */ break; // .TN.SheetStyleArchive
      case 3045: break; // .TSD.CanvasSelectionArchive

      case 6008: {
        [1].forEach(j => delete dmeta[j]);
        [2].forEach(j => delete_ref(dmeta, j, w.id));
        // deleting 3 -> -[TSWPLayoutManager initWithStorage:owner:] /Library/Caches/com.apple.xbs/Sources/iWorkDependenciesMacOS/iWorkDependenciesMacOS-7032.0.145/shared/text/TSWPLayoutManager.mm:82 Cannot initialize with a nil storage.
      } break; // .TST.TableStylePresetArchive

      case 6247: {
        [ 10, 11, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35 ].forEach(j => delete_ref(dmeta, j, w.id));
      } break; // .TST.TableStyleNetworkArchive

      case 210: { // .TSK.ViewStateArchive
        [2, 3].forEach(i => delete dmeta[i]);
      } break;
      case 12026: { // .TN.UIStateArchive
        delete dmeta[13];
      } break;

      case 219: delete dmeta[1]; break; // .DocumentSelectionArchive

      case 12028: break; // .TN.SheetSelectionArchive
      case 3061: { // .TSD.DrawableSelectionArchive
        [3].forEach(j => delete_ref(dmeta, j, w.id));
      } break;
      case 3091: { // .TSD.FreehandDrawingToolkitUIState
        [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15].forEach(i => delete dmeta[i]);
      } break;

      /* Metadata.iwa */
      case 11006: { // .TSP.PackageMetadata
        [10].forEach(j => delete_ref(dmeta, j, w.id));
      } break;
      case 11014: // .TSP.DataMetadata
        [1].forEach(j => delete dmeta[j]);
        break;
      case 11015: // .TSP.DataMetadataMap
        dmeta?.[1]?.forEach((r, idx) => {
          var ddmeta = parse_shallow(r.data);
          [2].forEach(j => delete_ref(ddmeta, j, w.id));
        });
        delete dmeta[1];
        dmeta.length = 0;
        break;

      /* AnnotationAuthorStorage.iwa */
      case 213: break; // TSK.AnnotationAuthorStorageArchive

      default:
        console.log("!!", fi.name, type, w.id, w.messages.length);
    }
    w.messages[0].data = write_shallow(dmeta);
  }
}