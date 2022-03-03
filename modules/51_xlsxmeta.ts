/// <reference path="src/types.ts"/>

RELS.XLMETA = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";

/* 12.3.10 Metadata Part */
function parse_xlmeta_xml(data: string, name: string, opts?: ParseXLMetaOptions): XLMeta {
	var out: XLMeta = { Types: [] };
	if(!data) return out;
	var pass = false;

	data.replace(tagregex, (x: string, idx: number) => {
		var y: any = parsexmltag(x);
		switch(strip_ns(y[0])) {
			case '<?xml': break;

			/* 18.9.8 */
			case '<metadata': case '</metadata>': break;

			/* 18.9.11 */
			case '<metadataTypes': case '</metadataTypes>': break;

			/* 18.9.10 */
			case '<metadataType':
				out.Types.push({ name: y.name });
				break;

			/* 18.9.4 */
			case '<futureMetadata': break;
			case '</futureMetadata>': break;

			/* 18.9.1 */
			case '<bk>': break;
			case '</bk>': break;

			/* 18.9.15 */
			case '<rc': break;
			case '</rc>': break;

			/* 18.9.3 */
			case '<cellMetadata':
			case '</cellMetadata>': break;

			/* 18.9.17 */
			case '<valueMetadata': break;
			case '</valueMetadata>': break;

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst': case '<extLst>': case '</extLst>': case '<extLst/>': break;

			/* 18.2.7  ext CT_Extension + */
			case '<ext': pass=true; break; //TODO: check with versions of excel
			case '</ext>': pass=false; break;

			default: if(!pass && opts.WTF) throw new Error('unrecognized ' + y[0] + ' in metadata');
		}
		return x;
	});
	return out;
}
/* TODO: coordinate with cell writing, pass flags */
function write_xlmeta_xml(): string {
	var o = [XML_HEADER];
	o.push(`\
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">
  <metadataTypes count="1">
    <metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>
  </metadataTypes>
  <futureMetadata name="XLDAPR" count="1">
    <bk>
      <extLst>
        <ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">
          <xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <cellMetadata count="1">
    <bk>
      <rc t="1" v="0"/>
    </bk>
  </cellMetadata>
</metadata>`);

	return o.join("");
}
