/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
import { Ptr, ProtoMessage, parse_shallow, parse_varint49, varint_to_i32 } from './proto';

interface IWAMessage {
	/** Metadata in .TSP.MessageInfo */
	meta: ProtoMessage;
	data: Uint8Array;
}
interface IWAArchiveInfo {
	id?: number;
	messages?: IWAMessage[];
}
export { IWAMessage, IWAArchiveInfo };

function parse_iwa(buf: Uint8Array): IWAArchiveInfo[] {
	var out: IWAArchiveInfo[] = [], ptr: Ptr = [0];
	while(ptr[0] < buf.length) {
		/* .TSP.ArchiveInfo */
		var len = parse_varint49(buf, ptr);
		var ai = parse_shallow(buf.slice(ptr[0], ptr[0] + len));
		ptr[0] += len;

		var res: IWAArchiveInfo = {
			id: varint_to_i32(ai[1][0].data),
			messages: []
		};
		ai[2].forEach(b => {
			var mi = parse_shallow(b.data);
			var fl = varint_to_i32(mi[3][0].data);
			res.messages.push({
				meta: mi,
				data: buf.slice(ptr[0], ptr[0] + fl)
			});
			ptr[0] += fl;
		});
		out.push(res);
	}
	return out;
}
export { parse_iwa };
