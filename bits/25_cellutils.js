/* XLS ranges enforced */
function shift_cell_xls(cell, tgt) {
	if(tgt.s) {
		if(cell.cRel) cell.c += tgt.s.c;
		if(cell.rRel) cell.r += tgt.s.r;
	} else {
		cell.c += tgt.c;
		cell.r += tgt.r;
	}
	cell.cRel = cell.rRel = 0;
	while(cell.c >= 0x100) cell.c -= 0x100;
	while(cell.r >= 0x10000) cell.r -= 0x10000;
	return cell;
}

function shift_range_xls(cell, range) {
	cell.s = shift_cell_xls(cell.s, range.s);
	cell.e = shift_cell_xls(cell.e, range.s);
	return cell;
}

