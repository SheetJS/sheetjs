function make_vba_xls(cfb/*:CFBContainer*/) {
	var newcfb = CFB.utils.cfb_new({root:"R"});
	cfb.FullPaths.forEach(function(p, i) {
		if(p.slice(-1) === "/" || !p.match(/_VBA_PROJECT_CUR/)) return;
		var newpath = p.replace(/^[^/]*/,"R").replace(/\/_VBA_PROJECT_CUR\u0000*/, "");
		CFB.utils.cfb_add(newcfb, newpath, cfb.FileIndex[i].content);
	});
	return CFB.write(newcfb);
}
