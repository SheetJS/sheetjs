function xlml_set_prop(Props, tag/*:string*/, val) {
	/* TODO: Normalize the properties */
	switch(tag) {
		case 'Description': tag = 'Comments'; break;
	}
	Props[tag] = val;
}

