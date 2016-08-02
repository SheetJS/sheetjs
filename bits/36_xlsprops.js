function xlml_set_prop(Props, tag, val) {
	/* TODO: Normalize the properties */
	switch(tag) {
		case 'Description': tag = 'Comments'; break;
	}
	Props[tag] = val;
}

