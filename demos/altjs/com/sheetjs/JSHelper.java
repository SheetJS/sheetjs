/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
package com.sheetjs;

import java.lang.Integer;
import java.lang.StringBuilder;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import org.mozilla.javascript.Context;
import org.mozilla.javascript.NativeArray;
import org.mozilla.javascript.NativeObject;
import org.mozilla.javascript.Scriptable;

public class JSHelper {
	static String read_file(String file) throws IOException {
		byte[] b = Files.readAllBytes(Paths.get(file));
		System.out.println(b.length);
		StringBuilder sb = new StringBuilder();
		for(int i = 0; i < b.length; ++i) sb.append(Character.toString((char)(b[i] < 0 ? b[i] + 256 : b[i])));
		return sb.toString();
	}

	static Object get_object(String path, Object base) throws ObjectNotFoundException {
		int idx = path.indexOf(".");
		Scriptable b = (Scriptable)base;
		if(idx == -1) return b.get(path, b);
		Object o = b.get(path.substring(0,idx), b);
		if(o == Scriptable.NOT_FOUND) throw new ObjectNotFoundException("not found: |" + path.substring(0,idx) + "|" + Integer.toString(idx));
		return get_object(path.substring(idx+1), (NativeObject)o);
	}

	static Object[] get_array(String path, Object base) throws ObjectNotFoundException {
		NativeArray arr = (NativeArray)get_object(path, base);
		Object[] out = new Object[(int)arr.getLength()];
		int idx;
		for(Object o : arr.getIds()) out[idx = (Integer)o] = arr.get(idx, arr); 
		return out;
	}

	static String[] get_string_array(String path, Object base) throws ObjectNotFoundException {
		NativeArray arr = (NativeArray)get_object(path, base);
		String[] out = new String[(int)arr.getLength()];
		int idx;
		for(Object o : arr.getIds()) out[idx = (Integer)o] = arr.get(idx, arr).toString(); 
		return out;
	}
	
	public static void close() { Context.exit(); }

}
