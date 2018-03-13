/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
package com.sheetjs;

import java.lang.Integer;
import java.util.Scanner;
import java.io.IOException;
import java.io.File;
import org.mozilla.javascript.Context;
import org.mozilla.javascript.Function;
import org.mozilla.javascript.NativeObject;
import org.mozilla.javascript.Scriptable;

public class SheetJS {
  public Scriptable scope;
  public Context cx;
  public NativeObject nXLSX;

  public SheetJS() throws Exception {
    this.cx = Context.enter();
    this.scope = this.cx.initStandardObjects();

    /* boilerplate */
    cx.setOptimizationLevel(-1);
    String s = "var global = (function(){ return this; }).call(null);";
    cx.evaluateString(scope, s, "<cmd>", 1, null);

    /* eval library */
    s = new Scanner(SheetJS.class.getResourceAsStream("/xlsx.full.min.js")).useDelimiter("\\Z").next();
    //s = new Scanner(new File("xlsx.full.min.js")).useDelimiter("\\Z").next();
    cx.evaluateString(scope, s, "<cmd>", 1, null);

    /* grab XLSX variable */
    Object XLSX = scope.get("XLSX", scope);
    if(XLSX == Scriptable.NOT_FOUND) throw new Exception("XLSX not found");
    this.nXLSX = (NativeObject)XLSX;
  }

  public SheetJSFile read_file(String filename) throws IOException, ObjectNotFoundException {
    /* open file */
    String d = JSHelper.read_file(filename);

    /* options argument */
    NativeObject q = (NativeObject)this.cx.evaluateString(this.scope, "q = {'type':'binary', 'WTF':1};", "<cmd>", 2, null);

    /* set up function arguments */
    Object args[] = {d, q};

    /* call read -> wb workbook */
    Function readfunc = (Function)JSHelper.get_object("XLSX.read",this.scope);
    NativeObject wb = (NativeObject)readfunc.call(this.cx, this.scope, this.nXLSX, args);

    return new SheetJSFile(wb, this);
  }

  public static void close() { JSHelper.close(); }
}

