var IEBinaryToArray_ByteStr_Script =
   "<!-- IEBinaryToArray_ByteStr -->\r\n"+
   "<script type='text/vbscript'>\r\n"+
   "Function IEBinaryToArray_ByteStr(Binary) : IEBinaryToArray_ByteStr = CStr(Binary) : End Function\r\n"+
   "Function IEBinaryToArray_ByteStr_Last(Binary)\r\n"+
   "   Dim lastIndex\r\n"+
   "   lastIndex = LenB(Binary)\r\n"+
   "   if lastIndex mod 2 Then\r\n"+
   "       IEBinaryToArray_ByteStr_Last = Chr( AscB( MidB( Binary, lastIndex, 1 ) ) )\r\n"+
   "   Else\r\n"+
   "       IEBinaryToArray_ByteStr_Last = "+'""'+"\r\n"+
   "   End If\r\n"+
   "End Function\r\n"+
   "</script>\r\n";

document.write(IEBinaryToArray_ByteStr_Script);

