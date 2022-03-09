function fix_hijri(date/*:Date*/, o/*:[number, number, number]*/) {
  /* TODO: properly adjust y/m/d and  */
  o[0] -= 581;
  var dow = date.getDay();
  if(date < 60) dow = (dow + 6) % 7;
  return dow;
}
