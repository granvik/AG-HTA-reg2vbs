function getCurrentDir(){
  var strHTALocation=location.href;
  ptnFile=/file\:\/\/\/(.*)/i;
  strHTALocation = strHTALocation.replace(ptnFile,'$1');
  ptnSlash=/\//gi;
  strHTALocation = strHTALocation.replace(ptnSlash,'\\');
  ptnSpace=/\%20/gi;
  strHTALocation = strHTALocation.replace(ptnSpace,' ');
  ptnEnd=/(.*\\).*/i;
  var strHTADir = strHTALocation.replace(ptnEnd,'$1');
  return strHTADir;
}
