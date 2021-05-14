//********************************************************************
//*
//*  File:           convertRegfileToVbs.js
//*  Created:        2007-08-01
//*  Changed:        2009-08-19
//*  Version:        0.9.5
//*  Author:         Andrey Grevtsov 
//*                  http://aradm.ru
//*
//*  Description:    JavaScript with function convert from registry file to  
//*                  VB script using WMI
//*
//*  Copyleft (C):   2009 Andrey Grevtsov
//*
//*  License:        GPL
//*
//********************************************************************
function convertRegfileToVbs(registry_keys) {
  var output ='';
  output = 'const HKEY_CLASSES_ROOT = &H80000000\n';
  output += 'const HKEY_CURRENT_USER = &H80000001\n';
  output += 'const HKEY_LOCAL_MACHINE = &H80000002\n';
  output += 'QUOT = chr(34)\n';
  output += 'strComputer = "."\n';
  output += 'Set objRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\\\" &strComputer & "\\root\\default:StdRegProv")\n';
  var arrRegistryKeys = [];
  var arrRegistryKeys1 = [];
  var strSection ="";
  var strSectionDelete ="";
  ptnPrepare = /^\s*(.*)/;
  ptnSlashQuot = /\\\"/gi;
  ptnDoubleSlash = /\\\\/gi;  
  ptnContinueLine = /^(.*)\\$/;
  ptnComment = /^\;(.*)/;
  ptnSection = /^\[(\w*)\\(.*)\]$/;
  ptnSectionDelete = /^\[-(\w*)\\(.*)\]$/;
  ptnKeyDword = /^\"(.*)\"=dword\:(.*)/;
  ptnKeyString = /^\"(.*)\"=\"(.*)\"/;
  ptnHex = /^\"(.*)\"=hex\:(.*)/;  
  ptnHex2 = /^\"(.*)\"=hex\(2\)\:(.*)/;  
  ptnHex7 = /^\"(.*)\"=hex\(7\)\:(.*)/;
  ptnDefault = /^\@=(.*)/;
  ptnQuotTxt = /\+QUOT\+\"$/;
  
  arrRegistryKeys = registry_keys.split("\r\n");
  var ii=0;
  for (var i=0; i<arrRegistryKeys.length; i++) {
    arrRegistryKeys[i] = arrRegistryKeys[i].replace(ptnPrepare,'$1');
// объединяем строку, разделённую для удобочитаемости    
    if (arrRegistryKeys[i].match(ptnContinueLine)) {
      if (arrRegistryKeys1[ii]){
       arrRegistryKeys1[ii] += arrRegistryKeys[i].replace(ptnContinueLine,'$1') ;
      }
      else {
       arrRegistryKeys1[ii] = arrRegistryKeys[i].replace(ptnContinueLine,'$1') ;
      }
    }
    else {
      if (arrRegistryKeys1[ii] ) {
        arrRegistryKeys1[ii] += arrRegistryKeys[i];
      }
      else {
        arrRegistryKeys1[ii] = arrRegistryKeys[i];
      }
      ii++;
    }
  }  
  for (var i=0; i<arrRegistryKeys1.length; i++) {
// удаляем пробелы в начале и конце строки  
    strR = arrRegistryKeys1[i].replace(ptnPrepare,'$1');
    strR = strR.replace(ptnSlashQuot,'"+QUOT+"');
    strR = strR.replace(ptnDoubleSlash,'\\');
    strR = strR.replace(ptnQuotTxt,'');
    if (strR.match(ptnSectionDelete)) {
      strSectionDelete = strR.replace(ptnSectionDelete,'$1');
      strKeyPath = strR.replace(ptnSectionDelete,'$2');
      output += 'objRegistry.DeleteKey '+strSectionDelete+', "'+strKeyPath+'"\r\n';
    }
    if (strR.match(ptnSection)) {
      strSection = strR.replace(ptnSection,'$1');
      strKeyPath = strR.replace(ptnSection,'$2');
      output += 'objRegistry.CreateKey '+strSection+', "'+strKeyPath+'"\r\n';
    }
    if (strSection.length>0 && strR.match(ptnKeyDword)) {
      intDWORD =parseInt('0x'+strR.replace(ptnKeyDword,'$2'));
      output += strR.replace(ptnKeyDword,'objRegistry.SetDWORDValue '+strSection+',"'+strKeyPath+'","$1",'+intDWORD+'\r\n');
    }
    if (strSection.length>0 && strR.match(ptnKeyString)) {
      output += strR.replace(ptnKeyString,'objRegistry.SetStringValue '+strSection+',"'+strKeyPath+'","$1","$2"\r\n');      
    }
    if (strSection.length>0 && strR.match(ptnHex)) {
      strName = strR.replace(ptnHex,'$1');
      strValueString = strR.replace(ptnHex,'$2');
      arrValueString = strValueString.split(",");
      strSS = "";
      strS ="";
      for (var j=0; j<arrValueString.length; j++) {
        if (strS.length>0){
          strS = ',&h'+arrValueString[j];
        }
        else {
          strS = '&h'+arrValueString[j];
        }
        strSS += strS;
      }
      if (strValueString.length>0){
        output += 'objRegistry.SetBinaryValue '+strSection+',"'+strKeyPath+'","'+strName+'",array('+strSS+')\r\n';
      }  
    }
    if (strSection.length>0 && strR.match(ptnHex2)) {
      strName = strR.replace(ptnHex2,'$1');
      strValueString = strR.replace(ptnHex2,'$2');
      arrValueString = strValueString.split(",");
      strSS = "";
      for (var j=0; j<arrValueString.length; j=j+2) {
        if (arrValueString[j] != '00'){
          strS = String.fromCharCode('0x'+arrValueString[j]);
          strSS += strS;
        }
      }
      output += 'objRegistry.SetExpandedStringValue '+strSection+',"'+strKeyPath+'","'+strName+'","'+strSS+'"\r\n';
    }
    if (strSection.length>0 && strR.match(ptnHex7)) {
      strName = strR.replace(ptnHex7,'$1');
      strValueString = strR.replace(ptnHex7,'$2');
      arrValueString = strValueString.split(",");
      strSS = "";
      strS ="";
      for (var j=0; j<arrValueString.length; j=j+2) {
        if (arrValueString[j] != '00'){
          strS = String.fromCharCode('0x'+arrValueString[j]);
          strSS += strS;
        }
        else {
          if (arrValueString.length >= j+2){
            strSS += '","';
          }
        }
      }
      output += 'objRegistry.SetMultiStringValue '+strSection+',"'+strKeyPath+'","'+strName+'",array("'+strSS+'")\r\n';
    }
    if (strSection.length>0 && strR.match(ptnDefault)) {
      output.value += strR.replace(ptnDefault,'objRegistry.SetStringValue '+strSection+',"'+strKeyPath+'","",$1\r\n');      
    }            
  }
  return output;
}
