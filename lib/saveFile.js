function saveFile(fileType, fileName) {
  var objFSO = new ActiveXObject("Scripting.FileSystemObject")
  var objDialog = new ActiveXObject("SAFRCFileDlg.FileSave");
  objDialog.fileType=fileType;
  objDialog.fileName=fileName;
  var r = objDialog.OpenFileSaveDlg();
  strFileName = objDialog.FileName;
  if (strFileName.length>0) {
    objFile=objFSO.CreateTextFile(strFileName,true)
    objFile.Write (output.value);
    r = objFile.Close();
    btnSaveFile.style.display='none';
  }
}
