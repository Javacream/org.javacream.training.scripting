Option Explicit
Dim objFSO
Dim objFolder
Dim objFile
Dim varTemp

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\Dokumente und Einstellungen\Administrator.MMSOFT\Desktop\Integrata 5155\5155\5155_neu\")
For Each objFile in objFolder.Files
  varTemp = varTemp & objFile.Name & vbCrLf
Next
MsgBox varTemp, vbInformation, "Ordnerinhalt"
