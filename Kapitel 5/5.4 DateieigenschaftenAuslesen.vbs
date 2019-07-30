Option Explicit
Dim objFSO
Dim objFile
Dim varTemp

'FILESYSTEMOBJECT ERSTELLEN 
Set objFSO = CreateObject("Scripting.FileSystemObject")

'AUF SPEZIELLE DATEI ZUGREIFEN
Set objFile = objFSO.GetFile("C:\Dokumente und Einstellungen\Administrator.MMSOFT\Desktop\Integrata 5155\Beispiele\Kapitel 5\5.5 Datenbankzugriff.vbs")

'DATEIEIGENSCHAFTEN ERMITTELN UND ANZEIGEN
varTemp = objFile.Path & vbCrLf
varTemp = varTemp & _
  "Erstellt: " & objFile.DateCreated & vbCrLf
varTemp = varTemp & _
  "Letzter Zugriff: " & objFile.DateLastAccessed & vbCrLf
varTemp = varTemp & _
  "Letzte Änderung: " & objFile.DateLastModified
MsgBox varTemp, vbInformation, "Dateieigenschaften"
