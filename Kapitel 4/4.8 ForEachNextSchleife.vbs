Option Explicit
Dim varGetraenk(3)
Dim varElement
Dim varAusgabe
Dim i

varGetraenk(0) = "Kaffee"
varGetraenk(1) = "Tee"
varGetraenk(2) = "Mineralwasser"
varGetraenk(3) = "Apfelsaft"
varAusgabe = "Folgende Getr�nke sind verf�gbar:" & _
  vbCrLf & vbCrLf

For Each varElement in varGetraenk
  varAusgabe = varAusgabe & varElement & vbCrLf
Next
MsgBox varAusgabe, vbInformation, "Getr�nkeangebot"