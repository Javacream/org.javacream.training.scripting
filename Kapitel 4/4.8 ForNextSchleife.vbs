Option Explicit
Dim varGetraenk(3)
Dim varAusgabe
Dim i

varGetraenk(0) = "Kaffee"
varGetraenk(1) = "Tee"
varGetraenk(2) = "Mineralwasser"
varGetraenk(3) = "Apfelsaft"
varAusgabe = "Folgende Getr�nke sind verf�gbar:" & _
  vbCrLf & vbCrLf

For i = 0 To 3
  varAusgabe = varAusgabe & varGetraenk(i) & vbCrLf
Next
MsgBox varAusgabe, vbInformation, "Getr�nkeangebot"