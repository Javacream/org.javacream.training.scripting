Option Explicit
Dim varPreis
Dim varGeldeingabe

varPreis = 0.5
varGeldeingabe = 0.4

If varGeldeingabe < varPreis Then
  MsgBox "Sie haben nicht genug Geld eingegeben." & _
    vbCrLf & "Bitte geben Sie genau " & _
    FormatCurrency(varPreis) & " ein!", _
    vbExclamation, _
    "Prüfung Geldeingabe"
End If
