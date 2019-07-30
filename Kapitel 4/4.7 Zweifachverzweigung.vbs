Option Explicit
Dim varPreis
Dim varGeldeingabe

varPreis = 0.5
varGeldeingabe = _
  CDbl(Inputbox("Bitte Betrag in Cent einwerfen", _
    "Geldeinwurf")) / 100

If varGeldeingabe < varPreis Then
  MsgBox "Sie haben nicht genug Geld eingegeben." & _
    vbCrLf & "Bitte geben Sie genau " & _
    FormatCurrency(varPreis) & " ein!", _
    vbExclamation, "Prüfung Geldeingabe"
Else
  Msgbox "Vielen Dank! Das Getränk wird ausgegeben!", _
    vbInformation, "Getränkausgabe"
End If
