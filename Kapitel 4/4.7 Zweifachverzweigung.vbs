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
    vbExclamation, "Pr�fung Geldeingabe"
Else
  Msgbox "Vielen Dank! Das Getr�nk wird ausgegeben!", _
    vbInformation, "Getr�nkausgabe"
End If
