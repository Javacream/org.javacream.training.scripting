Option Explicit
Dim varGetraenk
Dim varPreis
Dim varVorhanden

varGetraenk = _
  Inputbox("Bitte geben Sie den Getränkewunsch ein.", _
    "Getränk wählen")
varVorhanden = True

Select Case varGetraenk
  Case "Kaffee", "Espresso", "Cappucino"
    varPreis = 0.5
  Case "Tee", "Schwarztee", "Früchtetee"
    varPreis = 0.4
  Case "Mineralwasser", "Wasser"
    varPreis = 0.3
  Case Else
    varVorhanden = False
End Select
If varVorhanden Then
  MsgBox "Bitte geben Sie " & varPreis * 100 & " Cent ein. ", _
    vbInformation, "Geldeingabe"
Else
  Msgbox "Dieses Getränk ist nicht verfügbar!", _
    vbExclamation, "Achtung"
End If
