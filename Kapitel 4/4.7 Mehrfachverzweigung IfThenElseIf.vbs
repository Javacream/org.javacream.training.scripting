Option Explicit
Dim varGetraenk
Dim varPreis
Dim varVorhanden

varGetraenk = _
  Inputbox("Bitte geben Sie den Getränkewunsch ein.", _
    "Getränk wählen")
varVorhanden = True

If varGetraenk = "Kaffee" Then
  varPreis = 0.5
ElseIf varGetraenk = "Tee" Then
  varPreis = 0.4
ElseIf varGetraenk = "Mineralwasser" Then
  varPreis = 0.3
Else
  varVorhanden = False
End If
If varVorhanden Then
  MsgBox "Bitte geben Sie " & varPreis * 100 & " Cent ein. ", _
    vbInformation, "Geldeingabe"
Else
  Msgbox "Dieses Getränk ist nicht verfügbar!", _
    vbExclamation, "Achtung"
End If
