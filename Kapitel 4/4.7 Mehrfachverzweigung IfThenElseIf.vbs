Option Explicit
Dim varGetraenk
Dim varPreis
Dim varVorhanden

varGetraenk = _
  Inputbox("Bitte geben Sie den Getr�nkewunsch ein.", _
    "Getr�nk w�hlen")
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
  Msgbox "Dieses Getr�nk ist nicht verf�gbar!", _
    vbExclamation, "Achtung"
End If
