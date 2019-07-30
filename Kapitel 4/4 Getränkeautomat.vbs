Option Explicit
Dim varAppTitle
Dim varName
Dim varGetraenk
Dim varPreis
Dim varGeldeingabe
Dim varVorhanden

varAppTitle = "Marks Getr�nkeautomat"

' ABFRAGE DES ANWENDERNAMENS
varName = Inputbox("Bitte geben Sie Ihren Namen ein.", varAppTitle)

' GETR�NKEANGEBOT PR�SENTIEREN
Msgbox "Folgende Getr�nke sind verf�gbar:" & _
  vbCrLf & _
  vbCrLf & "Orangensaft" & _
  vbCrLf & "Apfelschorle" & _
  vbCrLf & "Mineralwasser", _
  vbInformation, varAppTitle

' GETR�NKEWUNSCH ABFRAGEN
varGetraenk = _
  Inputbox("Bitte geben Sie den Getr�nkewunsch ein.", _
    varAppTitle)
varVorhanden = True

Select Case varGetraenk
  Case "Orangensaft"
    varPreis = 0.5
  Case "Apfelschorle"
    varPreis = 0.4
  Case "Mineralwasser"
    varPreis = 0.3
  Case Else
    varVorhanden = False
End Select
If varVorhanden Then
  ' BETRAG PR�FEN
  Do 
    varGeldeingabe = _
    CLng(Inputbox ("Bitte geben Sie " & varPreis * 100 & " Cent ein. ", _
    varAppTitle)) / 100
  Loop While varGeldeingabe < varPreis
  MsgBox "Das Getr�nk wird ausgegeben! Vielen Dank!", _
    vbInformation, varAppTitle
Else
  Msgbox "Dieses Getr�nk ist nicht verf�gbar!", _
    vbExclamation, varAppTitle
End If

