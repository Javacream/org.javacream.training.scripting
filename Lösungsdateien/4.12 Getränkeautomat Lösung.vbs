Option Explicit
Dim varApp
Dim varPreis
Dim varDrink
Dim varEingabe

varApp = "Getr�nkeautomat"

' BEGR�SSUNG
MsgBox "Herzlich Willkommen! Bestellen Sie das Getr�nk Ihrer Wahl.", _
  vbInformation, varApp

' GETR�NKEANGEBOT VORSTELLEN
MsgBox "Folgende Getr�nke sind verf�gbar:" & vbCrLf & _
  vbCrLf & _
  "Orangensaft" & vbCrLf & _
  "Apfelschorle" & vbCrLf & _
  "Mineralwasser",vbInformation, varApp

' ABFRAGE GETR�NKEWUNSCH
varDrink = InputBox("Ihr Getr�nkewunsch?", varApp)

' PREIS "BERECHNEN"
Select Case varDrink
  Case "Orangensaft"
    varPreis = 0.75
  Case "Apfelschorle"
    varPreis = 0.5   
  Case "Mineralwasser"
    varPreis = 0.3
End Select

' GELDEINWURF
varEingabe = InputBox("Geldeinwurf " & (varPreis * 100) & " Cent", varApp)
Do Until Clng(varEingabe) = Clng(varPreis * 100) 
  varEingabe = InputBox("Geldeinwurf " & (varPreis * 100) & " Cent", varApp)
Loop 

' GETR�NKEAUSGABE
Msgbox "Vielen Dank! Hier kommt Ihr " & varDrink & ".", vbInformation, varApp


