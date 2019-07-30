Option Explicit
Dim varApp
Dim varPreis
Dim varDrink
Dim varEingabe

varApp = "Getränkeautomat"

' BEGRÜSSUNG
MsgBox "Herzlich Willkommen! Bestellen Sie das Getränk Ihrer Wahl.", _
  vbInformation, varApp

' GETRÄNKEANGEBOT VORSTELLEN
MsgBox "Folgende Getränke sind verfügbar:" & vbCrLf & _
  vbCrLf & _
  "Orangensaft" & vbCrLf & _
  "Apfelschorle" & vbCrLf & _
  "Mineralwasser",vbInformation, varApp

' ABFRAGE GETRÄNKEWUNSCH
varDrink = InputBox("Ihr Getränkewunsch?", varApp)

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

' GETRÄNKEAUSGABE
Msgbox "Vielen Dank! Hier kommt Ihr " & varDrink & ".", vbInformation, varApp


