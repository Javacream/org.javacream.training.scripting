Option Explicit
Dim varApp
Dim varPreis
Dim varEingabe

' BEISPIEL EINES AUFRUFS EINER PROZEDUR
Sub GeldeinwurfPruefen(varPreis)
  MsgBox "Das gew�hlte Getr�nk kostet " & _
    FormatCurrency(varPreis) & ". Bitte werfen " & vbCrLf & _
    "Sie den richtigen Betrag ein!", vbExclamation, varApp
End Sub

varApp = "Getr�nkeautomat"
varPreis = 0.5
varEingabe = _
  Clng(InputBox("Werfen Sie bitte den Centbetrag ein", varApp)) 
If varEingabe / 100 < varPreis then
  Call GeldeinwurfPruefen(varPreis)
End If
