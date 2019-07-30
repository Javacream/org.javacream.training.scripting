'DIESE FUNKTION RECHNET EINE DOLLARBETRAG IN EURO UM
'UND GIBT DAS ERGEBNIS IN EINEM MELDUNGSFENSTER AUS
Option Explicit
Dim varTageskurs
Dim varEingabe

varTageskurs = 1.25
varEingabe = InputBox("Geben Sie den Dollarbetrag ein")
MsgBox FormatCurrency(DollarInEuro(varEingabe, varTageskurs))

Function DollarInEuro(varDollarBetrag, varTageskurs)
  DollarInEuro = varDollarBetrag * varTageskurs
End Function


