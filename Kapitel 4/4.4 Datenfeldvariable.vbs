Option Explicit
Dim varDrink(2,1)
Dim varSprache
Dim varNummer

varDrink(0,0) = "Tee"
varDrink(1,0) = "Kaffee"
varDrink(2,0) = "Wasser"
varDrink(0,1) = "Tea"
varDrink(1,1) = "Coffee"
varDrink(2,1) = "Water"
varNummer = _
  CLng(Inputbox("Welches Getränk? 1 für Tee, " & _
  "2 für Kaffee oder 3 für Wasser", _
  "Getränk wählen")) - 1
varSprache = _
  CLng(Inputbox("Welche Sprache? 1 für Deutsch " & _
  "oder 2 für Englisch?", _
  "Sprache wählen")) - 1
MsgBox varDrink(varNummer, varSprache), 0, _
  "Datenfeldvariable ausgeben"
