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
  CLng(Inputbox("Welches Getr�nk? 1 f�r Tee, " & _
  "2 f�r Kaffee oder 3 f�r Wasser", _
  "Getr�nk w�hlen")) - 1
varSprache = _
  CLng(Inputbox("Welche Sprache? 1 f�r Deutsch " & _
  "oder 2 f�r Englisch?", _
  "Sprache w�hlen")) - 1
MsgBox varDrink(varNummer, varSprache), 0, _
  "Datenfeldvariable ausgeben"
