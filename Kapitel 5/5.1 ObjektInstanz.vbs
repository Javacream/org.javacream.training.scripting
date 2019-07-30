Option Explicit

Class Drink

  Public Name
  Public Preis
  Public Menge
  Public Kalorienfrei
  Public Koffeinfrei
  Private EK

  Public Function entnehmeGetraenk()
    If Menge > 0 Then
      Menge = Menge - 1
      entnehmeGetraenk = True
    Else
      entnehmeGetraenk = False
    End If
  End Function

End Class

Dim objDrink
Dim varCheckEntnahme

' HIER INSTANZIEREN WIR DAS OBJEKT OBJDRINK...
Set objDrink = New Drink

' ...UND SETZEN EIGENSCHAFTEN FÜR DAS OBJEKT
objDrink.Name = "Mineralwasser"
objDrink.Preis = 0.3
objDrink.Koffeinfrei = True
objDrink.Kalorienfrei = True
objDrink.Menge = 25

' WIR SIMULIEREN EINE GETRÄNKEENTNAHME...
varCheckEntnahme = objDrink.entnehmeGetraenk()

' ...UND LASSEN UND DIE VERBLEIBENDE MENGE ANZEIGEN
Msgbox objDrink.Menge, vbInformation, "Verbleibende Menge"
