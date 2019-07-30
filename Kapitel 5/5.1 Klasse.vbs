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
