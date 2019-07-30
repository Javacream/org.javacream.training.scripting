Option Explicit

Class Drink
  Public Name
  Public Preis
  Public Menge
  Public Kalorienfrei
  Public Koffeinfrei
  Private EK
  Public ServiceEMail

  Public Function entnehmeGetraenk()
    If Menge > 0 Then
      Menge = Menge - 1
      entnehmeGetraenk = True
      If Menge < 3 Then sendeBestellung()
    Else
      entnehmeGetraenk = False
    End If
  End Function

  Private Function sendeBestellung()
    Dim objOLApp
    Dim objOLNameSpace
    Dim objOLMail
    
    Set objOLApp = CreateObject("Outlook.Application")
    Set objOLNamespace = objOLApp.GetNameSpace("MAPI")
    objOLNamespace.Logon
    Set objOLMail = objOLApp.CreateItem(0)
    With objOLMail
      .Recipients.Add ServiceEMail
      .Subject = "Bestellung"
      .Body = "Bitte Getränkeschacht für " & Name & _
         " schnellstmöglich auffüllen. Vielen Dank!"
      .Send
    End With
    Set objOLMail = Nothing
    Set objOLNameSpace = Nothing
    objOLApp.Quit
    Set objOLApp = Nothing
  End Function

End Class


Dim objDrink
Dim varCheckEntnahme

' HIER INSTANZIEREN WIR DAS OBJEKT OBJDRINK...
Set objDrink = New Drink

' ...UND SETZEN EIGENSCHAFTEN FÜR DAS OBJEKT
objDrink.ServiceEMail = "info@markmuendl.de"
objDrink.Name = "Mineralwasser"
objDrink.Preis = 0.3
objDrink.Koffeinfrei = True
objDrink.Kalorienfrei = True
objDrink.Menge = 2

' WIR SIMULIEREN EINE GETRÄNKEENTNAHME...
varCheckEntnahme = objDrink.entnehmeGetraenk()

' ...UND LASSEN UND DIE VERBLEIBENDE MENGE ANZEIGEN
Msgbox objDrink.Menge, vbInformation, "Verbleibende Menge"
