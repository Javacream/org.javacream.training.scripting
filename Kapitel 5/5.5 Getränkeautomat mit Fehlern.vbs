Option Explicit
Dim objConnection
Dim objRecordset
Dim objDrink
Dim varDBPfad
Dim varSQL
Dim varDrinks()
Dim i
Dim varAusgabe
Dim varApp
Dim varDrink

' ERSTELLEN WIR ZUNÄCHST UNSERE GETRÄNKEKLASSE
Class Drink
  Public ID
  Public Name
  Public Preis
  Public Menge
  Public Koffeinfrei
  Public Kalorienfrei
End Class

varApp = "Getränkeautomat"

' PFAD ZUR DATENBANK FESTLEGEN
varDBPfad = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - 3) & "mdb"

' DATENBANKZUGRIFF PER ADO (ACTIVEX DATA OBJECTS)
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "Microsoft.Jet.OLEDB.4.0"
objConnection.Open varDBPfad


' ZUGRIFF AUF ALLE DATEN IN DER TABELLE GETRAENKE
varSQL = "SELECT * FROM Getraenke"
Set objRecordset = objConnection.Execute(varSQL)

' PRÜFEN OB TABELLE DATENSÄTZE ENTHÄLT
If Not objRecordset.EOF Then

  ' DANN GEHEN WIR ZUM ERSTEN DATENSATZ IN DER TABELLE
  objRecordset.MoveFirst
  i = 0

  ' JETZT MACHEN WIR EINE SCHLEIFE DURCH ALLE DATENSÄTZE
  Do While not objRecordset.EOF

    ' HIER WIRD EIN NEUES GETRÄNKEOBJEKT INSTANZIERT...
    Set objDrink = New Drink

    ' UND MIT DATEN BEFÜLLT
    objDrink.ID = objRecordset.Fields("GID").Value
    objDrink.Name = objRecordset.Fields("GName").Value
    objDrink.Preis = objRecordset.Fields("GPreis").Value
    objDrink.Menge = objRecordset.Fields("GMenge").Value

    ' UNSER GETRÄNKEDATENFELD WIRD REDIMENSIONIERT
    ReDim Preserve varDrinks(i)
    Set varDrinks(i) = objDrink
    i = i + 1

    ' DANN GEHEN WIR ZUM NÄCHSTEN DATENSATZ
    objRecordset.MoveNext

  ' UND JETZT DAS GANZE VON VORNE
  Loop
End If

' DATENBANKVERBINDUNG SCHLIESSEN
objConnection.Close
' OBJEKTE AUS DEM ARBEITSSPEICHER ENTFERNEN
Set objRecordset = Nothing
Set objConnection = Nothing

' JETZT KÜMMERN WIR UNS UM DIE DATENAUSGABE
' DA WIR ALLE DATEN IN EINEM DATENFELD MIT GETRÄNKEOBJEKTEN
' HABEN, BRAUCHEN WIR NUR NOCH IN EINER SCHLEIFE DIE DATEN
' ABZURUFEN...
For Each objDrink in varDrinks
  varAusgabe = varAusgabe & _
    objDrink.Name & vbTab & _
    FormatCurrency(objDrink.Preis) & vbTab & _
    objDrink.Menge & "      " & vbCrLf
Next

' ...NOCH EINE ÜBERSCHRIFT DAVOR SETZTEN ...
varAusgabe = "Getränkevorrat:" & vbCrLf & vbCrLf & varAusgabe

' ...UND FERTIG IST DIE DATENAUSGABE
Msgbox varAusgabe, vbInformation, varApp

Msgbox "Leider ist die Getränkewahltaste defekt." & _
  vbCrLf & "Soll der Automat für Sie entscheiden?", vbQuestion, varApp

varDrink = "Orangensaft"

Msgbox "Hier kommt Ihr " & varDrink & "!", vbInformation, varApp







