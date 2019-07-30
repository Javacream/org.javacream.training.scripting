Option Explicit
Dim objConnection
Dim objRecordset
Dim objDrink
Dim varDBPfad
Dim varSQL
Dim varDrinks()
Dim i
Dim varAusgabe

' ERSTELLEN WIR ZUN�CHST UNSERE GETR�NKEKLASSE
Class Drink
  Public ID
  Public Name
  Public Preis
  Public Menge
  Public Koffeinfrei
  Public Kalorienfrei
End Class

' PFAD ZUR DATENBANK FESTLEGEN
varDBPfad = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - 3) & "mdb"

' DATENBANKZUGRIFF PER ADO (ACTIVEX DATA OBJECTS)
set objConnection=CreateObject("ADODB.Connection")
objConnection.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0"
objConnection.open varDBPfad


' ZUGRIFF AUF ALLE DATEN IN DER TABELLE GETRAENKE
varSQL = "SELECT * FROM Getraenke"
Set objRecordset = objConnection.Execute(varSQL)

' PR�FEN OB TABELLE DATENS�TZE ENTH�LT
If Not objRecordset.EOF Then

  ' DANN GEHEN WIR ZUM ERSTEN DATENSATZ IN DER TABELLE
  objRecordset.MoveFirst
  i = 0

  ' JETZT MACHEN WIR EINE SCHLEIFE DURCH ALLE DATENS�TZE
  Do While not objRecordset.EOF

    ' HIER WIRD EIN NEUES GETR�NKEOBJEKT INSTANZIERT...
    Set objDrink = New Drink

    ' UND MIT DATEN BEF�LLT
    objDrink.ID = objRecordset.Fields("GID").Value
    objDrink.Name = objRecordset.Fields("GName").Value
    objDrink.Preis = objRecordset.Fields("GPreis").Value
    objDrink.Menge = objRecordset.Fields("GMenge").Value

    ' UNSER GETR�NKEDATENFELD WIRD REDIMENSIONIERT
    ReDim Preserve varDrinks(i)
    Set varDrinks(i) = objDrink
    i = i + 1

    ' DANN GEHEN WIR ZUM N�CHSTEN DATENSATZ
    objRecordset.MoveNext

  ' UND JETZT DAS GANZE VON VORNE
  Loop
End If

' DATENBANKVERBINDUNG SCHLIESSEN
objConnection.Close
' OBJEKTE AUS DEM ARBEITSSPEICHER ENTFERNEN
Set objRecordset = Nothing
Set objConnection = Nothing

' JETZT K�MMERN WIR UNS UM DIE DATENAUSGABE
' DA WIR ALLE DATEN IN EINEM DATENFELD MIT GETR�NKEOBJEKTEN
' HABEN, BRAUCHEN WIR NUR NOCH IN EINER SCHLEIFE DIE DATEN
' ABZURUFEN...
For Each objDrink in varDrinks
  varAusgabe = varAusgabe & _
    objDrink.Name & vbTab & _
    FormatCurrency(objDrink.Preis) & vbTab & _
    objDrink.Menge & "    " & vbCrLf
Next

' ...NOCH EINE �BERSCHRIFT DAVOR SETZTEN ...
varAusgabe = "Getr�nkevorrat:" & vbCrLf & vbCrLf & varAusgabe

' ...UND FERTIG IST DIE DATENAUSGABE
Msgbox varAusgabe, vbInformation, "Getr�nkeautomat"
