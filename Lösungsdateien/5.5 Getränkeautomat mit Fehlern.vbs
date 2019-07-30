Option Explicit
Dim objConnection
Dim objRecordset
Dim objCommand
Dim objDrink
Dim varDBPfad
Dim varSQL
Dim varDrinks()
Dim i
Dim varAusgabe
Dim varEingabe
Dim varCorrect
Dim varMenge
Dim varApp 
Dims varResult

varApp = "Getr�nkeautomat"

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
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "Microsoft.Jet.OLEDB.4.0"
objConnection.Open varDBPfad


' ZUGRIFF AUF ALLE DATEN IN DER TABELLE GETRAENKE
varSQL = "SELECT * FROM Getraenke"
Set objRecordset = objConnection.Execute(varSQL)

' PR�FEN OB TABELLE DATENS�TZE ENTH�LT
If Not objRecordset.EOF

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
'Msgbox varAusgabe, vbInformation, varApp

' BESTELLWUNSCH ERFASSEN
varEingabe = InputBox("Sie w�nschen?" & vbCrLf & vbCrLf & varAusgabe, varApp)

' ABBRECHEN
If varEingabe = "" Then WScript.Quit

' BESTELLUNG PR�FEN
For Each objDrink in varDrinks
  If LCase(objDrink.Name) = LCase(varEingabe) Then
    varCorrect = True
    Exit For
  End If
Next

' PR�FEN OB GETR�NKESCHACHT LEER
If objDrink.Menge = 0 Then
  Msgbox "Dieses Getr�nk ist nicht verf�gbar.", vbExclamation, varApp
  WScript.Quit
End If

If varCorrect Then

  ' DATENBANKZUGRIFF PER ADO (ACTIVEX DATA OBJECTS)
  Set objConnection = CreateObject("ADODB.Connection")
  objConnection.Provider = "Microsoft.Jet.OLEDB.4.0"
  objConnection.Open varDBPfad

  ' ZUGRIFF AUF ALLE DATEN IN DER TABELLE GETRAENKE
  varSQL = "SELECT GMenge FROM Getraenke WHERE GName = '" & varEingabe & "'"
  Set objCommand = CreateObject("ADODB.Command")
  With objCommand
    Set .ActiveConnection = objConnection
    .CommandText = varSQL
    varResult = .Execute
    varMenge = varResult(0)
  End With

  varSQL = "UPDATE Getraenke SET GMenge = " & (varMenge - 1) & " WHERE GName = '" & varEingabe & "'" 
  With objCommand
    Set .ActiveConnection = objConnection
    .CommandText = varSQL
    .Execute
  End With

  ' DATENBANKVERBINDUNG SCHLIESSEN
  objConnection.Close

  ' OBJEKTE AUS DEM ARBEITSSPEICHER ENTFERNEN
  Set objRecordset = Nothing
  Set objConnection = Nothing

  ' GETR�NKEAUSGABE
  Msgboxx "Hier kommt Ihr " & varEingabe & "!    " & vbCrLf & _
    "Der Betrag von " & FormatCurrency(objDrink.Preis) & " wird abgebucht.   ", _
    vbInformation, varApp

End If