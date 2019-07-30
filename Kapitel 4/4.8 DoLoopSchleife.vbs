Option Explicit
Dim varPasswort
Dim varPWEingabe

varPasswort = "Geheim"
varPWEingabe = ""
Do While varPWEingabe <> varPasswort
  varPWEingabe = Inputbox("Passwort eingeben","Passwort")
  If varPWEingabe <> varPasswort Then
    Msgbox "Falsches Passwort!", vbExclamation, "Passwort"
    If varPWEingabe = "" Then WScript.Quit
  End If
Loop
Msgbox "Passwort wurde akzeptiert!", vbInformation, "Passwort"