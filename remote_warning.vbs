Set shell = CreateObject("WScript.Shell")
Dim pfad, dateiname

' Name deiner .exe Datei hier anpassen:
dateiname = "remote.exe"

' Ermittelt den aktuellen Ordner des Scripts
pfad = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' Erstellt den Dialog (4 = Ja/Nein Buttons, 32 = Fragezeichen-Icon)
antwort = MsgBox("Moechtest du das wirklich die Remote Hilfe starten? Die Remote Hilfe ist unsicher! Benutze die Remote Hilfe nur wenn nichts anderes funktioniert und du Maikel oder eine Person die die Remote Hilfe benutzen kann in deiner Klasse hast!", 4 + 32, "Warnung")

' Wenn JA (6) geklickt wurde
If antwort = 6 Then
    shell.Run Chr(34) & pfad & "\" & dateiname & Chr(34)
End If

' Das Script beendet sich danach automatisch