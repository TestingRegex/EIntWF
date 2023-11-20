'''
'   Ein Excel Makro was an den Button im Add-in Tab gebunden ist und
'   die Aufgabe des Pushen übernimmt
'
'   Allgemeines:
'       Das Programm gibt die gewünschten Git-Befehle an eine Shell-Instanz weiter damit diese ausgeführt werden.
'
'   Verwendete Funktionen:
'       Pathing,
'''

Option Explicit

Sub PushToGit(ByRef control As Office.IRibbonControl)

    Push
    
End Sub

Function Push()

    Dim GitCommand As String
    Dim temp As Integer

'------------------------------------------------------------------------
' Das richtige Directory finden

    Pathing
    
'-----------------------------------------------------------------------
' git push ausführen
    
    GitCommand = "git push"
    
    temp = ShellCommand(GitCommand, "Committed Änderungen wurden gepusht.", "Der Push-Vorgang ist gescheitert.")
    

End Function