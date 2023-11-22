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

    Dim gitCommand As String
    Dim temp As Integer

'------------------------------------------------------------------------
' Das richtige Directory finden

    Pathing
    
'-----------------------------------------------------------------------
' git push ausführen
    
    gitCommand = "git push"
    
    temp = ShellCommand(gitCommand, "Committed Änderungen wurden gepusht.", "Der Push-Vorgang ist gescheitert.")
    

End Function