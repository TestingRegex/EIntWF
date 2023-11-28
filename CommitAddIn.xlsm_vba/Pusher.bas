Attribute VB_Name = "Pusher"
'''
'   Ein Excel Makro was an den Button im Add-in Tab gebunden ist und
'   die Aufgabe des Pushen �bernimmt
'
'   Allgemeines:
'       Das Programm gibt die gew�nschten Git-Befehle an eine Shell-Instanz weiter damit diese ausgef�hrt werden.
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
' get desired path

    Pathing
    
'-----------------------------------------------------------------------
' execute commands
    
    gitCommand = "git push"
    
    temp = ShellCommand(gitCommand, "Committed �nderungen wurden gepusht.", "Der Push-Vorgang ist gescheitert.")
    

End Function
