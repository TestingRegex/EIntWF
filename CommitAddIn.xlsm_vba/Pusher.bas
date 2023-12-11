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

Private Sub PushToGit(ByVal control As Office.IRibbonControl)
On Error GoTo ErrHandler

        Push
    
ExitSub:
    Exit Sub
    
ErrHandler:

    ErrorHandler Err.Number, Err.Source, Err.Description
    Resume ExitSub
    Resume

End Sub

Public Sub Push()

    Dim gitCommand As String

'------------------------------------------------------------------------
' get desired path

    Pathing
    
'-----------------------------------------------------------------------
' execute commands
    
    gitCommand = "git push"
    ShellCommand gitCommand, "Die gecommiteten �nderungen wurden hochgeladen.", "Der Push Vorgang ist gescheitert.", "Push"
    

End Sub
