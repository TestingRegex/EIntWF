Attribute VB_Name = "Pusher"
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
On Error GoTo ErrHandler

    If AnnoyUsers = vbYes Then
        Push
    End If
    
ExitSub:
    Exit Sub
    
ErrHandler:
    MsgBox "Im " & Err.Source & " Vorgang ist ein Fehler aufgetreten." & vbCrLf & Err.Description
    Resume ExitSub
    Resume
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
    
    temp = ShellCommand(gitCommand, "Die gecommiteten Änderungen wurden hochgeladen.", "Der Push Vorgang ist gescheitert.", "Push")
    

End Function
