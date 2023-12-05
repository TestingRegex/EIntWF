Attribute VB_Name = "Puller"
'''
'   Ein Excel Makro was an einen Button im Add-in Tab gebunden ist und
'   die Aufgabe des Export, Commit, und Push aufeinmal übernimmt
'
'   Allgemeines:
'       Das Programm gibt die gewünschten Git-Befehle an eine Shell-Instanz weiter damit diese ausgeführt werden.
'
'   Verwendete Funktionen:
'       Pathing,
'''

Option Explicit

Sub GitPull(ByRef contral As Office.IRibbonControl)
On Error GoTo ErrHandler

    If AnnoyUsers = vbYes Then
        Pull
    End If
    
ExitSub:
    Exit Sub
    
ErrHandler:
    MsgBox "Im " & Err.Source & " Vorgang ist ein Fehler aufgetreten." & vbCrLf & Err.Description
    Resume ExitSub
    Resume

End Sub

Function Pull()

    Dim gitCommand As String
    Dim temp As Integer

'------------------------------------------------------------------------
' Get the desired path

    Pathing
    
'-----------------------------------------------------------------------
' execute commands
    
    gitCommand = "git pull"
    temp = ShellCommand(gitCommand, "Updates wurden von GitHub heruntergeladen.", "Es konnten keine Updates heruntergeladen werden.", "Pull")


End Function

