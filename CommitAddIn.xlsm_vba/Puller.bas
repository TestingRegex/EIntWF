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

    Pull

End Sub

Function Pull()

    Dim gitCommand As String
    Dim temp As Integer

'------------------------------------------------------------------------
' Das richtige Directory finden

    Pathing
    
'-----------------------------------------------------------------------
' git push ausführen
    
    gitCommand = "git pull"
    temp = ShellCommand(gitCommand, "Updates wurden von GitHub gepulled.", "Es konnten keine Updates gepulled werden.")


End Function
