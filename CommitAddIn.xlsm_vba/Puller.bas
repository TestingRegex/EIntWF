'''
'   Ein Excel Makro was an einen Button im Add-in Tab gebunden ist und
'   die Aufgabe des Export, Commit, und Push aufeinmal �bernimmt
'
'   Allgemeines:
'       Das Programm gibt die gew�nschten Git-Befehle an eine Shell-Instanz weiter damit diese ausgef�hrt werden.
'
'   Verwendete Funktionen:
'       Pathing,
'''

Option Explicit

Sub GitPull(ByRef contral As Office.IRibbonControl)

    Pull

End Sub

Function Pull()

    Dim GitCommand As String

'------------------------------------------------------------------------
' Das richtige Directory finden

    Pathing
    
'-----------------------------------------------------------------------
' git push ausf�hren
    
    GitCommand = "git pull"
    shell GitCommand, vbNormalFocus
    
    MsgBox "Updates wurden von GitHub gepullt."



End Function
