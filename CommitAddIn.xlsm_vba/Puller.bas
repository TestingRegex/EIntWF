'''
'   Ein Excel Makro was an einen Button im Add-in Tab gebunden ist und
'   die Aufgabe des Export, Commit, und Push aufeinmal übernimmt
'
'   Allgemeines:
'       Das Programm gibt die gewünschten Git-Befehle an eine Shell-Instanz weiter damit diese ausgeführt werden.
'
'''

Option Explicit

Sub GitPull(ByRef contral As Office.IRibbonControl)

    Pull

End Sub

Function Pull()

    Dim GitCommand As String
    Dim WorkbookPath As String

'------------------------------------------------------------------------
' Das richtige Directory finden

    ' Get the path of the current workbook
    WorkbookPath = ActiveWorkbook.path

    ' Moving into the git repo
    ChDir WorkbookPath
    
'-----------------------------------------------------------------------
' git push ausführen
    
    GitCommand = "git pull"
    Shell GitCommand, vbNormalFocus
    
    MsgBox "Updates wurden von GitHub gepullt."



End Function
