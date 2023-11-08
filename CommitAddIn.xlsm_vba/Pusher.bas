'''
'   Ein Excel Makro was an den Button im Add-in Tab gebunden ist und
'   die Aufgabe des Pushen �bernimmt
'
'
'
'''

Option Explicit

Sub PushToGit(ByRef control As Office.IRibbonControl)

    Push
    
End Sub

Function Push()

    Dim GitCommand As String
    Dim WorkbookPath As String

'------------------------------------------------------------------------
' Das richtige Directory finden

    ' Get the path of the current workbook
    WorkbookPath = ActiveWorkbook.path

    ' Moving into the git repo
    ChDir WorkbookPath
    
'-----------------------------------------------------------------------
' git push ausf�hren
    
    GitCommand = "git push"
    Shell GitCommand, vbNormalFocus
    
    MsgBox "Committed �nderungen wurden gepusht."

End Function