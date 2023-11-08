Option Explicit

Sub Push_Test()

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
    
    GitCommand = "git push"
    Shell GitCommand, vbNormalFocus
    
    MsgBox "Committed Änderungen wurden gepusht."

End Sub