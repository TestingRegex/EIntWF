Option Explicit

Sub GitPush(ByRef control As Office.IRibbonControl)

    Dim GitCommand As String
    Dim WorkbookPath As String
    
    ' Get the path of the current workbook
    WorkbookPath = ActiveWorkbook.path

    MsgBox WorkbookPath
    ' Moving into the git repo
    ChDir WorkbookPath
    
    GitCommand = "git push"
    Shell GitCommand, vbNormalFocus
    
    MsgBox "Committed �nderungen wurden gepusht."
End Sub