Sub CommitToGit(control As Office.IRibbonControl)

    Dim GitCommand As String
    Dim WorkbookPath As String
    
    ' Get the path of the current workbook
    WorkbookPath = ActiveWorkbook.Path

    ' Moving into the git repo
    ChDir "C:\Users\d60157\Documents\Projects\Swissgrid\GitTests"

    ' Add the current workbook to the Git repository
    GitCommand = "git add """ & WorkbookPath & "\" & ThisWorkbook.Name & """"
    Shell GitCommand, vbNormalFocus
    GitCommand = "git add """ & WorkbookPath & "\" & ThisWorkbook.Name & "_vba" & """"
    Shell GitCommand, vbNormalFocus
    
    Application.Wait Now + TimeSerial(0, 0, 0.5)
    
    ' Commit the changes
    GitCommand = "git commit -m ""Auto-Committing the current workbook"""
    Shell GitCommand, vbNormalFocus
    
    MsgBox "The Commit macro has been run!"
End Sub

