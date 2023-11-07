Sub CommitToGit(control As Office.IRibbonControl)

    Dim GitCommand As String
    Dim WorkbookPath As String
    
    ' Get the path of the current workbook
    WorkbookPath = ActiveWorkbook.path

    MsgBox WorkbookPath
    ' Moving into the git repo
    ChDir WorkbookPath

    ' Add the current workbook to the Git repository
    GitCommand = "git add --all"
    Shell GitCommand, vbNormalFocus
    GitCommand = "git add """ & WorkbookPath & "\" & ActiveWorkbook.Name & "_vba" & """"
    Shell GitCommand, vbNormalFocus
    
    Application.Wait Now + TimeSerial(0, 0, 0.1)
    
    ' Commit the changes
    GitCommand = "git commit -m ""Auto-Committing the current workbook for " & GetUser() & """"
    Shell GitCommand, vbNormalFocus
    
    MsgBox "The Commit macro has been run!"
End Sub

