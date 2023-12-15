Attribute VB_Name = "Committer"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   This module contains the macros and major functions used in the 'Änderungen Commiten'
'   button.
'
'   Purpose:
'       All changes to the tracked files in the repository are staged, as well as explicitly
'       staging any changes to the active workbook or the workbook _vba directory.
'       Then the user is prompted to either create a custom commit message or use the
'       standard commit message.
'
'   Used homemade functions:
'       AnnoyUsers, Saver, Pathing, BadCharacterFilter, UserPromptYesNo, UserPromptText,
'       ShellCommand
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Sub CommitToGit(ByVal control As Office.IRibbonControl)
    
On Error GoTo ErrHandler:

    
    Commit False, False
    
ExitSub:
    Exit Sub
    
ErrHandler:
    ErrorHandler Err.Number, Err.Source, Err.Description
    Resume ExitSub
    Resume
    
End Sub

Public Sub Commit(ByVal ForcedStandardCommit As Boolean, Optional ByVal SelectIndividualFiles As Boolean = False)

    Dim gitCommand As String
    Dim customCommit As Long
    Dim customCommitMessage As String
    Dim commitMessage As String

'---------------------------------------------------------------------------------------------
' Save everything in the workbook before committing it to git.

    Saver

'-----------------------------------------------------------------------------------
' Move to the desired git repo.
'
    Pathing
    
'-----------------------------------------------------------------------------------
' Staging files to be committed
' Currently we add all already tracked files and the new workbook and directory
' New: Try to add individually selected files.
    
    gitCommand = "git add "
    If SelectIndividualFiles Then
        Dim selectedFiles As Variant
        Dim i As Long
        
        selectedFiles = SelectFiles("Commit")
        
        If IsArray(selectedFiles) Then
            For i = LBound(selectedFiles) To UBound(selectedFiles)
                gitCommand = gitCommand & selectedFiles(i) & " "
            Next i
        Else
            gitCommand = gitCommand & selectedFiles
        End If
        
        'ShellCommand gitCommand, "Die gewählten Dateien wurden committet.", "Die gewählten Dateien konnten leider nicht committet werden", "Commit"
        Debug.Print gitCommand
        Exit Sub
    Else
        ' All Änderungen im Git Repo werden aufeinmal hinzugefügt
        gitCommand = gitCommand & " -u"
        shell gitCommand, vbNormalFocus
        
        ' Nochmal spezifisch den Exportierordner angeben
        gitCommand = gitCommand & activeWorkbook.Name & "_vba" & "/* " & activeWorkbook.Name
        shell gitCommand, vbNormalFocus
    End If
        
'-------------------------------------------------------------------------------------
' Commit Message Dialoge:

    If Not ForcedStandardCommit Then
        customCommit = UserPromptYesNo("Möchten Sie eine Commit Nachricht selber erstellen?")
        
        If customCommit = vbYes Then
            ' Get user input for commit message.
            customCommitMessage = UserPromptText("Geben Sie hier Ihre Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier eingeben", "Commit")
            
            commitMessage = customCommitMessage & " - " & GetUser()
        Else
            ' Standardized commit message
            commitMessage = "Commit erstellt von " & GetUser()
        End If
    Else
        ' Standardized Commit message
        commitMessage = "Commit erstellt von " & GetUser()
    End If
    
    gitCommand = "git commit -m """ & commitMessage & """"
    'Debug.Print "GitCommand:"; gitCommand
    
'-------------------------------------------------------------------------------------------
' Executing commit command.

    ShellCommand gitCommand, "Die Änderungen wurden commitet.", "Die Änderungen konnten nicht commitet werden. Versuchen Sie es bitte manuell über eine Shellinstanz.", "Commit"
    
End Sub
