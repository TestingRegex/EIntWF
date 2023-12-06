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

Private Sub CommitToGit(control As Office.IRibbonControl)
    
On Error GoTo ErrHandler:

    If AnnoyUsers = vbYes Then
        Commit (False)
    End If
    
ExitSub:
    Exit Sub
    
ErrHandler:
    MsgBox "Im " & Err.Source & " Vorgang ist ein Fehler aufgetreten." & vbCrLf & Err.Description
    Resume ExitSub
    Resume
    
End Sub
' The function
Public Function Commit(ByVal ForcedStandardCommit As Boolean) As Variant

    Dim gitCommand As String
    Dim WorkbookPath As String
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
    
    
    ' All Änderungen im Git Repo werden aufeinmal hinzugefügt
    gitCommand = "git add -u"
    shell gitCommand, vbNormalFocus
    
    ' Nochmal spezifisch den Exportierordner angeben
    gitCommand = "git add " & ActiveWorkbook.Name & "_vba" & "/* " & ActiveWorkbook.Name
    shell gitCommand, vbNormalFocus
    
        
'-------------------------------------------------------------------------------------
' Commit Message Dialoge:

    If Not ForcedStandardCommit Then
        customCommit = UserPromptYesNo("Möchten Sie eine Commit Nachricht selber erstellen?")
        
        If customCommit = vbYes Then
            ' Get user input for commit message.
            customCommitMessage = UserPromptText("Bitte gebe hier deine Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier angeben", "Commit")
            
            ' Commit messages should not be empty
            If customCommitMessage = vbNullString Then
                MsgBox "Es wurde keine Commit Nachricht eingegeben der Commit Vorgang wird abgebrochen."
                Exit Function
            End If
            
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

    Dim temp As Integer
    
    temp = ShellCommand(gitCommand, "Die Änderungen wurden commitet.", "Die Änderungen konnten nicht commitet werden. Versuchen Sie es bitte manuell über eine Shellinstanz.", "Commit")
    
End Function
