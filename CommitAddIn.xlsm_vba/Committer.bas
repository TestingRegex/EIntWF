Attribute VB_Name = "Committer"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   This module contains the macros and major functions used in the '�nderungen Commiten'
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

    If AnnoyUsers = vbYes Then
        Commit False, False
    End If
    
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

    If SelectIndividualFiles Then
        MsgBox "You want to import individual files."
        Exit Sub
    Else
        ' All �nderungen im Git Repo werden aufeinmal hinzugef�gt
        gitCommand = "git add -u"
        shell gitCommand, vbNormalFocus
        
        ' Nochmal spezifisch den Exportierordner angeben
        gitCommand = "git add " & activeWorkbook.Name & "_vba" & "/* " & activeWorkbook.Name
        shell gitCommand, vbNormalFocus
    End If
        
'-------------------------------------------------------------------------------------
' Commit Message Dialoge:

    If Not ForcedStandardCommit Then
        customCommit = UserPromptYesNo("M�chten Sie eine Commit Nachricht selber erstellen?")
        
        If customCommit = vbYes Then
            ' Get user input for commit message.
            customCommitMessage = UserPromptText("Bitte gebe hier deine Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier angeben", "Commit")
            
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

    ShellCommand gitCommand, "Die �nderungen wurden commitet.", "Die �nderungen konnten nicht commitet werden. Versuchen Sie es bitte manuell �ber eine Shellinstanz.", "Commit"
    
End Sub
