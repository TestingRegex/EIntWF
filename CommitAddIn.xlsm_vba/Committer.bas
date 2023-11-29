'''
'   Ein Excel Makro was an den Button im Add-in Tab gebunden ist und
'   die Aufgabe des Committen übernimmt
'
'   Allgemeines:
'       Das Programm gibt die gewünschten Git-Befehle an eine Shell-Instanz weiter damit diese ausgeführt werden.
'
'   Verwendete Funktionen:
'       Saver, Pathing, BadCharacterFilter, UserPromptYesNo, UserPromptText
'''

Option Explicit

Sub CommitToGit(control As Office.IRibbonControl)
    
    AnnoyUsers
    Commit (False)
    
End Sub

Function Commit(ByVal ForcedStandardCommit As Boolean)

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
            customCommitMessage = UserPromptText("Bitte gebe hier deine Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier angeben")
            
            ' Commit messages should not be empty
            If customCommitMessage = "" Then
                MsgBox "Es wurde keine Commit Nachricht eingegeben der Commit Vorgang wird abgebrochen."
                Exit Function
            End If
            
            Do While BadCharacterFilter(customCommitMessage, "Commit")
            
                customCommitMessage = UserPromptText("Die eingegebene Commit Nachricht war ungültig. Bitte geben Sie hier deine Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier angeben")
                If customCommitMessage = "" Then
                    MsgBox "Es wurde keine Commit Nachricht eingegeben der Commit Vorgang wird abgebrochen."
                    Exit Function
                End If
            Loop
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
    
    temp = ShellCommand(gitCommand, "Die Änderungen wurden commitet.", "Die Änderungen konnten nicht commitet werden. Versuchen Sie es bitte manuell über eine Shellinstanz.")
    
End Function