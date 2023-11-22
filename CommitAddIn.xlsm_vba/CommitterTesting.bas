Option Explicit

Sub CommitToGit_Test()

    Dim gitCommand As String
    Dim WorkbookPath As String
    Dim customCommit As Long
    Dim customCommitMessage As String
    Dim commitMessage As String
    Dim ForcedStandardCommit As Boolean
    
    ForcedStandardCommit = False

'---------------------------------------------------------------------------------------------
' Einmal Alles Speichern.

    Saver

'-----------------------------------------------------------------------------------
' Git Repo wird ausgew�hlt
' Momentan wird angenommen dass das Workbook im gleichen Ort liegt wie das Repo

    
    Pathing
    
'-----------------------------------------------------------------------------------
' Die Dateien die gestaged werden werden ausgew�hlt
' Momentan werden alle �nderung gestaged
    
    
    ' All �nderungen im Git Repo werden aufeinmal hinzugef�gt
    gitCommand = "git add --all"
    shell gitCommand, vbNormalFocus
    
    ' Nochmal spezifisch den Exportierordner angeben
    ' Eigentlich nicht mehr notwendig!!
    gitCommand = "git add """ & WorkbookPath & "\" & ActiveWorkbook.Name & "_vba" & """"
    shell gitCommand, vbNormalFocus
    
    ' Spezifisch das Aktive Workbook stagen
    
    gitCommand = "git add """ & ActiveWorkbook.Name & """"
    shell gitCommand, vbNormalFocus
    
'-------------------------------------------------------------------------------------
' Commit Prozess f�ngt an
    If Not ForcedStandardCommit Then
        customCommit = UserPromptYesNo("M�chten Sie eine Commit Nachricht selber erstellen?")
        
        If customCommit = vbYes Then
            ' Custom Commit Nachricht wird erstellt
            customCommitMessage = UserInputText("Bitte gebe hier deine Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier angeben")
            
            ' Leere Commit Nachricht pr�fen:
            If customCommitMessage = "" Then
                MsgBox "Es wurde keine Commit Nachricht eingegeben der Commit Vorgang wird abgebrochen."
                Exit Sub
            End If
            
            Do While BadCharacterFilter(customCommitMessage, "Commit")
            
                customCommitMessage = UserInputText("Die eingegebene Commit Nachricht war ung�ltig. Bitte geben Sie hier deine Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier angeben")
                If customCommitMessage = "" Then
                    MsgBox "Es wurde keine Commit Nachricht eingegeben der Commit Vorgang wird abgebrochen."
                    Exit Sub
                End If
            Loop
            
            
            commitMessage = customCommitMessage & " - " & GetUser()
        Else
            ' Standard Commit Nachricht wird erstellt
            commitMessage = "Commit erstellt von " & GetUser()
        End If
    Else
        ' Standard Commit Nachricht wird erstellt
        commitMessage = "Commit erstellt von " & GetUser()
    End If
    
    gitCommand = "git commit -m """ & commitMessage & """"
    Debug.Print gitCommand
    
    'Dim temp As Integer
    
    'temp = ShellCommand(gitCommand, "Die �nderungen wurden commitet.", "Die �nderungen konnten nicht commitet werden. Versuchen Sie es bitte manuell �ber eine Shellinstanz.")
    
End Sub
