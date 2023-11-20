Option Explicit

Sub CommitToGit_Test()

    Dim GitCommand As String
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
' Git Repo wird ausgewählt
' Momentan wird angenommen dass das Workbook im gleichen Ort liegt wie das Repo

    
    Pathing
    
'-----------------------------------------------------------------------------------
' Die Dateien die gestaged werden werden ausgewählt
' Momentan werden alle Änderung gestaged
    
    
    ' All Änderungen im Git Repo werden aufeinmal hinzugefügt
    GitCommand = "git add --all"
    shell GitCommand, vbNormalFocus
    
    ' Nochmal spezifisch den Exportierordner angeben
    ' Eigentlich nicht mehr notwendig!!
    GitCommand = "git add """ & WorkbookPath & "\" & ActiveWorkbook.Name & "_vba" & """"
    shell GitCommand, vbNormalFocus
    
    ' Spezifisch das Aktive Workbook stagen
    
    GitCommand = "git add """ & ActiveWorkbook.Name & """"
    shell GitCommand, vbNormalFocus
    
'-------------------------------------------------------------------------------------
' Commit Prozess fängt an
    If Not ForcedStandardCommit Then
        customCommit = UserPromptYesNo("Möchten Sie eine Commit Nachricht selber erstellen?")
        
        If customCommit = vbYes Then
            ' Custom Commit Nachricht wird erstellt
            customCommitMessage = UserInputText("Bitte gebe hier deine Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier angeben")
            
            ' Leere Commit Nachricht prüfen:
            If customCommitMessage = "" Then
                MsgBox "Es wurde keine Commit Nachricht eingegeben der Commit Vorgang wird abgebrochen."
                Exit Sub
            End If
            
            Do While BadCharacterFilter(customCommitMessage, "Commit")
            
                customCommitMessage = UserInputText("Die eingegebene Commit Nachricht war ungültig. Bitte geben Sie hier deine Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier angeben")
                If customCommitMessage = "" Then
                    MsgBox "Es wurde keine Commit Nachricht eingegeben der Commit Vorgang wird abgebrochen."
                    Exit Sub
                End If
            Loop
            
            
            commitMessage = customCommitMessage & " - " & GetUser()
        End If
    Else
        ' Standard Commit Nachricht wird erstellt
        commitMessage = "Commit erstellt von " & GetUser()
    End If
    
    GitCommand = "git commit -m """ & commitMessage & """"
    shell GitCommand, vbNormalFocus
    
    MsgBox "Die Änderungen wurden committet."
End Sub
