'''
'   Ein Excel Makro was an den Button im Add-in Tab gebunden ist und
'   die Aufgabe des Committen übernimmt
'
'   Allgemeines:
'       Das Programm gibt die gewünschten Git-Befehle an eine Shell-Instanz weiter damit diese ausgeführt werden.
'
'   Verwendete Funktionen:
'       Saver, Pathing, BadCharacterFilter, UserPromptYesNo, UserInputText
'''

Option Explicit

Sub CommitToGit(control As Office.IRibbonControl)

    Commit (False)
    
End Sub

Function Commit(ByVal ForcedStandardCommit As Boolean)

    Dim gitCommand As String
    Dim WorkbookPath As String
    Dim customCommit As Long
    Dim customCommitMessage As String
    Dim commitMessage As String

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
' Commit Prozess fängt an
    If Not ForcedStandardCommit Then
        customCommit = UserPromptYesNo("Möchten Sie eine Commit Nachricht selber erstellen?")
        
        If customCommit = vbYes Then
            ' Custom Commit Nachricht wird erstellt
            customCommitMessage = UserInputText("Bitte gebe hier deine Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier angeben")
            
            ' Leere Commit Nachricht prüfen:
            If customCommitMessage = "" Then
                MsgBox "Es wurde keine Commit Nachricht eingegeben der Commit Vorgang wird abgebrochen."
                Exit Function
            End If
            
            Do While BadCharacterFilter(customCommitMessage, "Commit")
            
                customCommitMessage = UserInputText("Die eingegebene Commit Nachricht war ungültig. Bitte geben Sie hier deine Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier angeben")
                If customCommitMessage = "" Then
                    MsgBox "Es wurde keine Commit Nachricht eingegeben der Commit Vorgang wird abgebrochen."
                    Exit Function
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
    Debug.Print "GitCommand:"; gitCommand
    
    Dim temp As Integer
    
    temp = ShellCommand(gitCommand, "Die Änderungen wurden commitet.", "Die Änderungen konnten nicht commitet werden. Versuchen Sie es bitte manuell über eine Shellinstanz.")
    
End Function