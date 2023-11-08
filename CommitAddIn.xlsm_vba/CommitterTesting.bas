Option Explicit

Sub CommitToGit_Test()

    Dim GitCommand As String
    Dim WorkbookPath As String
    Dim customCommit As Long
    Dim customCommitMessage As String
    Dim commitMessage As String
    
'-----------------------------------------------------------------------------------
' Git Repo wird ausgew�hlt
' Momentan wird angenommen dass das Workbook im gleichen Ort liegt wie das Repo
    
    
    ' Get the path of the current workbook
    WorkbookPath = ActiveWorkbook.path

    ' Zum Gew�nschten Ordner hinbewegen
    ChDir WorkbookPath
    
'-----------------------------------------------------------------------------------
' Die Dateien die gestaged werden werden ausgew�hlt
' Momentan werden alle �nderung gestaged
    
    
    ' All �nderungen im Git Repo werden aufeinmal hinzugef�gt
    GitCommand = "git add --all"
    'Shell GitCommand, vbNormalFocus
    MsgBox GitCommand
    
    ' Nochmal spezifisch den Exportierordner angeben
    ' Eigentlich nicht mehr notwendig!!
    GitCommand = "git add """ & WorkbookPath & "\" & ActiveWorkbook.Name & "_vba" & """"
    'Shell GitCommand, vbNormalFocus
    MsgBox GitCommand
    
    ' Spezifisch das Aktive Workbook stagen
    
    GitCommand = "git add """ & ActiveWorkbook.Name & """"
    ' Shell GitCommand, vbNormalFocus
    MsgBox GitCommand
    
    
'-------------------------------------------------------------------------------------
' Commit Prozess f�ngt an
    
    customCommit = UserPromptYesNo("M�chten Sie eine Commit Nachricht selber erstellen?")
    
    If customCommit = vbYes Then
        ' Custom Commit Nachricht wird erstellt
        customCommitMessage = UserInputText("Bitte gebe hier deine Commit Nachricht an.", "Custom Commit Nachricht erstellen.", "")
        
        commitMessage = customCommitMessage & " - " & GetUser()
    Else
        ' Standard Commit Nachricht wird erstellt
        commitMessage = "Commit erstellt von " & GetUser()
    End If
    
    GitCommand = "git commit -m """ & commitMessage & """"
    MsgBox GitCommand
    'Shell GitCommand, vbNormalFocus
    
    MsgBox "Die �nderungen wurden committet."
End Sub
