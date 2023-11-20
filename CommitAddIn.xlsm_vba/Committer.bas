'''
'   Ein Excel Makro was an den Button im Add-in Tab gebunden ist und
'   die Aufgabe des Committen �bernimmt
'
'   Allgemeines:
'       Das Programm gibt die gew�nschten Git-Befehle an eine Shell-Instanz weiter damit diese ausgef�hrt werden.
'
'   Verwendete Funktionen:
'       Saver, Pathing, BadCharacterFilter, UserPromptYesNo, UserInputText
'''

Option Explicit

Sub CommitToGit(control As Office.IRibbonControl)

    Commit (False)
    
End Sub

Function Commit(ByVal ForcedStandardCommit As Boolean)

    Dim GitCommand As String
    Dim WorkbookPath As String
    Dim customCommit As Long
    Dim customCommitMessage As String
    Dim commitMessage As String

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
' Commit Prozess f�ngt an
    If Not ForcedStandardCommit Then
        customCommit = UserPromptYesNo("M�chten Sie eine Commit Nachricht selber erstellen?")
        
        If customCommit = vbYes Then
            ' Custom Commit Nachricht wird erstellt
            customCommitMessage = UserInputText("Bitte gebe hier deine Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier angeben")
            
            ' Leere Commit Nachricht pr�fen:
            If customCommitMessage = "" Then
                MsgBox "Es wurde keine Commit Nachricht eingegeben der Commit Vorgang wird abgebrochen."
                Exit Function
            End If
            
            Do While BadCharacterFilter(customCommitMessage, "Commit")
            
                customCommitMessage = UserInputText("Die eingegebene Commit Nachricht war ung�ltig. Bitte geben Sie hier deine Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier angeben")
                If customCommitMessage = "" Then
                    MsgBox "Es wurde keine Commit Nachricht eingegeben der Commit Vorgang wird abgebrochen."
                    Exit Function
                End If
            Loop
            
            
            commitMessage = customCommitMessage & " - " & GetUser()
        End If
    Else
        ' Standard Commit Nachricht wird erstellt
        commitMessage = "Commit erstellt von " & GetUser()
    End If
    
    GitCommand = "git commit -m """ & commitMessage & """"""
    Debug.Print "GitCommand:"; GitCommand
    
    Dim temp As Integer
    
    temp = ShellCommand(GitCommand, "Die �nderungen wurden commitet.", "Die �nderungen konnten nicht commitet werden. Versuchen Sie es bitte manuell �ber eine Shellinstanz.")
    
End Function