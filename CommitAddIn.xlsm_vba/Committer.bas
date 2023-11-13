'''
'   Ein Excel Makro was an den Button im Add-in Tab gebunden ist und
'   die Aufgabe des Committen �bernimmt
'
'
'
'''

Option Explicit

Sub CommitToGit(control As Office.IRibbonControl)

    Commit
    
End Sub

Function Commit()

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
    Shell GitCommand, vbNormalFocus
    
    ' Nochmal spezifisch den Exportierordner angeben
    ' Eigentlich nicht mehr notwendig!!
    GitCommand = "git add """ & WorkbookPath & "\" & ActiveWorkbook.Name & "_vba" & """"
    Shell GitCommand, vbNormalFocus
    
    ' Spezifisch das Aktive Workbook stagen
    
    GitCommand = "git add """ & ActiveWorkbook.Name & """"
    Shell GitCommand, vbNormalFocus
    
'-------------------------------------------------------------------------------------
' Commit Prozess f�ngt an
    
    customCommit = UserPromptYesNo("M�chten Sie eine Commit Nachricht selber erstellen?")
    
    If customCommit = vbYes Then
        ' Custom Commit Nachricht wird erstellt
        customCommitMessage = UserInputText("Bitte gebe hier deine Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier angeben")
        
        commitMessage = customCommitMessage & " - " & GetUser()
    Else
        ' Standard Commit Nachricht wird erstellt
        commitMessage = "Commit erstellt von " & GetUser()
    End If
    
    GitCommand = "git commit -m """ & commitMessage & """"
    Shell GitCommand, vbNormalFocus
    
    MsgBox "Die �nderungen wurden committet."

End Function