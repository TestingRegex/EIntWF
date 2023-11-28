'''
'
' A module to contain temporary tests that are used to test functions of my own or ones that are new to me.
' Dev Versions of the main functions are saved here.
'
'
' We note that the language in this module is not consistent.
'''

Option Explicit




Sub Testing()

    
    
End Sub




Sub Push_Test()

    Dim gitCommand As String
    Dim temp As Integer

'------------------------------------------------------------------------
' Das richtige Directory finden

    Pathing
    
'-----------------------------------------------------------------------
' git push ausführen
    
    gitCommand = "git push"
    temp = ShellCommand(gitCommand, "Committed Änderungen wurden gepusht.", "Der Push-Vorgang ist gescheitert.")

End Sub

Sub ImportMacros_Test()

    Dim selectedFolder As String ' Der Pfad zum Importordner
    
    Dim fs As Object 'FileSystemObject um mit System außerhalb von Excel interagieren zu können
    Dim folder As Object 'FileSystemObject: Der Ordner aus dem imortiert wird
    Dim file As Object 'FileSystemObject: Der Iterator beim Importieren
    Dim wb As Workbook ' Das Aktive Workbook
    Dim vbComp As Object ' Eine VBA Componente des Aktiven Workbooks
    Dim moduleName As String ' Der Name der importierten Module
    Dim newModuleName As String ' Der neue Modul Name des zu importierenden Moduls
    Dim benutzerMeinung As Long ' Entscheidung ob bereitsvorhandene Module überschrieben werden sollen oder nicht

    ' Set a reference to the Microsoft Scripting Runtime library.
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    
'---------------------------------------------------------------------------------------------------------
' Der Ordner aus dem importiert werden soll, wird ausgewählt.

    'Dialog damit User weiß was gleich zu tun ist.
    MsgBox "Bitte wählen Sie den Ordner aus, aus dem Sie die Makros importieren möchten."
    
    
    ' Der Pfad zum gewünschten Importordner wird erhoben.
    selectedFolder = SelectFolder()

    ' Falls kein Ordner ausgesucht wird, brechen wir ab.
    If selectedFolder = "" Then
        MsgBox "Kein Ordner ausgewählt. Import abgebrochen."
        Exit Sub
    End If
    
    ' Prüfen ob der ausgewählte Ordner existiert
    If Not fs.FolderExists(selectedFolder) Then
        MsgBox "Der gewünschte Ordner konnte nicht gefunden werden."
        Exit Sub
    End If
    
    
    ' Alle .bas Dateien werden aus dem Ordner importiert
    Set folder = fs.GetFolder(selectedFolder)
    Set wb = ActiveWorkbook

'---------------------------------------------------------------------------------------------------------
' .bas Dateien werden aus dem gewünschten Ordner importiert.

    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".bas" Then
             moduleName = Left(file.Name, Len(file.Name) - 4) ' Remove the last 4 characters (".bas").
            '-----------------------------------------------------------------------------------
            ' Namengebung des importierten Moduls:
            If ModulNamenSuchen(moduleName) Then
            
                benutzerMeinung = UserPromptYesNo(" Es gibt bereits ein Modul mit dem Namen '" + moduleName + "'. Soll das bereitsexistierende Modul überschrieben werden?")
                
                If benutzerMeinung = vbYes Then
                    
                    ' Remove the old Modul
                    RemoveModule (moduleName)
                    ' Import the .bas file into the workbook's VBA project.
                    Set vbComp = wb.VBProject.VBComponents.Import(file.path)
                    vbComp.Name = moduleName
                
                Else
                    
                    benutzerMeinung = UserPromptYesNo(" Möchten Sie das Modul '" + moduleName + "' unter einem anderen Namen speichern? (Bei 'Nein' wird das Modul übersprungen.)")
                    
                    If benutzerMeinung = vbYes Then
                        
                        newModuleName = UserPromptText("Wie soll das Modul heißen?", "", "")
                        
                        Do While ModulNamenSuchen(newModuleName)
                            '---------------------------------------------------------------------------------------------------------
                            ' Soll die Datei doch nicht importiert werden?
                            benutzerMeinung = UserPromptYesNo("Dieser Name ist bereits vergeben. Soll dieses Modul doch Übersprungen werden?")
                            
                            If benutzerMeinung = vbYes Then
                                
                                Dim skip As Boolean
                                skip = True
                                Exit Do
                            
                            End If
                            
                            If Not skip Then
                            newModuleName = UserPromptText("Wählen Sie bitte einen neuen Namen für das importierte Modul aus.", "", "Neuer Modulname")
                            End If
                        
                        Loop
                        
                        If Not skip Then
                            ' Datei wird unter dem neuen Namen importiert.
                            Set vbComp = wb.VBProject.VBComponents.Import(file.path)
                            vbComp.Name = newModuleName
                        End If
                    Else
                        
                        MsgBox "Das Modul '" + moduleName + "' wird nicht neu importiert."
                        
                    End If
                End If
            
            ' Es gibt kein Modul mit dem gleichen Namen wie die Datei:
            Else
                Set vbComp = wb.VBProject.VBComponents.Import(file.path)
                vbComp.Name = moduleName
            End If
        End If
    Next file
    
'---------------------------------------------------------------------------------------------------------
    ' Clean up.
    Set fs = Nothing
    Set folder = Nothing
    Set file = Nothing
    Set wb = Nothing
    Set vbComp = Nothing

    MsgBox "Alle gewünschten .bas Dateien aus " & selectedFolder & " wurden importiert."
    
End Sub




Sub Tag_Test()

' Benötigten Variablen init:

    Dim gitCommand As String
    Dim VersionInput As String
    Dim TagMessage As String
    Dim StringCheck As Boolean
    
'------------------------------------------------------
' Git-Pfad finden

    Pathing
    
'------------------------------------------------------
' Basic Ablauf:

    VersionInput = UserPromptText("Welche Version des Workbooks möchten Sie taggen?", "Versionsname", "_._")
    StringCheck = BadCharacterFilter(VersionInput, "Tag")
    If VersionInput = "" Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Sub
    End If
    Do While StringCheck
        VersionInput = UserPromptText("Der Eingebene Versionsname ist ungültig. Bitte geben Sie einen anderen Namen ein und vermeiden Sie die Zeichen: ' ~!@#$%^&*()+,{}[]|\;:'""<>/?='", "Versionsname", "_._")
        StringCheck = BadCharacterFilter(VersionInput, "Tag")
        If VersionInput = "" Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Sub
    End If
    Loop
    
    
    TagMessage = UserPromptText("Bitte geben Sie eine Kurze Beschreibung der Version oder ihrer Relevanz an:", "Versionsbeschreibung", "")
    StringCheck = BadCharacterFilter(TagMessage)
    If TagMessage = "" Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Sub
    End If
    Do While StringCheck
        TagMessage = UserPromptText("Bitte geben Sie eine Kurze Beschreibung der Version oder ihrer Relevanz an:", "Versionsbeschreibung", "")
        StringCheck = BadCharacterFilter(TagMessage)
        If TagMessage = "" Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Sub
        End If
    Loop
    
    gitCommand = "git tag -a " & VersionInput & " -m  """ & TagMessage & " - " & GetUser() & """"
    
    'MsgBox GitCommand
    shell gitCommand, vbNormalFocus


End Sub





Sub Export_Test()

    Dim wb As Workbook
    Dim WorkbookName As String
    Dim vbComp As Object
    Dim vbProj As Object
    Dim moduleName As String
    Dim moduleCode As String
    Dim outPath As String
    Dim modulePath As String
    Dim fileSysObj As Object
    Dim fs As Object

'---------------------------------------------------------------------------------------------
' Der Pfad zum Exportordner wird gefunden

    Set wb = ActiveWorkbook
    WorkbookName = wb.Name


    outPath = wb.path
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim vbaDirectory As String
    vbaDirectory = outPath & "\" & WorkbookName & "_vba\"
    
'---------------------------------------------------------------------------------------------
' Exportordner wird erstellt falls noch nicht vorhanden

    If Not fs.FolderExists(vbaDirectory) Then
        fs.CreateFolder vbaDirectory
    End If

'---------------------------------------------------------------------------------------------
' Die Module des VBA Projekts werden an den gewünschten Ort Exportiert


    ' Es wird durch alle Komponenten des VBA Projekts durch iteriert und alle Module werden exportiert.
    For Each vbProj In wb.VBProject.VBComponents
        If vbProj.Type = 1 Then ' Module
            moduleName = vbProj.Name
            
            ' Prüfen ob das Modul nicht einfach leer ist.
            If vbProj.CodeModule.CountOfLines > 0 Then
            
                moduleCode = vbProj.CodeModule.Lines(1, vbProj.CodeModule.CountOfLines)
            
                ' Inhalt des Moduls wird als String Variable geladen
                modulePath = vbaDirectory & moduleName & ".bas"
                
                ' Prüfen ob das Modul oder ein Modul mit diesem Namen bereits im Exportordner existiert
                If fs.FileExists(modulePath) Then
                
                    ' Inhalt der gleichnamigen Datei einladen
                    Dim textStream As Object
                    Set textStream = fs.OpenTextFile(modulePath, 1) ' 1: ForReading

                    Dim fileContent As String
                    fileContent = textStream.ReadAll
                    textStream.Close

                    ' Prüfen ob der Inhalt der Datei und der des Moduls sich unterscheiden falls ja wird die Datei überschrieben
                    If fileContent <> moduleCode Then
                        
                        Dim textStreamOverwrite As Object
                        Set textStreamOverwrite = fs.CreateTextFile(modulePath, True)
                        textStreamOverwrite.Write moduleCode
                        textStreamOverwrite.Close
                    End If
                ' Modul wurde unter dem jetzigen Namen noch nicht exportiert, dementsprechend einfach exportieren.
                Else
                
                ' Neue .bas Datei wird erstellt und mit dem Modul inhalt gefüllt
                Dim textStreamNew As Object
                Set textStreamNew = fs.CreateTextFile(modulePath, True)
            
                textStreamNew.Write moduleCode
                textStreamNew.Close
            
                'Debug.Print "Module Name: " & moduleName
                'Debug.Print moduleCode
                End If
            End If
        End If
    Next vbProj
    
'---------------------------------------------------------------------------------------------
' Aufräumen
    
    Set fs = Nothing
    Set vbComp = Nothing
    Set wb = Nothing
    Set vbProj = Nothing
    Set fileSysObj = Nothing

End Sub





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
            customCommitMessage = UserPromptText("Bitte gebe hier deine Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier angeben")
            
            ' Leere Commit Nachricht prüfen:
            If customCommitMessage = "" Then
                MsgBox "Es wurde keine Commit Nachricht eingegeben der Commit Vorgang wird abgebrochen."
                Exit Sub
            End If
            
            Do While BadCharacterFilter(customCommitMessage, "Commit")
            
                customCommitMessage = UserPromptText("Die eingegebene Commit Nachricht war ungültig. Bitte geben Sie hier deine Commit Nachricht an.", "Custom Commit Nachricht", "Commit Nachricht hier angeben")
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
    
    'temp = ShellCommand(gitCommand, "Die Änderungen wurden commitet.", "Die Änderungen konnten nicht commitet werden. Versuchen Sie es bitte manuell über eine Shellinstanz.")
    
End Sub


