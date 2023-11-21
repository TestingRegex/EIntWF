Option Explicit

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
                        
                        newModuleName = UserInputText("Wie soll das Modul heißen?", "", "")
                        
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
                            newModuleName = UserInputText("Wählen Sie bitte einen neuen Namen für das importierte Modul aus.", "", "Neuer Modulname")
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