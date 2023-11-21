Option Explicit

Sub ImportMacros_Test()

    Dim selectedFolder As String ' Der Pfad zum Importordner
    
    Dim fs As Object 'FileSystemObject um mit System au�erhalb von Excel interagieren zu k�nnen
    Dim folder As Object 'FileSystemObject: Der Ordner aus dem imortiert wird
    Dim file As Object 'FileSystemObject: Der Iterator beim Importieren
    Dim wb As Workbook ' Das Aktive Workbook
    Dim vbComp As Object ' Eine VBA Componente des Aktiven Workbooks
    Dim moduleName As String ' Der Name der importierten Module
    Dim newModuleName As String ' Der neue Modul Name des zu importierenden Moduls
    Dim benutzerMeinung As Long ' Entscheidung ob bereitsvorhandene Module �berschrieben werden sollen oder nicht

    ' Set a reference to the Microsoft Scripting Runtime library.
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    
'---------------------------------------------------------------------------------------------------------
' Der Ordner aus dem importiert werden soll, wird ausgew�hlt.

    'Dialog damit User wei� was gleich zu tun ist.
    MsgBox "Bitte w�hlen Sie den Ordner aus, aus dem Sie die Makros importieren m�chten."
    
    
    ' Der Pfad zum gew�nschten Importordner wird erhoben.
    selectedFolder = SelectFolder()

    ' Falls kein Ordner ausgesucht wird, brechen wir ab.
    If selectedFolder = "" Then
        MsgBox "Kein Ordner ausgew�hlt. Import abgebrochen."
        Exit Sub
    End If
    
    ' Pr�fen ob der ausgew�hlte Ordner existiert
    If Not fs.FolderExists(selectedFolder) Then
        MsgBox "Der gew�nschte Ordner konnte nicht gefunden werden."
        Exit Sub
    End If
    
    
    ' Alle .bas Dateien werden aus dem Ordner importiert
    Set folder = fs.GetFolder(selectedFolder)
    Set wb = ActiveWorkbook

'---------------------------------------------------------------------------------------------------------
' .bas Dateien werden aus dem gew�nschten Ordner importiert.

    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".bas" Then
             moduleName = Left(file.Name, Len(file.Name) - 4) ' Remove the last 4 characters (".bas").
            '-----------------------------------------------------------------------------------
            ' Namengebung des importierten Moduls:
            If ModulNamenSuchen(moduleName) Then
            
                benutzerMeinung = UserPromptYesNo(" Es gibt bereits ein Modul mit dem Namen '" + moduleName + "'. Soll das bereitsexistierende Modul �berschrieben werden?")
                
                If benutzerMeinung = vbYes Then
                    
                    ' Remove the old Modul
                    RemoveModule (moduleName)
                    ' Import the .bas file into the workbook's VBA project.
                    Set vbComp = wb.VBProject.VBComponents.Import(file.path)
                    vbComp.Name = moduleName
                
                Else
                    
                    benutzerMeinung = UserPromptYesNo(" M�chten Sie das Modul '" + moduleName + "' unter einem anderen Namen speichern? (Bei 'Nein' wird das Modul �bersprungen.)")
                    
                    If benutzerMeinung = vbYes Then
                        
                        newModuleName = UserInputText("Wie soll das Modul hei�en?", "", "")
                        
                        Do While ModulNamenSuchen(newModuleName)
                            '---------------------------------------------------------------------------------------------------------
                            ' Soll die Datei doch nicht importiert werden?
                            benutzerMeinung = UserPromptYesNo("Dieser Name ist bereits vergeben. Soll dieses Modul doch �bersprungen werden?")
                            
                            If benutzerMeinung = vbYes Then
                                
                                Dim skip As Boolean
                                skip = True
                                Exit Do
                            
                            End If
                            
                            If Not skip Then
                            newModuleName = UserInputText("W�hlen Sie bitte einen neuen Namen f�r das importierte Modul aus.", "", "Neuer Modulname")
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

    MsgBox "Alle gew�nschten .bas Dateien aus " & selectedFolder & " wurden importiert."
    
End Sub