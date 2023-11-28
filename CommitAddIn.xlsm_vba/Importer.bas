'''
'   Ein Excel Makro was an einen Button im Add-in Tab gebunden ist und
'   die Aufgabe des Importieren der Module eines bestimmten Ordners übernimmt
'
'   Allgemeines:
'       Das Programm benötigt zugriff auf das VBA-Projekt als Objekt, um die externen .bas Dateien
'       als VBA-Module ins VBA-Projekt speichern zu können. Dies muss im Trust-Center bei den Makro Einstellungen genehmigt werden.
'
'   Verwendete Funktionen:
'       SelectFolder,ModulNamenSuchen, UserPromptYesNo, UserPromptText,
'''

Option Explicit

Sub ImportMacros(ByRef control As Office.IRibbonControl)

    Import
    
End Sub

Function Import()

    Dim selectedFolder As String ' Der Pfad zum Importordner
    
    Dim fs As Object 'FileSystemObject
    Dim folder As Object 'FileSystemObject: The directory we import from
    Dim file As Object 'FileSystemObject: iterator representing the files we iterate over
    Dim wb As Workbook ' Current Workbook
    Dim vbComp As Object ' Component of the current Workbooks vba project.
    Dim moduleName As String ' name of the module we are trying to import
    Dim newModuleName As String ' name of the module we will save the code in
    Dim benutzerMeinung As Long ' Varaible to save user input (yes/no)

'-----------------------------------------------------------------
' Choosing the directory to import from.

    ' Set a reference to the Microsoft Scripting Runtime library.
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    'Dialogue to inform the user of what has to be done.
    MsgBox "Bitte wählen Sie den Ordner aus, aus dem Sie die Makros importieren möchten."
    
    
    ' Get User input for desired import directory
    selectedFolder = SelectFolder()

    ' Check validity of userinput
    If selectedFolder = "" Then
        MsgBox "Kein Ordner ausgewählt. Import abgebrochen."
        Exit Function
    End If
    
    If Not fs.FolderExists(selectedFolder) Then
        MsgBox "Der gewünschte Ordner konnte nicht gefunden werden."
        Exit Function
    End If
    
'------------------------------------------------------------------
' The import process begins

    ' Currently we are only import .bas files.
    Set folder = fs.GetFolder(selectedFolder)
    Set wb = ActiveWorkbook
    'Debug.Print "ActiveWorkbook: " & wb.Name
    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".bas" Then
             moduleName = Left(file.Name, Len(file.Name) - 4) ' Remove the last 4 characters (".bas") to get module name.
             
            '-----------------------------------------------------------------------------------
            ' VBA does not do well with placing modules
            If moduleName = "ThisWorkbook" Or Left(moduleName, 5) = "Sheet" Then
                moduleName = moduleName & "_import"
            End If
            
            '-----------------------------------------------------------------------------------
            ' The module name is not allowed to already be ascribed to a module in the current workbook the following tries to resolve this conflict:
            If ModulNamenSuchen(moduleName) Then
                benutzerMeinung = UserPromptYesNo("Es gibt bereits ein Modul mit dem Namen '" + moduleName + "'. Soll das bereitsexistierende Modul überschrieben werden?")
                If benutzerMeinung = vbYes Then
                    MsgBox "Sie wollen ein altes Modul überschreiben."
                    ' Remove the old Modul
                    RemoveModule wb, moduleName
                    ' Import the .bas file into the workbook's VBA project.
                    Set vbComp = wb.VBProject.VBComponents.Import(file.path)
                    vbComp.Name = moduleName
                
                Else
                    benutzerMeinung = UserPromptYesNo(" Möchten Sie das Modul '" + moduleName + "' unter einem anderen Namen speichern? (Bei 'Nein' wird das Modul übersprungen.)")
                    If benutzerMeinung = vbYes Then
                        newModuleName = UserPromptText("Wie soll das Modul heißen?", "", "")
                        Do While ModulNamenSuchen(newModuleName)
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
                        ' Import the .bas file into the workbook's VBA project.
                        Set vbComp = wb.VBProject.VBComponents.Import(file.path)
                        vbComp.Name = newModuleName
                        End If
                    Else
                        MsgBox "Das Modul '" + moduleName + "' wird nicht neu importiert."
                    End If
                End If
            Else
            '----------------------------------------------------
            ' No module of the same name exists in the current workbook
            
                'Debug.Print moduleName
                Set vbComp = wb.VBProject.VBComponents.Import(file.path)
                vbComp.Name = moduleName
            End If
        End If
    Next file

    MsgBox "Alle gewünschten .bas Dateien aus " & selectedFolder & " wurden importiert."

End Function