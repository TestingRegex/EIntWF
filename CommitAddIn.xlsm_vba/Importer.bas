Attribute VB_Name = "Importer"
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
'
'''

Option Explicit

Private Sub ImportMacros(ByVal control As Office.IRibbonControl)
On Error GoTo ErrHandler

    Import
    
ExitSub:
    Exit Sub
    
ErrHandler:
    
    ErrorHandler Err.Number, Err.Source, Err.Description
    Resume ExitSub
    Resume
 
End Sub

Public Sub Import()

    Dim selectedFolder As String ' Der Pfad zum Importordner
    
    Dim fileSystemObject As Object 'FileSystemObject
    Dim folder As Object 'FileSystemObject: The directory we import from
    Dim file As Object 'FileSystemObject: iterator representing the files we iterate over
    Dim liveWorkbook As Workbook ' Current Workbook
    Dim vbComp As Object ' Component of the current Workbooks vba project.
    Dim moduleName As String ' name of the module we are trying to import
    Dim newModuleName As String ' name of the module we will save the code in
    Dim benutzerMeinung As Long ' Varaible to save user input (yes/no)
    Dim suffix As String

'-----------------------------------------------------------------
' Choosing the directory to import from.

    ' Set a reference to the Microsoft Scripting Runtime library.
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    
    'Dialogue to inform the user of what has to be done.
    MsgBox "Bitte wählen Sie den Ordner aus, aus dem Sie die Makros importieren möchten."
    
    
    ' Get User input for desired import directory
    selectedFolder = SelectFolder()

    ' Check validity of userinput
    If selectedFolder = vbNullString Then
        Err.Raise 1239, "Import Prozess", "Es wurde kein Ordner ausgewählt."
    End If
    
    If Not fileSystemObject.FolderExists(selectedFolder) Then
        Err.Raise 1239, "Import Prozess", "Der gewünschte Ordner konnte nicht gefunden werden."
    End If
    
'------------------------------------------------------------------
' The import process begins

    ' Currently we are only import .bas files.
    Set folder = fileSystemObject.GetFolder(selectedFolder)
    Set liveWorkbook = activeWorkbook
    'Debug.Print "ActiveWorkbook: " & liveWorkbook.Name
    For Each file In folder.Files
        suffix = LCase(Right(file.Name, 4))
        If suffix = ".bas" Or suffix = ".frm" Or suffix = ".cls" Then ' Do not import .frx files or other files that we not exported properly!!!
        'If LCase(Right(file.Name, 4)) = ".bas" Then
        moduleName = Left(file.Name, Len(file.Name) - 4) ' Remove the last 4 characters (".bas") to get module name.
        Debug.Print file.Name
            If LCase(Right(file.Name, 4)) = ".frm" Then ' If we change the name of a userform everything breaks.
                If ModulNamenSuchen(moduleName) Then
                    RemoveModule liveWorkbook, moduleName
                End If
                Set vbComp = liveWorkbook.VBProject.VBComponents.Import(file.path)
                vbComp.Name = moduleName
            Else
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
                        RemoveModule liveWorkbook, moduleName
                        ' Import the .bas file into the workbook's VBA project.
                        Set vbComp = liveWorkbook.VBProject.VBComponents.Import(file.path)
                        vbComp.Name = moduleName
                    
                    Else
                        benutzerMeinung = UserPromptYesNo(" Möchten Sie das Modul '" + moduleName + "' unter einem anderen Namen speichern? (Bei 'Nein' wird das Modul übersprungen.)")
                        If benutzerMeinung = vbYes Then
                            newModuleName = UserPromptText("Wie soll das Modul heißen?", vbNullString, vbNullString, "Module")
                            If newModuleName = vbNullString Then
                                MsgBox "Vorgang abgebrochen."
                                Exit Sub
                            End If
                            Do While ModulNamenSuchen(newModuleName)
                                benutzerMeinung = UserPromptYesNo("Dieser Name ist bereits vergeben. Soll dieses Modul doch Übersprungen werden?")
                                If benutzerMeinung = vbYes Then
                                    Dim skip As Boolean
                                    skip = True
                                    Exit Do
                                End If
                                If Not skip Then
                                newModuleName = UserPromptText("Wählen Sie bitte einen neuen Namen für das importierte Modul aus.", vbNullString, "Neuer Modulname", "Module")
                                    If newModuleName = vbNullString Then
                                        MsgBox "Vorgang abgebrochen."
                                        Exit Sub
                                    End If
                                End If
                            Loop
                            If Not skip Then
                            ' Import the .bas file into the workbook's VBA project.
                            Set vbComp = liveWorkbook.VBProject.VBComponents.Import(file.path)
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
                    Set vbComp = liveWorkbook.VBProject.VBComponents.Import(file.path)
                    vbComp.Name = moduleName
                End If
            End If
        End If
    Next file

    MsgBox "Alle gewünschten VBA-Dateien aus " & selectedFolder & " wurden importiert."

End Sub
