Option Explicit

Sub ImportMacros_Test()

    Dim selectedFolder As String ' Der Pfad zum Importordner
    
    Dim fs As Object 'FileSystemObject um mit System außerhalb von Excel interagieren zu können
    Dim folder As Object 'FileSystemObject: Der Ordner aus dem imortiert wird
    Dim file As Object 'FileSystemObject: Der Iterator beim Importieren
    Dim wb As Workbook ' Das Aktive Workbook
    Dim vbComp As Object ' Das VBA Projekt des Aktiven Workbooks
    Dim moduleName As String ' Der zukünfitge Name der importierten Module

    ' Set a reference to the Microsoft Scripting Runtime library.
    Set fs = CreateObject("Scripting.FileSystemObject")
    
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

    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".bas" Then
            
             moduleName = Left(file.Name, Len(file.Name) - 4) ' Remove the last 4 characters (".bas").
             
            ' Import the .bas file into the workbook's VBA project.
            Set vbComp = wb.VBProject.VBComponents.Import(file.path)
            vbComp.Name = moduleName
        End If
    Next file

    ' Clean up.
    Set fs = Nothing
    Set folder = Nothing
    Set file = Nothing
    Set wb = Nothing
    Set vbComp = Nothing

    MsgBox "Alle .bas Dateien aus " & selectedFolder & " wurden importiert."
    
End Sub
