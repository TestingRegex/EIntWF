Option Explicit

Sub ImportMacros_Test()

    Dim selectedFolder As String ' Der Pfad zum Importordner
    
    Dim fs As Object 'FileSystemObject um mit System au�erhalb von Excel interagieren zu k�nnen
    Dim folder As Object 'FileSystemObject: Der Ordner aus dem imortiert wird
    Dim file As Object 'FileSystemObject: Der Iterator beim Importieren
    Dim wb As Workbook ' Das Aktive Workbook
    Dim vbComp As Object ' Das VBA Projekt des Aktiven Workbooks
    Dim moduleName As String ' Der zuk�nfitge Name der importierten Module

    ' Set a reference to the Microsoft Scripting Runtime library.
    Set fs = CreateObject("Scripting.FileSystemObject")
    
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
