'''
'   Ein Excel Makro was an einen Button im Add-in Tab gebunden ist und
'   die Aufgabe des Exportieren der Module im VBA Projekt übernimmt
'
'   Allgemeines:
'       Das Programm benötigt zugriff auf das VBA-Projekt als Objekt, um die externen .bas Dateien
'       als VBA-Module ins VBA-Projekt speichern zu können. Dies muss im Trust-Center bei den Makro Einstellungen genehmigt werden.
'
'   Verwendete Funktionen:
'       Saver,
'''

Option Explicit


Sub ExportSub(control As Office.IRibbonControl)

    Export
    
End Sub

Function Export()

    Dim wb As Workbook 'Zeigt auf das aktive Workbook
    Dim WorkbookName As String 'Beinhaltet den namen des Workbooks
    Dim vbComp As Object 'Used as Iterator for components of the VBA project
    Dim vbProj As Object ' Points to the vba project
    Dim moduleName As String 'Contains the name of the current module
    Dim moduleCode As String 'Contains the code within the current module
    Dim outPath As String 'The location of our workbook
    Dim modulePath As String 'the location of the module
    Dim fs As Object 'the object that allows us to interact with the FileSystem
    
    
'---------------------------------------------------------------------------------------------
' Save the workbook before we begin the export process.

    Saver

'---------------------------------------------------------------------------------------------
' The path to the export-directory is found

    Set wb = ActiveWorkbook
    WorkbookName = wb.Name

    outPath = wb.path
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim vbaDirectory As String
    
    vbaDirectory = Replace(outPath & "\" & WorkbookName & "_vba\", " ", "_")
    
'---------------------------------------------------------------------------------------------
' Creating the export directory if it does not exist yet

    If Not fs.FolderExists(vbaDirectory) Then
        fs.CreateFolder vbaDirectory
    End If

'---------------------------------------------------------------------------------------------
' The actual export process:


    ' We iterate through all components of the vba project and export them.
    For Each vbProj In wb.VBProject.VBComponents
        'If vbProj.Type = 1 Then    ' this line can be uncommented if one wants to export only the modules.
            moduleName = vbProj.Name
            
            ' If the component is empty we do not need to export it.
            If vbProj.CodeModule.CountOfLines > 0 Then
            
                moduleCode = vbProj.CodeModule.Lines(1, vbProj.CodeModule.CountOfLines)

                modulePath = vbaDirectory & moduleName & ".bas"

                ' If the module has been exported before we check whether we need to overwrite it incase of an update.
                If fs.FileExists(modulePath) And Dir(modulePath) <> "" And Not moduleName = "ThisWorkbook" Then
                
                    ' Get content of already exported version.
                    Dim textStream As Object
                    Set textStream = fs.OpenTextFile(modulePath, 1) ' 1: ForReading
                                        
                    Dim fileContent As String
                    fileContent = textStream.ReadAll
                    textStream.Close

                    ' Check if the content of file and component differ, if yes then overwrite file content with component content.
                    If fileContent <> moduleCode Then
                        
                        Dim textStreamOverwrite As Object
                        Set textStreamOverwrite = fs.CreateTextFile(modulePath, True)
                        textStreamOverwrite.Write moduleCode
                        textStreamOverwrite.Close
                    End If
                    
                ' Component has not been exported (under its current name) as of yet, so simply export.
                Else
                
                ' new .bas file is created.
                Dim textStreamNew As Object
                Set textStreamNew = fs.CreateTextFile(modulePath, True)
            
                textStreamNew.Write moduleCode
                textStreamNew.Close
            
                Debug.Print "Module Name: " & moduleName
                Debug.Print moduleCode
                End If
            End If
        'End If
    Next vbProj
    'Clean Up
    Set textStream = Nothing
    Set vbProj = Nothing
    Set textStreamNew = Nothing
    Set textStreamOverwrite = Nothing
    Set fs = Nothing
End Function
