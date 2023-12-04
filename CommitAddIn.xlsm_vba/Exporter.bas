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

On Error GoTo ErrHandler:

    AnnoyUsers
    AltExporter
    
ExitSub:
        
    Exit Sub
      
ErrHandler:
    
    MsgBox "Something went wrong."
    Resume ExitSub
    Resume
    
End Sub

'A firts export function, that does its job but does not respect the various vba project component types, replaced by "AltExporter"
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
            
                moduleCode = vbProj.CodeModule.lines(1, vbProj.CodeModule.CountOfLines)

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

' This is a more sophisticated version of the export function that respects the various types of VBA project components!

Function AltExporter()

    Dim wb As Workbook
    Dim vbComp As Object
    Dim suffix As String
    Dim vbaDirectory As String
    Dim fs As Object 'the object that allows us to interact with the FileSystem
    Dim fileContent As String
    Dim moduleContent As String
    Dim modulePath As String
    Dim textStream As Object
    Dim startLine As Integer
    Dim i As Integer
    
    Set wb = ActiveWorkbook
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    vbaDirectory = Replace(wb.path & "\" & wb.Name & "_vba", " ", "_")
    
'---------------------------------------------------------------------------------------------
' Creating the export directory if it does not exist yet
    If Not fs.FolderExists(vbaDirectory) Then
        fs.CreateFolder vbaDirectory
    End If
'---------------------------------------------------------------------------------------------
' The actual export process:

    For Each vbComp In wb.VBProject.VBComponents
    
        Select Case vbComp.Type
            Case 2
                suffix = ".cls" ' Class modules
            Case 3
                suffix = ".frm" ' Userforms
            Case 1
                suffix = ".bas" ' Standard Modules
            Case 100
                suffix = ".cls" 'Objects contained in the "Microsoft Excel Objects" folder
            Case Else
                suffix = ""
        End Select
        

        ' Checking if the component needs to be exported at all, does it contain any code?
        If suffix <> "" And vbComp.CodeModule.CountOfLines > 0 Then
        

            modulePath = vbaDirectory & "\" & _
                            vbComp.Name & suffix
        
            ' Add check to see if code content has changed.
            ' If the module has been exported before we check whether we need to overwrite it incase of a change.
            If fs.FileExists(modulePath) And Dir(modulePath) <> "" Then
                
                
                moduleContent = vbComp.CodeModule.lines(1, vbComp.CodeModule.CountOfLines)
                
                ' Get content of already exported version.
                Set textStream = fs.OpenTextFile(modulePath, 1) ' 1: ForReading
                fileContent = textStream.ReadAll
                textStream.Close
                        
                
                startLine = FindLine(fileContent, "Option Explicit")
                
                If Not startLine = 0 Then
                
                    ' If we find either Option Explicit or ''' at the beginning of a line in an exported module we use this as our first line of code visible in the VBE
                    Set textStream = fs.OpenTextFile(modulePath, 1) ' 1: ForReading
                    ' Not so elegant and could/should be improved
                        For i = 1 To startLine
                            textStream.Skipline
                        Next i
                                
                    fileContent = textStream.ReadAll
                    textStream.Close
                    
                    ' Check if the content of file and component differ, if yes then overwrite file content with component content.
                    If Mid(fileContent, 1, Len(fileContent) - 2) <> moduleContent Then
                        vbComp.Export _
                            filename:=modulePath
                    End If
                End If
            Else
                vbComp.Export _
                        filename:=modulePath
            End If
        End If
    Next vbComp

    'Clean Up
    Set textStream = Nothing
    Set vbComp = Nothing
    Set fs = Nothing
    
End Function


Function FindLine(ByVal content As String, ByVal term As String)
    ' A function that should help with finding the "proper" start to the code often either Option Explicit or a comment, _
    to avoid the overhead lines created when exporting with the inbuild export method.
    
    If term = "" Or content = "" Then
        MsgBox "Invalid input for FindString"
    Else
        Dim lines As Variant
        Dim i As Integer
        
        lines = Split(content, vbCrLf)
        For i = LBound(lines) To UBound(lines)
            If Left(lines(i), Len(term)) = term Or Left(lines(i), 3) = "'''" Then
                FindLine = i
                Exit For
            End If
        Next i
    End If
End Function