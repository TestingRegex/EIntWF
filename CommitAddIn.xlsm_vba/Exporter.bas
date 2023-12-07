Attribute VB_Name = "Exporter"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   This module contains the macros and major functions used in the 'VBA Projekt exportieren'
'   button.
'
'   Purpose:
'       The module exports all vba project components ( this requires access to the vba
'       project object) as the correct file type so that they may be imported correctly later
'       on.
'
'   Verwendete Funktionen:
'       AnnoyUsers, Saver, FindLine
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit


Private Sub ExportSub(ByVal control As Office.IRibbonControl)

On Error GoTo ErrHandler:

    If AnnoyUsers = vbYes Then
        Export
    End If
    
ExitSub:
        
    Exit Sub
      
ErrHandler:
    
    ErrorHandler Err.Number, Err.Source, Err.Description
    Resume ExitSub
    Resume
    
End Sub

'A firts export function, that does its job but does not respect the various vba project component types, replaced by "AltExporter"
Public Sub RetiredExport()

    Dim liveWorkbook As Workbook 'Zeigt auf das aktive Workbook
    Dim WorkbookName As String 'Beinhaltet den namen des Workbooks
    Dim vbComp As Object 'Used as Iterator for components of the VBA project
    Dim moduleName As String 'Contains the name of the current module
    Dim moduleCode As String 'Contains the code within the current module
    Dim outPath As String 'The location of our workbook
    Dim modulePath As String 'the location of the module
    Dim fileSystemObject As Object 'the object that allows us to interact with the FileSystem
    
    
'---------------------------------------------------------------------------------------------
' Save the workbook before we begin the export process.

    Saver

'---------------------------------------------------------------------------------------------
' The path to the export-directory is found

    Set liveWorkbook = activeWorkbook
    WorkbookName = liveWorkbook.Name

    outPath = liveWorkbook.path
    
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    
    Dim vbaDirectory As String
    
    vbaDirectory = Replace(outPath & "\" & WorkbookName & "_vba\", " ", "_")
    
'---------------------------------------------------------------------------------------------
' Creating the export directory if it does not exist yet

    If Not fileSystemObject.FolderExists(vbaDirectory) Then
        fileSystemObject.CreateFolder vbaDirectory
    End If

'---------------------------------------------------------------------------------------------
' The actual export process:


    ' We iterate through all components of the vba project and export them.
    For Each vbComp In liveWorkbook.VBProject.VBComponents
        'If vbProj.Type = 1 Then    ' this line can be uncommented if one wants to export only the modules.
            moduleName = vbComp.Name
            
            ' If the component is empty we do not need to export it.
            If vbComp.CodeModule.CountOfLines > 0 Then
            
                moduleCode = vbComp.CodeModule.lines(1, vbComp.CodeModule.CountOfLines)

                modulePath = vbaDirectory & moduleName & ".bas"

                ' If the module has been exported before we check whether we need to overwrite it incase of an update.
                If fileSystemObject.FileExists(modulePath) And Dir(modulePath) <> vbNullString And Not moduleName = "ThisWorkbook" Then
                
                    ' Get content of already exported version.
                    Dim textStream As Object
                    Set textStream = fileSystemObject.OpenTextFile(modulePath, 1) ' 1: ForReading
                                        
                    Dim fileContent As String
                    fileContent = textStream.ReadAll
                    textStream.Close

                    ' Check if the content of file and component differ, if yes then overwrite file content with component content.
                    If fileContent <> moduleCode Then
                        
                        Dim textStreamOverwrite As Object
                        Set textStreamOverwrite = fileSystemObject.CreateTextFile(modulePath, True)
                        textStreamOverwrite.Write moduleCode
                        textStreamOverwrite.Close
                    End If
                    
                ' Component has not been exported (under its current name) as of yet, so simply export.
                Else
                
                ' new .bas file is created.
                Dim textStreamNew As Object
                Set textStreamNew = fileSystemObject.CreateTextFile(modulePath, True)
            
                textStreamNew.Write moduleCode
                textStreamNew.Close
                
                End If
            End If
        'End If
    Next vbComp
    'Clean Up
    Set textStream = Nothing
    Set vbComp = Nothing
    Set textStreamNew = Nothing
    Set textStreamOverwrite = Nothing
    Set fileSystemObject = Nothing
End Sub

' This is a more sophisticated version of the export function that respects the various types of VBA project components!

Public Sub Export()

    Dim liveWorkbook As Workbook
    Dim vbComp As Object
    Dim suffix As String
    Dim vbaDirectory As String
    Dim fileSystemObject As Object 'the object that allows us to interact with the FileSystem
    Dim fileContent As String
    Dim moduleContent As String
    Dim modulePath As String
    Dim textStream As Object
    Dim startLine As Long
    Dim i As Long
    Dim UserInput As Long
    
    Set liveWorkbook = activeWorkbook
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    
    ' Remove the vbNo definition when other functions are updated to reflect the flexibility as well.
    UserInput = vbNo 'UserPromptYesNo("Möchten Sie Ihr VBA Projekt in einen spezifischen Ordner exportieren?" _
                                    & vbCrLf & "Ansonsten wird das Projekt in den " _
                                    & liveWorkbook.Name & "_vba Ordner exportiert.")
                                    
    If UserInput = vbYes Then
        vbaDirectory = SelectFolder
    Else
        vbaDirectory = Replace(liveWorkbook.path & "\" & liveWorkbook.Name & "_vba", " ", "_")
    End If
'---------------------------------------------------------------------------------------------
' Creating the export directory if it does not exist yet
    If Not fileSystemObject.FolderExists(vbaDirectory) Then
        fileSystemObject.CreateFolder vbaDirectory
    End If
'---------------------------------------------------------------------------------------------
' The actual export process:

    For Each vbComp In liveWorkbook.VBProject.VBComponents
    
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
                suffix = vbNullString
        End Select
        

        ' Checking if the component needs to be exported at all, does it contain any code?
        If suffix <> vbNullString And vbComp.CodeModule.CountOfLines > 0 Then
        

            modulePath = vbaDirectory & "\" & _
                            vbComp.Name & suffix
        
            ' Add check to see if code content has changed.
            ' If the module has been exported before we check whether we need to overwrite it incase of a change.
            If fileSystemObject.FileExists(modulePath) Then
                
                
                moduleContent = vbComp.CodeModule.lines(1, vbComp.CodeModule.CountOfLines)
                
                ' Get content of already exported version.
                Set textStream = fileSystemObject.OpenTextFile(modulePath, 1) ' 1: ForReading
                fileContent = textStream.ReadAll
                textStream.Close
                        
                
                startLine = FindLine(fileContent, "Option Explicit")
                If Not startLine = -1 Then
                
                    ' If we find either Option Explicit or ''' at the beginning of a line in an exported module we use this as our first line of code visible in the VBE
                    Set textStream = fileSystemObject.OpenTextFile(modulePath, 1) ' 1: ForReading
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
    Set fileSystemObject = Nothing
    
End Sub
