Option Explicit

' A sub connected to the test button to test new functions once ready

Sub TestSub(ByRef control As Office.IRibbonControl)

    alternativeExporter
    'ModuleTypeChange
    
End Sub

Function alternativeExporter()

    Dim wb As Workbook
    Dim vbComp As Object
    Dim suffix As String
    Dim vbaDirectory As String
    Dim fs As Object 'the object that allows us to interact with the FileSystem
    
    Set wb = ActiveWorkbook
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    vbaDirectory = Replace(wb.path & "\" & wb.Name & "_vba\", " ", "_")
    
'---------------------------------------------------------------------------------------------
' Creating the export directory if it does not exist yet
    
    If Not fs.FolderExists(vbaDirectory) Then
        fs.CreateFolder vbaDirectory
    End If

'---------------------------------------------------------------------------------------------
' The actual export process:

    For Each vbComp In wb.VBProject.VBComponents
    
'        Debug.Print "Component name: " & vbComp.Name; "Compenent type: " & vbComp.Type
        
        Select Case vbComp.Type
            Case 2
                suffix = ".cls" ' Class modules
            Case 3
                suffix = ".frm" ' Userforms
            Case 1
                suffix = ".bas" ' Standard Modules
            Case 100
                suffix = ".txt" 'Objects contained in the "Microsoft Excel Objects" folder
            Case Else
                suffix = ""
        End Select
        
        If vbComp.Name = "Module1" Then
            Debug.Print vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
        End If

'        If suffix <> "" And vbComp.CodeModule.CountOfLines > 0 Then
            ' Add check to see if code content has changed.
           
           
'            vbComp.Export _
                filename:=vbaDirectory & "\" & _
                vbComp.Name & suffix
'        Else
'            vbComp.Export _
                filename:=vbaDirectory & "\" & _
                vbComp.Name & ".txt"
'        End If
        
    Next vbComp

End Function

Function ModuleTypeChange()
    
    Dim vbComp As Object
    
    For Each vbComp In ActiveWorkbook.VBProject.VBComponents
        Debug.Print vbComp.Name; "  "; vbComp.Type
        If vbComp.Name = "A_module" Then
            vbComp.Name = "Z_module"
            'vbComp.Type = 2
        End If
    
    Next vbComp

End Function