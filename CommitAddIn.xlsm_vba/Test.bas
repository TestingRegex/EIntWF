'''
' A test for the exporter
'''
Option Explicit

' A sub connected to the test button to test new functions once ready

Sub TestSub(ByRef control As Office.IRibbonControl)

    AltExporter
    
End Sub



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
    
    vbaDirectory = Replace(wb.path & "\" & wb.Name & "_vbaTesting\", " ", "_")
    
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
                suffix = ".txt" 'Objects contained in the "Microsoft Excel Objects" folder
            Case Else
                suffix = ""
        End Select
        

        'If vbComp.Name = "Module1" Or vbComp.Name = "Module2" Then
        If suffix <> "" And vbComp.CodeModule.CountOfLines > 0 Then
        

            modulePath = vbaDirectory & "\" & _
                            vbComp.Name & suffix
        
            ' Add check to see if code content has changed.
            ' If the module has been exported before we check whether we need to overwrite it incase of an update.
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
                    'Not so elegant and could/should be improved
                        For i = 1 To startLine
                            textStream.Skipline
                        Next i
                                
                    fileContent = textStream.ReadAll
                    textStream.Close
                    
                    ' Check if the content of file and component differ, if yes then overwrite file content with component content.
                    If Mid(fileContent, 1, Len(fileContent) - 2) <> moduleContent Then
                    
                        Debug.Print vbComp.Name
                        Debug.Print startLine
                        vbComp.Export _
                            filename:=modulePath
                    End If
                End If
            Else
                vbComp.Export _
                        filename:=modulePath
            End If
        End If
        'End If
    Next vbComp

    'Clean Up
    Set textStream = Nothing
    Set vbComp = Nothing
    Set fs = Nothing
    
End Function

Function FindLine(ByVal content As String, ByVal term As String)
    
    If term = "" Or content = "" Then
        MsgBox "Invalid input for FindString"
    Else
        Dim lines As Variant
        Dim i As Integer
        
        lines = Split(content, vbCrLf)
        For i = LBound(lines) To UBound(lines)
            If Left(lines(i), Len(term)) = term Or Left(lines(i), 3) = "'''" Then
                'Debug.Print lines(i)
                FindLine = i
                Exit For
            End If
        Next i
    End If
    'Debug.Print "FindLine value: "; FindLine
End Function


Function AnnoyUsers()

    MsgBox "Have you cleaned up your code and spreadsheets?"

End Function
