Attribute VB_Name = "Exporter"
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

    Dim wb As Workbook
    Dim vbComp As Object
    Dim suffix As String
    Dim vbaDirectory As String
    Dim fs As Object 'the object that allows us to interact with the FileSystem
    Dim fileContent As String
    Dim moduleContent As String
    Dim modulePath As String
    Dim textStream As Object
    
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
        
        'If vbComp.Name = "SimpleWorkflows" Then
        'Debug.Print vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
        modulePath = vbaDirectory & "\" & _
                        vbComp.Name & suffix

        If suffix <> "" And vbComp.CodeModule.CountOfLines > 0 Then
        
        ' Add check to see if code content has changed.
        ' If the module has been exported before we check whether we need to overwrite it incase of an update.
            If fs.FileExists(modulePath) And Dir(modulePath) <> "" And Not vbComp.Name = "ThisWorkbook" Then
                
                ' Get content of already exported version.
                    
                Set textStream = fs.OpenTextFile(modulePath, 1) ' 1: ForReading
                textStream.Skipline
                    
                fileContent = textStream.ReadAll
                textStream.Close
                moduleContent = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
                ' Check if the content of file and component differ, if yes then overwrite file content with component content.
                Debug.Print vbComp.Name
                Debug.Print "file"; Len(fileContent)
                Debug.Print "module"; Len(moduleContent)
                If Mid(fileContent, 1, Len(fileContent) - 2) <> moduleContent Then
                                           
                    vbComp.Export _
                        filename:=modulePath
                End If
                    
            End If
        End If
        'End If
    Next vbComp

    'Clean Up
    Set textStream = Nothing
    Set vbComp = Nothing
    Set fs = Nothing
    
End Function
