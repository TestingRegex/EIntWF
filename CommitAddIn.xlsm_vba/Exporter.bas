Sub ExtractModulesFromWorkbook(control As Office.IRibbonControl)

    Dim wb As Workbook
    Dim workbookName As String
    Dim vbComp As Object
    Dim vbProj As Object
    Dim moduleName As String
    Dim moduleCode As String
    Dim outPath As String
    Dim modulePath As String
    Dim fileSysObj As Object


    Set wb = ActiveWorkbook
    workbookName = wb.Name

    ' Get the current Directory
    outPath = wb.Path
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim vbaDirectory As String
    vbaDirectory = outPath & "\" & workbookName & "_vba\"

    ' Check if the directory exists; if not, create it
    If Not fs.FolderExists(vbaDirectory) Then
        fs.CreateFolder vbaDirectory
    End If

    ' Iterate through each component of the VB Project attached to the current workbook and
    ' extract all of those that are modules.
    For Each vbProj In wb.VBProject.VBComponents
        If vbProj.Type = 1 Then ' Module
            moduleName = vbProj.Name
            
            ' Check whether the module contains any lines of code and is even worth exporting
            If vbProj.CodeModule.CountOfLines > 0 Then
            
                moduleCode = vbProj.CodeModule.Lines(1, vbProj.CodeModule.CountOfLines)
            
                ' Save the module code or do something with it as needed
                modulePath = vbaDirectory & moduleName & ".bas"
                
                ' Check whether the module has changed since it was last exported
                If fs.FileExists(modulePath) Then
                    ' Read the content of the .bas file
                    Dim textStream As Object
                    Set textStream = fs.OpenTextFile(modulePath, 1) ' 1: ForReading

                    ' Read the entire content of the .bas file
                    Dim fileContent As String
                    fileContent = textStream.ReadAll
                    textStream.Close

                    ' Compare the file content with the current module code
                    If fileContent <> moduleCode Then
                        ' Module code has changed so we overwrite the old code with the new
                        Dim textStreamOverwrite As Object
                        Set textStreamOverwrite = fs.CreateTextFile(modulePath, True)
                        textStreamOverwrite.Write moduleCode
                        textStreamOverwrite.Close
                    End If
                Else
                ' .bas file doesn't exist, indicating a change or the module has not been exported yet
                
                ' Create a new .bas file to save the content of the modules into
                Dim textStreamNew As Object
                Set textStreamNew = fs.CreateTextFile(modulePath, True)
            
                ' Write the module code into the .bas file
                textStreamNew.Write moduleCode
                textStreamNew.Close
            
                Debug.Print "Module Name: " & moduleName
                Debug.Print moduleCode
                End If
            End If
        End If
    Next vbProj
    
    Set fs = Nothing
End Sub
