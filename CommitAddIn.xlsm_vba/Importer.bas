Option Explicit

Sub ImportMacros(ByRef control As Office.IRibbonControl)

    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    Dim directory As String
    
     
    Set oFSO = CreateObject("Scripting.FileSystemObject")
     
    Set oFolder = oFSO.GetFolder(ActiveWorkbook.Path & "\VisualBasic")
     
    For Each oFile In oFolder.Files
     
        directory = ActiveWorkbook.Path & "\VisualBasic\" & oFile.Name
        ActiveWorkbook.VBProject.VBComponents.Import directory
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to import " & oFile.Name, vbCritical)
        End If
     
    Next oFile

End Sub