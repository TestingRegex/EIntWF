' Wir wollen die Komponenten des VBA-Projekts exportieren ohne auf das VBA Projekt als Objekt
' zu zu greifen.

Option Explicit
Sub Export()

    Dim file As String
    

End Sub



Private Function exportvba(Path As String)
    Dim objVbComp As VBComponent
    Dim strPath As String
    Dim varItem As Variant
    Dim fso As New FileSystemObject
    Dim filename As String
    
    filename = fso.GetFileName(Path)
    
    On Error Resume Next
        MkDir ("C:\Create\directory\for\VBA\Code\" & filename & "\")
    On Error GoTo 0
    
    'Change the path to suit the users needs
    strPath = "C:\Give\directory\to\save\Code\in\" & filename & "\"
    
      For Each varItem In ActiveWorkbook.VBProject.VBComponents
      Set objVbComp = varItem
    
      Select Case objVbComp.Type
         Case vbext_ct_StdModule
            objVbComp.Export strPath & "\" & objVbComp.Name & ".bas"
         Case vbext_ct_Document, vbext_ct_ClassModule
            ' ThisDocument and class modules
            objVbComp.Export strPath & "\" & objVbComp.Name & ".cls"
         Case vbext_ct_MSForm
            objVbComp.Export strPath & "\" & objVbComp.Name & ".frm"
         Case Else
            objVbComp.Export strPath & "\" & objVbComp.Name
      End Select
    Next varItem
End Function