Attribute VB_Name = "SimpleWorkflows"
'''
'   Eine Sammlung von Excel Makros, die die einzelnen Arbeitsschritte zusammenlegen.
'
'
'
'''

Option Explicit

Sub workflowExportCommitPush(ByRef control As Office.IRibbonControl)
On Error GoTo ErrHandler
    If AnnoyUsers = vbYes Then
        Export
        Commit (False)
        Push
    End If
    
ExitSub:
    Exit Sub
    
ErrHandler:
    MsgBox "Im " & Err.Source & " Vorgang ist ein Fehler aufgetreten." & vbCrLf & Err.Description
    Resume ExitSub
    Resume

End Sub


Sub workflowPullImport(ByRef control As Office.IRibbonControl)
On Error GoTo ErrHandler

    Pull
    Import

ExitSub:
    Exit Sub
    
ErrHandler:
    MsgBox "Im " & Err.Source & " Vorgang ist ein Fehler aufgetreten." & vbCrLf & Err.Description
    Resume ExitSub
    Resume
    
End Sub
