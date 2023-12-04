'''
'   Eine Sammlung von Excel Makros, die die einzelnen Arbeitsschritte zusammenlegen.
'
'
'
'''

Option Explicit

Sub workflowExportCommitPush(ByRef control As Office.IRibbonControl)
On Error GoTo ErrHandler

    Export
    Commit (False)
    Push

ExitSub:
    Exit Sub
    
ErrHandler:
    MsgBox "Something went wrong."
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
    MsgBox "Something went wrong."
    Resume ExitSub
    Resume
    
End Sub