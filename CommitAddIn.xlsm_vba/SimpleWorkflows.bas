Attribute VB_Name = "SimpleWorkflows"
'''
'   Eine Sammlung von Excel Makros, die die einzelnen Arbeitsschritte zusammenlegen.
'
'
'
'''

Option Explicit

Private Sub workflowExportCommitPush(ByVal control As Office.IRibbonControl)
On Error GoTo ErrHandler

    Export
    Commit (False)
    Push
        
ExitSub:
    Exit Sub
    
ErrHandler:

    ErrorHandler Err.Number, Err.Source, Err.Description
    Resume ExitSub
    Resume
    
End Sub


Private Sub workflowPullImport(ByVal control As Office.IRibbonControl)
On Error GoTo ErrHandler

    Pull
    Import

ExitSub:
    Exit Sub
    
ErrHandler:

    ErrorHandler Err.Number, Err.Source, Err.Description
    Resume ExitSub
    Resume
    
End Sub
