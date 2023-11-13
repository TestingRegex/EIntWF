'''
'   Eine Sammlung von Excel Makros, die die einzelnen Arbeitsschritte zusammenlegen.
'
'
'
'''

Option Explicit

Sub workflowExportCommitPush(ByRef control As Office.IRibbonControl)
    
    Export
    Commit
    Push

End Sub


Sub workflowPullImport(ByRef control As Office.IRibbonControl)

    Pull
    Import

End Sub