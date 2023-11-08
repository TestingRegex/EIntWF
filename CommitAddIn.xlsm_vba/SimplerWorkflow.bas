'''
'   Ein Excel Makro was an einen Button im Add-in Tab gebunden ist und
'   die Aufgabe des Export, Commit, und Push aufeinmal übernimmt
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