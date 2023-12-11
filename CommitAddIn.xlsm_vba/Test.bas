Attribute VB_Name = "Test"
'''
' A test for the exporter
'''
Option Explicit '

' A sub connected to the test button to test new functions once ready

Public Sub TestSub()
On Error GoTo ErrHandler

    Commit True, True
    
ExitSub:
    Exit Sub
    
ErrHandler:
    ErrorHandler Err.Number, Err.Source, Err.Description
    Resume ExitSub
    Resume

End Sub


