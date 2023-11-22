'''
'
' A module to contain temporary tests that are used to test functions of my own or ones that are new to me,
' this module is regularly reset and cleared.
'
'''

Option Explicit




Sub Testing()

    Dim gitCommand As String
    Dim temp As Integer
    
    gitCommand = "git tag"
    
    MsgBox GetShellOutput(gitCommand)
    
End Sub


