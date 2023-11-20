'''
'
' A module to contain temporary tests that are used to test functions of my own or ones that are new to me,
' this module is regularly reset and cleared.
'
'''

Option Explicit


Sub TagRetrieval()

' Variablen:
    Dim temp As Integer

'-------------------------------------------------------
' Pfad:
    Pathing
    
'-------------------------------------------------------
    
    temp = ShellCommand("git show v1.0 > temp/", "Yay", "Nay")
    
End Sub

Sub Testing()

    ChDir "C:\Users\d60157\"
    MsgBox GetShellOutput("cmd /c cd")
    
End Sub

