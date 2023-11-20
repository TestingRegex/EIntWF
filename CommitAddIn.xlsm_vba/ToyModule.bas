'''
'
' A module to contain temporary tests that are used to test functions of my own or ones that are new to me,
' this module is regularly reset and cleared.
'
'''

Option Explicit


Sub ModuleSearch()

    
    MsgBox ModulNamenSuchen("pusher")


End Sub

Sub TestRemoveModule()

    RemoveModule ("Module1")

End Sub


Sub TagRetrieval()

' Variablen:


'-------------------------------------------------------
' Pfad:
    Pathing
    
'-------------------------------------------------------

    MsgBox GetShellOutput("git tag")


End Sub

Function GetShellOutput(ByVal command As String)

    Dim shell As Object
    Dim executioner As Object
    Dim output As String
    
    'Shellinstanz erstellen
    Set shell = CreateObject("WScript.Shell")

    ' Command ausführen und Output schnappen
    Set executioner = shell.exec(command)
    output = executioner.StdOut.ReadAll

    ' Return the output
    GetShellOutput = output

End Function