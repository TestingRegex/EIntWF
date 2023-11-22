Option Explicit

Sub Push_Test()

    Dim gitCommand As String
    Dim temp As Integer

'------------------------------------------------------------------------
' Das richtige Directory finden

    Pathing
    
'-----------------------------------------------------------------------
' git push ausführen
    
    gitCommand = "git push"
    temp = ShellCommand(gitCommand, "Committed Änderungen wurden gepusht.", "Der Push-Vorgang ist gescheitert.")

End Sub