Option Explicit

Sub Push_Test()

    Dim GitCommand As String
    Dim WorkbookPath As String

'------------------------------------------------------------------------
' Das richtige Directory finden

    Pathing
    
'-----------------------------------------------------------------------
' git push ausführen
    
    GitCommand = "git push"
    temp = ShellCommand(GitCommand, "Committed Änderungen wurden gepusht.", "Der Push-Vorgang ist gescheitert.")

End Sub