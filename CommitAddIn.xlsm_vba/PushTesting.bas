Option Explicit

Sub Push_Test()

    Dim GitCommand As String
    Dim WorkbookPath As String

'------------------------------------------------------------------------
' Das richtige Directory finden

    Pathing
    
'-----------------------------------------------------------------------
' git push ausf�hren
    
    GitCommand = "git push"
    temp = ShellCommand(GitCommand, "Committed �nderungen wurden gepusht.", "Der Push-Vorgang ist gescheitert.")

End Sub