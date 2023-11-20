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
    shell GitCommand, vbNormalFocus
    
    MsgBox "Committed Änderungen wurden gepusht."

End Sub