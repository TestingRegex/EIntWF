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
    shell GitCommand, vbNormalFocus
    
    MsgBox "Committed �nderungen wurden gepusht."

End Sub