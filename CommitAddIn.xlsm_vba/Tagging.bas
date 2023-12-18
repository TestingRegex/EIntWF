Attribute VB_Name = "Tagging"
'''
'   Ein Excel Makro was an einen Button im Add-in Tab gebunden ist und
'   die Aufgabe des Taggens der letzten Änderungen und des letzten Commits übernimmt.
'
'   Hier für werden alle Module exportiert und die Änderungen committet mit einer Standard Commit-Nachricht.
'   Danach wird der Benutzer dazu aufgefordert eine Tag Nachricht zu erstellen was genau in dieser Version des
'   Codes erreicht wird.
'
'   Verwendete Funktionen:
'       Pathing, UserPromptText
'''

Option Explicit

Private Sub GitTag(ByVal control As Office.IRibbonControl)
On Error GoTo ErrHandler

    Export
    Commit True
    Tag
        
    
ExitSub:
    Exit Sub
    
ErrHandler:
    
    ErrorHandler Err.Number, Err.Source, Err.Description
    Resume ExitSub
    Resume

End Sub




Public Sub Tag()
' Variables:

    Dim gitCommand As String
    Dim VersionInput As String
    Dim TagMessage As String
    Dim userYesNo As Long
    
'------------------------------------------------------
' Find desired path

    Pathing
    
'------------------------------------------------------
' Core:
'

    VersionInput = UserPromptText("Wie soll diese Version des Workbooks heissen?", "Versionsname", "v_._", "Version")
    
    
    userYesNo = UserPromptYesNo("Möchten Sie eine eigene Versionsbeschreibung schreiben? (Empfohlen: Ja)")
    If userYesNo = vbYes Then
        TagMessage = UserPromptText("Bitte geben Sie eine Kurze Beschreibung der Version oder ihrer Relevanz an:", "Versionsbeschreibung", vbNullString, "Tag")
    Else
        TagMessage = "Version erstellt am " & Replace(Date, ".", "_")
    End If
    
    gitCommand = "git tag -a " & VersionInput & " -m  """ & TagMessage & " - " & GetUser() & """"
    
    
'-------------------------------------------------------------------------
'Commands are passed to the shell
        
    ShellCommand gitCommand, "Die Version wurde erfolgreich erstellt.", "Die Version konnte nicht erstellt werden.", "Tag"
    
    ShellCommand "git push origin --tags", "Die Version wurde hochgeladen.", "Die Version konnte nicht hochgeladen werden." & vbCrLf & "Bitte versuchen Sie es über die Commandline mit dem Befehl: " & vbCrLf & "git push origin --tags", "Tag"

End Sub
