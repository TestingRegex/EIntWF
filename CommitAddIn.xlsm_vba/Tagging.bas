Attribute VB_Name = "Tagging"
'''
'   Ein Excel Makro was an einen Button im Add-in Tab gebunden ist und
'   die Aufgabe des Taggens der letzten Änderungen und des letzten Commits übernimmt.
'
'   Hier für werden alle Module exportiert und die Änderungen commitet mit einer Standard commit Nachricht.
'   Danach wird der Benutzer dazu aufgefordert eine Tag Nachricht zu erstellen was genau in dieser Version des
'   Codes erreicht wird.
'
'   Verwendete Funktionen:
'       Pathing, UserPromptText
'''

Option Explicit

Private Sub GitTag(ByVal control As Office.IRibbonControl)
On Error GoTo ErrHandler

    If AnnoyUsers = vbYes Then
        Commit (True)
        Tag
    End If
    
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
    Dim shell As Object
    
'------------------------------------------------------
' Find desired path

    Pathing
    
'------------------------------------------------------
' Core:
'

    VersionInput = UserPromptText("Welche Version des Workbooks möchten Sie taggen?", "Versionsname", "_._", "Version")
    
    If VersionInput = vbNullString Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Sub
    End If
    '------------------------------------
    ' Validating userInput to not contain undesirable characters.
    
    
    TagMessage = UserPromptText("Bitte geben Sie eine Kurze Beschreibung der Version oder ihrer Relevanz an:", "Versionsbeschreibung", vbNullString, "Tag")
    If TagMessage = vbNullString Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Sub
    End If
    
    
    gitCommand = "git tag -a " & VersionInput & " -m  """ & TagMessage & " - " & GetUser() & """"
    
    'Debug.Print GitCommand
    
'-------------------------------------------------------------------------
'Commands are passed to the shell

    
        
    ShellCommand gitCommand, "Der Tag wurde erfolgreich erstellt.", "Der Tag konnte nicht erstellt werden.", "Tag"
    
    Set shell = CreateObject("WScript.Shell")
    
    shell.Run "git push origin --tags", vbNormalFocus, True

End Sub
