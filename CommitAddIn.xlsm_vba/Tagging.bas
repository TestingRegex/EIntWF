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

Sub GitTag(ByRef control As Office.IRibbonControl)
    
    Commit (True)
    Tag

End Sub




Function Tag()

' Variables:

    Dim gitCommand As String
    Dim VersionInput As String
    Dim TagMessage As String
    Dim StringCheck As Boolean
    Dim shell As Object
    
'------------------------------------------------------
' Find desired path

    Pathing
    
'------------------------------------------------------
' Core:
'

    VersionInput = UserPromptText("Welche Version des Workbooks möchten Sie taggen?", "Versionsname", "_._", "Tag")
    
    If VersionInput = "" Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Function
    End If
    '------------------------------------
    ' Validating userInput to not contain undesirable characters.
    
    
    TagMessage = UserPromptText("Bitte geben Sie eine Kurze Beschreibung der Version oder ihrer Relevanz an:", "Versionsbeschreibung", "")
    If TagMessage = "" Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Function
    End If
    
    
    gitCommand = "git tag -a " & VersionInput & " -m  """ & TagMessage & " - " & GetUser() & """"
    
    'Debug.Print GitCommand
    
'-------------------------------------------------------------------------
'Commands are passed to the shell

    Dim temp As Integer
        
    temp = ShellCommand(gitCommand, "Der Tag wurde erfolgreich erstellt.", "Der Tag konnte nicht erstellt werden.")
    
    Set shell = CreateObject("WScript.Shell")
    
    temp = shell.Run("git push origin --tags", vbNormalFocus, True)

End Function
