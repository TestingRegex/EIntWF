'''
'   Ein Excel Makro was an einen Button im Add-in Tab gebunden ist und
'   die Aufgabe des Taggens der letzten Änderungen und des letzten Commits übernimmt.
'
'   Hier für werden alle Module exportiert und die Änderungen commitet mit einer Standard commit Nachricht.
'   Danach wird der Benutzer dazu aufgefordert eine Tag Nachricht zu erstellen was genau in dieser Version des
'   Codes erreicht wird.
'
'   Verwendete Funktionen:
'       Pathing, UserInputText, BadCharacterFilter
'''

Option Explicit

Sub TagCommit(ByRef control As Office.IRibbonControl)
    
    Commit (True)
    Tag

End Sub


Function Tag()

' Benötigten Variablen init:

    Dim GitCommand As String
    Dim VersionInput As String
    Dim TagMessage As String
    Dim StringCheck As Boolean
    Dim shell As Object
    
'------------------------------------------------------
' Git-Pfad finden

    Pathing
    
'------------------------------------------------------
' Core:
'       Es wird auch noch geprüft ob der UserInput kosher ist.

    VersionInput = UserInputText("Welche Version des Workbooks möchten Sie taggen?", "Versionsname", "_._")
    StringCheck = BadCharacterFilter(VersionInput, "Tag")
    If VersionInput = "" Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Function
    End If
    Do While StringCheck
        VersionInput = UserInputText("Der Eingebene Versionsname ist ungültig. Bitte geben Sie einen anderen Namen ein und vermeiden Sie die Zeichen: ' ~!@#$%^&*()+,{}[]|\;:'""<>/?='", "Versionsname", "_._")
        StringCheck = BadCharacterFilter(VersionInput, "Tag")
        If VersionInput = "" Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Function
    End If
    Loop
    
    
    TagMessage = UserInputText("Bitte geben Sie eine Kurze Beschreibung der Version oder ihrer Relevanz an:", "Versionsbeschreibung", "")
    StringCheck = BadCharacterFilter(TagMessage)
    If TagMessage = "" Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Function
    End If
    Do While StringCheck
        TagMessage = UserInputText("Bitte geben Sie eine Kurze Beschreibung der Version oder ihrer Relevanz an:", "Versionsbeschreibung", "")
        StringCheck = BadCharacterFilter(TagMessage)
        If TagMessage = "" Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Function
        End If
    Loop
    
    GitCommand = "git tag -a " & VersionInput & " -m  """ & TagMessage & " - " & GetUser() & """"
    
    'Debug.Print GitCommand
    
'-------------------------------------------------------------------------
'Commands werden an die Shell weitergegeben

    Dim temp As Integer
        
    temp = ShellCommand(GitCommand, "Der Tag wurde erfolgreich erstellt.", "Der Tag konnte nicht erstellt werden.")
    
    Set shell = CreateObject("WScript.Shell")
    
    temp = shell.Run("git push origin --tags", vbNormalFocus, True)

End Function