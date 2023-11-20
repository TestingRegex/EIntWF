Option Explicit

Sub Tag_Test()

' Benötigten Variablen init:

    Dim GitCommand As String
    Dim VersionInput As String
    Dim TagMessage As String
    Dim StringCheck As Boolean
    
'------------------------------------------------------
' Git-Pfad finden

    Pathing
    
'------------------------------------------------------
' Basic Ablauf:

    VersionInput = UserInputText("Welche Version des Workbooks möchten Sie taggen?", "Versionsname", "_._")
    StringCheck = BadCharacterFilter(VersionInput, "Tag")
    If VersionInput = "" Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Sub
    End If
    Do While StringCheck
        VersionInput = UserInputText("Der Eingebene Versionsname ist ungültig. Bitte geben Sie einen anderen Namen ein und vermeiden Sie die Zeichen: ' ~!@#$%^&*()+,{}[]|\;:'""<>/?='", "Versionsname", "_._")
        StringCheck = BadCharacterFilter(VersionInput, "Tag")
        If VersionInput = "" Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Sub
    End If
    Loop
    
    
    TagMessage = UserInputText("Bitte geben Sie eine Kurze Beschreibung der Version oder ihrer Relevanz an:", "Versionsbeschreibung", "")
    StringCheck = BadCharacterFilter(TagMessage)
    If TagMessage = "" Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Sub
    End If
    Do While StringCheck
        TagMessage = UserInputText("Bitte geben Sie eine Kurze Beschreibung der Version oder ihrer Relevanz an:", "Versionsbeschreibung", "")
        StringCheck = BadCharacterFilter(TagMessage)
        If TagMessage = "" Then
        MsgBox "Der Tag Vorgang wird abgebrochen."
        Exit Sub
        End If
    Loop
    
    GitCommand = "git tag -a " & VersionInput & " -m  """ & TagMessage & " - " & GetUser() & """"
    
    'MsgBox GitCommand
    shell GitCommand, vbNormalFocus


End Sub