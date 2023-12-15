Attribute VB_Name = "FunktionenModul"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   A module that collects all homemade functions that take over a minor part of some process.
'
'   The following is a list of all functions in this module with a brief description and a
'   list of module where the function is called
'
'
'   GetUser
'       Description: A short function that retrieves the environment username
'       Called in: Committer, Tagging
'
'   Saver
'       Description: A short function that saves the active workbook
'       Called in: Committer, Exporter
'
'   Pathing
'       Description: A short function that retrieves the path to the current workbook, may be extended.
'       Called in: Exporter, Committer, GetTag, Puller, Pusher, Tagging
'
'   UserPromptYesNo
'       Description: A short function that prompts the user to decide yes or no,
'           was easier for me to remember for some reason.
'       Called in: Committer, AnnoyUser, Importer
'
'   UserPromptText
'       Description: A function that preformats a text prompt window, similar to UserPromptYesNo, but also
'           scans the user input for unwanted characters, could maybe be extended to unwanted phrases?
'       Called in:  Committer, Tagging,
'
'   SelectFolder
'       Description: A function that opens a folder selection window
'       Called in: Importer
'
'   ModulNamenSuchen
'       Description: A function that checks whether a module of a given name exists in the current vba project
'       Called in: Importer
'
'   RemoveModule
'       Description: A function that deletes a module of a given name from the current vba project
'       Called in: Importer
'
'   BadCharacterFilter
'       Description: A function that checks whether a user input contains any undesirable characters,
'            there are different cases included in the function that can be passed arguments.
'       Called in: UserInputText
'
'   BadCharacterLoop
'       Description: A function used by BadCharacterFilter to find the bad characters
'       Called in: BadCharacterFilter
'
'   ShellCommand
'       Description: A preformatted shell command function that also takes a positive and negative result message as inputs
'       Called in: Committer, GetTag, Tagging, Pusher, Puller
'
'   GetShellOutput
'       Description: A function that also passes commands to the shell, but also fetches the shell output
'       Called in: FindTags
'
'   FindLine
'       Description: A function used to ignore the extra data added to sourcecode files when using manual/proper export methods
'       Called in: (Alt)Exporter
'
'   FindTags
'       Description: A function that retrieves all tags created in the given repository so that users may choose
'           which tag they wish to checkout
'       Called in: GitVersionCheckForm
'
'   ErrorHandler
'       Description: A function to collect the Errorhandling processes so that they do not need to be adjusted in
'           every single major macro when I think of a new better way to do things.
'       Called in: Committer, Exporter, GetTag, Importer, Puller, Pusher, SimpleWorkflows, Tagging
'
'   SelectFiles
'       Description: A function that allows a user to select which files should be used in a given process.
'       Called in: Committer, potentially useful for Importer
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Public Function GetUser() As String

 GetUser = Environ("username")

End Function

' Eine Funktion um das Workbook und die Module zu speichern.
' Benutzt in: Committer, Exporter, Importer
Public Sub Saver()

    activeWorkbook.Save

End Sub



' Eine Funktion, die den Pfad zum Git Repo angeben kann.
' Momentan gibt es einfach den Pfad zum aktiven Workbook an.
' Benutzt in: Importer, Committer, Tagger,Pusher,Puller,

Public Sub Pathing()

    Dim WorkbookPath As String

    WorkbookPath = activeWorkbook.path

    ChDir WorkbookPath

End Sub

' Simpler weg Ja/Nein Userprompt zu starten
' Benutzt in: Committer;
Public Function UserPromptYesNo(ByVal message As String) As Long
    
    UserPromptYesNo = MsgBox(message, vbYesNo)
    
End Function

' Präformatiertes Benutzereingabe Fenster, weil ich mir InputBox nicht merken konnte...
' Benutzt in: Committer; Tagger;
Public Function UserPromptText(ByVal message As String, ByVal titleText As String, ByVal fillText As String, ByVal purpose As String) As String
    
    UserPromptText = InputBox(message, titleText, fillText)
    
    If UserPromptText = vbNullString Then
        Err.Raise 1239, "Fehlender Userinput", "Es wurde kein Userinput gefunden, der Vorgang wurde abgebrochen."
    End If
        
    Do While BadCharacterFilter(UserPromptText, purpose)
        MsgBox "Ihre Eingabe hat ungewünschte Zeichen enthalten. Bitte versuchen Sie es erneut."
        UserPromptText = InputBox(message, titleText, fillText)
        If UserPromptText = vbNullString Then
            Err.Raise 1239, "Fehlender Userinput", "Es wurde kein Userinput gefunden, der Vorgang wurde abgebrochen."
        End If
    Loop

End Function

' Funktion die ein Ordner-Auswahl-Fenster öffnet
' Bentutz in: Importer;
Public Function SelectFolder() As Variant

    Dim diaFolder As FileDialog
    Dim selected As Boolean

    ' Open the file dialog
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    selected = diaFolder.Show

    If selected Then
        SelectFolder = diaFolder.SelectedItems(1)
    End If

    Set diaFolder = Nothing
End Function

' Eine Funktion die prüft ob ein Modul mit dem gegebenen Namen bereits existiert.
' Benutzt in: Importer
Public Function ModulNamenSuchen(ByVal moduleName As String) As Boolean
    
    Dim vbComponent As Object
    
    For Each vbComponent In activeWorkbook.VBProject.VBComponents
    
        If vbComponent.Name = moduleName Then
            ModulNamenSuchen = True
            Exit Function
        End If
    Next vbComponent
    
    ModulNamenSuchen = False
    
End Function

' Eine Funktion die ein Modul mit dem gegebenen Namen entfernt sofern es existiert.
' Benutzt in: Importer
Public Sub RemoveModule(ByVal Workbook As Workbook, ByVal removeName As String)
    Dim moduleName As String
    Dim vbComponent As Object
    Dim liveWorkbook As Workbook
    
    moduleName = removeName ' Replace with the name of the module you want to remove
    Set liveWorkbook = Workbook
    ' Iterate through all VBComponents in the project
    For Each vbComponent In liveWorkbook.VBProject.VBComponents
        ' Check if the current component is a module and has the specified name
        If Not vbComponent.Type = 100 And vbComponent.Name = moduleName Then
            ' Remove the module
            liveWorkbook.VBProject.VBComponents.Remove vbComponent
            'MsgBox moduleName & " removed from the VBA project.", vbInformation
            Exit Sub
        End If
    Next vbComponent
    
    ' Module not found
    MsgBox moduleName & " wurde in diesem VBA-Projekt nicht gefunden.", vbExclamation
End Sub

' Eine Funktion die überprüft ob ein InputString unerwünschte Zeichen beinhaltet
' Benutzt in: Committer, Tagger,
Public Function BadCharacterFilter(ByVal inputString As String, ByVal purpose As String) As Boolean

    Dim validCharacters As String
    validCharacters = "1234567890abcdefghijklmnopqrstuvwxyzäöü"
    
    BadCharacterFilter = True
    
    Select Case purpose
    '----------------------------------------------------------------------
        Case "Tag", "Commit"
            validCharacters = validCharacters & ",;: ._-" & vbNullString
        '---------------------------------------------------------------------
        Case "Module"
            validCharacters = validCharacters & "_" & vbNullString
        '---------------------------------------------------------------------
        Case "Version", "Filename"
            validCharacters = validCharacters & "._" & vbNullString
        '---------------------------------------------------------------------
        
        Case Else
            Exit Function
    End Select
    
    BadCharacterFilter = BadCharacterLoop(validCharacters, LCase(inputString))
    
    If Len(inputString) > 300 Then
        Err.Raise 1241, "Userinput", "Die Benutzereingabe war zu lang, der Vorgang wurde abgebrochen."
    End If
    
End Function

Private Function BadCharacterLoop(ByVal validCharacters As String, ByVal inputString As String) As Boolean
    
    Dim i As Long
    
    For i = 1 To Len(inputString)
        If Not InStr(validCharacters, Mid(inputString, i, 1)) > 0 Then
            ' If an invalid character is found, return True
            BadCharacterLoop = True
            Exit Function
        End If
    Next i
    BadCharacterLoop = False
    
End Function

' Eine Funktion, die dafür sorgt das Shell commands ausgeführt werden
' und überprüft wird ob sie erfolgreich waren oder nicht
Public Sub ShellCommand(ByVal command As String, ByVal successMessage As String, ByVal failureMessage As String, Optional ByVal purpose As String)
    
    Dim shell As Object
    Dim errorCode As Long
    Dim ErrNumber As Long

    Set shell = CreateObject("WScript.Shell")
    
    errorCode = shell.Run(command, 0, True)
    
    If errorCode = 0 Then
    
        MsgBox successMessage

    Else
        Select Case purpose
            Case "Tag", "Version"
                ErrNumber = 1234
            Case "Commit"
                ErrNumber = 1235
            Case "Push"
                ErrNumber = 1236
            Case "Pull"
                ErrNumber = 1237
            Case "TagFileRetrieval"
                ErrNumber = 1238
            Case "TagFullRetrieval"
                ErrNumber = 1239
            Case Else
                ErrNumber = 0
        End Select
        Err.Raise ErrNumber, purpose, failureMessage
    End If
        
    Set shell = Nothing

End Sub


' Den Output der ShellCommands einlesen
Public Function GetShellOutput(ByVal command As String) As String

    Dim shell As Object
    Dim exec As Object
    Dim output As String
    
    'Shellinstanz erstellen
    Set shell = CreateObject("WScript.Shell")

    ' Command ausführen und Output schnappen
    Set exec = shell.exec(command)
    output = exec.StdOut.ReadAll

    ' Return the output
    GetShellOutput = output

End Function

Public Function FindLine(ByVal content As String, ByVal term As String) As Long
    ' A function that should help with finding the "proper" start to the code often either Option Explicit or a comment, _
    to avoid the overhead lines created when exporting with the inbuild export method.
    FindLine = -1
    If term = vbNullString Or content = vbNullString Then
        MsgBox "Invalid input for FindString"
        Exit Function
    Else
        Dim lines As Variant
        Dim i As Long
        
        lines = Split(content, vbCrLf)
        For i = LBound(lines) To UBound(lines)
            If Left(lines(i), Len(term)) = term Or Left(lines(i), 1) = "'" Then
                FindLine = i
                Exit For
            End If
        Next i
    End If
End Function

' A function that retrieves the tags that exist in the current repository.
Public Function FindTags() As Variant

    Dim existingTagsRaw As String
    Dim existingTags() As String
    
    Pathing
    
    existingTagsRaw = GetShellOutput("git tag")
    
    existingTags = Split(existingTagsRaw, vbLf)
    
    ReDim Preserve existingTags(UBound(existingTags) - 1)
    
    If UBound(existingTags) > 9 Then
    
        Dim tempArray(9) As String
        Dim i As Integer
        For i = 1 To 10
            tempArray(i - 1) = existingTags(UBound(existingTags) - 9 + i)
        Next i
        existingTags = tempArray
    End If
    
    FindTags = existingTags
    
End Function

Public Sub ErrorHandler(ByVal ErrNumber As Long, ByVal ErrSource As String, ByVal ErrDescription As String)

    If ErrNumber = 1239 Then
        MsgBox Err.Description, vbOKOnly, ErrSource
    Else
        MsgBox "Im " & ErrSource & " Vorgang ist ein Fehler aufgetreten." & vbCrLf & ErrDescription, vbOKOnly, "Fehlermeldung"
    End If

End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'                       Work in Progress
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function SelectFiles(ByVal purpose As String) As Variant

    Dim selectedFiles As Variant
    
    Select Case purpose
        Case "Import"
            selectedFiles = Application.GetOpenFilename("Visual Basic Files (*.bas; *.txt; *.frm, *.cls),*.frm, *.cls *.bas;*.txt", , "Import Dateien", , True)
        Case "Commit"
            selectedFiles = Application.GetOpenFilename(, , "Wählen Sie die Dateien aus die Sie geändert haben", , True)
    End Select
    
    If VarType(selectedFiles) = vbBoolean Then
        If selectedFiles = False Then
            Err.Raise 1240, "Dateiauswahl", "Es wurden keine Dateien für den " & purpose & " Prozess ausgewählt"
        End If
    End If
    
    SelectFiles = selectedFiles
    
End Function
