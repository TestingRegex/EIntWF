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
'   AnnoyUsers
'       Description: A placeholder function that is supposed to act as a reminder for users to adhere to coding best practices,
'           might not make it to final production
'       Called in: Committer, Exporter, Tagging, SimpleWorkflows
'
'   FindTags
'       Description: A function that retrieves all tags created in the given repository so that users may choose
'           which tag they wish to checkout
'       Called in: GitVersionCheckForm
'
'
'
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Function GetUser()

 GetUser = Environ("username")

End Function

' Eine Funktion um das Workbook und die Module zu speichern.
' Benutzt in: Committer, Exporter, Importer
Function Saver()

    ActiveWorkbook.Save

End Function



' Eine Funktion, die den Pfad zum Git Repo angeben kann.
' Momentan gibt es einfach den Pfad zum aktiven Workbook an.
' Benutzt in: Importer, Committer, Tagger,Pusher,Puller,

Function Pathing()

    Dim WorkbookPath As String

    WorkbookPath = ActiveWorkbook.path

    ChDir WorkbookPath

End Function

' Simpler weg Ja/Nein Userprompt zu starten
' Benutzt in: Committer;
Function UserPromptYesNo(ByVal message As String)
    
    UserPromptYesNo = MsgBox(message, vbYesNo)
    
End Function

' Präformatiertes Benutzereingabe Fenster, weil ich mir InputBox nicht merken konnte...
' Benutzt in: Committer; Tagger;
Function UserPromptText(ByVal message As String, ByVal titleText As String, ByVal fillText As String, ByVal purpose As String)
    
    UserPromptText = InputBox(message, titleText, fillText)
    
    If UserPromptText = "" Then
        Exit Function
    End If
        
    Do While BadCharacterFilter(UserPromptText, purpose)
        MsgBox "Ihre Eingabe hat ungewünschte Zeichen enthalten. Bitte versuchen Sie es erneut."
        UserPromptText = InputBox(message, titleText, fillText)
        If UserPromptText = "" Then
            Exit Function
        End If
    Loop

End Function

' Funktion die ein Ordner-Auswahl-Fenster öffnet
' Bentutz in: Importer;
Function SelectFolder()
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
Function ModulNamenSuchen(ByVal moduleName As String)
    
    Dim vbComponent As Object
    
    For Each vbComponent In ActiveWorkbook.VBProject.VBComponents
    
        If vbComponent.Name = moduleName Then
            ModulNamenSuchen = True
            Exit Function
        End If
    Next vbComponent
    
    ModulNamenSuchen = False
    
End Function

' Eine Funktion die ein Modul mit dem gegebenen Namen entfernt sofern es existiert.
' Benutzt in: Importer
Function RemoveModule(ByVal Workbook As Workbook, ByVal removeName As String)
    Dim moduleName As String
    Dim vbComponent As Object
    Dim wb As Workbook
    moduleName = removeName ' Replace with the name of the module you want to remove
    Set wb = Workbook
    ' Iterate through all VBComponents in the project
    For Each vbComponent In wb.VBProject.VBComponents
        ' Check if the current component is a module and has the specified name
        If Not vbComponent.Type = 100 And vbComponent.Name = moduleName Then
            ' Remove the module
            wb.VBProject.VBComponents.Remove vbComponent
            'MsgBox moduleName & " removed from the VBA project.", vbInformation
            Exit Function
        End If
    Next vbComponent
    
    ' Module not found
    MsgBox moduleName & " wurde in diesem VBA-Projekt nicht gefunden.", vbExclamation
End Function

' Eine Funktion die überprüft ob ein InputString unerwünschte Zeichen beinhaltet
' Benutzt in: Committer, Tagger,
Function BadCharacterFilter(ByVal inputString As String, ByVal purpose As String)

    Dim invalidCharacters As String
    
    BadCharacterFilter = False
    
    Select Case purpose
    '----------------------------------------------------------------------
        Case "Tag"
            invalidCharacters = " ~!@#$%^&*()+,{}[]|\;:'""<>/|?="
        '---------------------------------------------------------------------
        Case "Commit"
            invalidCharacters = """#$^:;'<>[]{}@|/\²³="
        '---------------------------------------------------------------------
        Case "File", "Module"
            ' Add Conditions we don't want in Filenames.
            invalidCharacters = """/\*~$!&=?!{[]}.;:'><,^@€²³| "
        '---------------------------------------------------------------------
        Case "Version"
            invalidCharacters = ""
        '---------------------------------------------------------------------
        Case Else
            Exit Function
    End Select
    
    BadCharacterFilter = BadCharacterLoop(invalidCharacters, inputString)
    
End Function

Function BadCharacterLoop(ByVal invalidCharacters As String, ByVal inputString As String)
    
    Dim i As Integer
    
    For i = 1 To Len(inputString)
        If InStr(invalidCharacters, Mid(inputString, i, 1)) > 0 Then
            ' If an invalid character is found, return True
            BadCharacterLoop = True
            Exit Function
        End If
    Next i

End Function

' Eine Funktion, die dafür sorgt das Shell commands ausgeführt werden
' und überprüft wird ob sie erfolgreich waren oder nicht
Function ShellCommand(command As String, successMessage As String, failureMessage As String, Optional ByVal purpose As String)
    
    Dim shell As Object
    Dim errorCode As Integer
    Dim ErrNumber As Long

    Set shell = CreateObject("WScript.Shell")
    
    errorCode = shell.Run(command, 0, True)
    
    If errorCode = 0 Then
    
        MsgBox successMessage

    Else
        Select Case purpose
            Case "Tag"
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
        Err.Raise 1234, purpose, failureMessage
    End If
    
    ShellCommand = errorCode
    
    Set shell = Nothing

End Function


' Den Output der ShellCommands einlesen
Function GetShellOutput(ByVal command As String)

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

Function FindLine(ByVal content As String, ByVal term As String)
    ' A function that should help with finding the "proper" start to the code often either Option Explicit or a comment, _
    to avoid the overhead lines created when exporting with the inbuild export method.
    FindLine = -1
    If term = "" Or content = "" Then
        MsgBox "Invalid input for FindString"
        Exit Function
    Else
        Dim lines As Variant
        Dim i As Integer
        
        lines = Split(content, vbCrLf)
        For i = LBound(lines) To UBound(lines)
            If Left(lines(i), Len(term)) = term Or Left(lines(i), 1) = "'" Then
                FindLine = i
                Exit For
            End If
        Next i
    End If
End Function

Function AnnoyUsers()

    AnnoyUsers = UserPromptYesNo("Have you cleaned up your code and spreadsheets?")
    
End Function

' A function that retrieves the tags that exist in the current repository.
Function FindTags()

    Dim existingTagsRaw As String
    Dim existingTags() As String
    Dim i As Integer
    
    Pathing
    
    existingTagsRaw = GetShellOutput("git tag")
    
    existingTags = Split(existingTagsRaw, vbLf)
    
    ReDim Preserve existingTags(UBound(existingTags) - 1)
    
    FindTags = existingTags
    
End Function
