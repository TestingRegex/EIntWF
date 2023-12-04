'''
'   Eine Sammlung von vielleicht nützlichen Funktionen die in verschiedenen
'   Makros wieder verwendet werden
'
'''

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
Function UserPromptText(ByVal message As String, ByVal titleText As String, ByVal fillText As String, ByVal Purpose As String)
    
    UserPromptText = InputBox(message, titleText, fillText)
    
    If UserPromptText = "" Then
        Exit Function
    End If
        
    Do While BadCharacterFilter(UserPromptText, Purpose)
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



' Eine Funktion, die den Pfad zum Git Repo angeben kann.
' Momentan gibt es einfach den Pfad zum aktiven Workbook an.
' Benutzt in: Importer, Committer, Tagger,Pusher,Puller,

Function Pathing()

    Dim WorkbookPath As String

    WorkbookPath = ActiveWorkbook.path

    ChDir WorkbookPath

End Function

' Eine Funktion die überprüft ob ein InputString unerwünschte Zeichen beinhaltet
' Benutzt in: Committer, Tagger,
Function BadCharacterFilter(ByVal inputString As String, ByVal Purpose As String)

    Dim invalidCharacters As String
    
    BadCharacterFilter = False
    
    Select Case Purpose
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
Function ShellCommand(command As String, successMessage As String, failureMessage As String)
    
    Dim shell As Object
    Dim errorCode As Integer

    Set shell = CreateObject("WScript.Shell")
    
    errorCode = shell.Run(command, 0, True)
    
    If errorCode = 0 Then
    
        MsgBox successMessage

    Else
        MsgBox failureMessage
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
