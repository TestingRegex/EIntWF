'''
'   Eine Sammlung von vielleicht nützlichen Funktionen die in verschiedenen
'   Makros wieder verwendet werden
'
'''

Option Explicit

Function GetUser()

 GetUser = Environ("username")

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

' Simpler weg Ja/Nein Userprompt zu starten
' Benutzt in: Committer;
Function UserPromptYesNo(ByVal message As String)
    
    UserPromptYesNo = MsgBox(message, vbYesNo)
    
End Function

' Präformatiertes Benutzereingabe Fenster, weil ich mir InputBox nicht merken konnte...
' Benutzt in: Committer; Tagger;
Function UserInputText(ByVal message As String, ByVal titleText As String, ByVal fillText As String)

    UserInputText = InputBox(message, titleText, fillText)

End Function

' Eine Funktion die prüft ob ein Modul mit dem gegebenen Namen bereits existiert.
' Benutzt in: Importer
Function ModulNamenSuchen(ByVal moduleName As String)
    
    Dim vbComponent As Object
    
    For Each vbComponent In ActiveWorkbook.VBProject.VBComponents
    
        If vbComponent.Type = 1 And vbComponent.Name = moduleName Then
            ModulNamenSuchen = True
            Exit Function
        End If
    Next vbComponent
    
    ModulNamenSuchen = False
    
End Function

' Eine Funktion die ein Modul mit dem gegebenen Namen entfernt sofern es existiert.
' Benutzt in: Importer
Function RemoveModule(ByVal removeName As String)
    Dim moduleName As String
    Dim vbComponent As Object
    moduleName = removeName ' Replace with the name of the module you want to remove
    
    ' Iterate through all VBComponents in the project
    For Each vbComponent In ThisWorkbook.VBProject.VBComponents
        ' Check if the current component is a module and has the specified name
        If vbComponent.Type = 1 And vbComponent.Name = moduleName Then
            ' Remove the module
            ThisWorkbook.VBProject.VBComponents.Remove vbComponent
            MsgBox moduleName & " removed from the VBA project.", vbInformation
            Exit Function
        End If
    Next vbComponent
    
    ' Module not found
    MsgBox moduleName & " wurde in diesem VBA-Projekt nicht gefunden.", vbExclamation
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

' Eine Funktion die überprüft ob ein InputString unerwünschte Zeichen beinhaltet
' Benutzt in: Committer, Tagger,
Function BadCharacterFilter(ByVal inputString As String, Optional ByVal Purpose As String)

    Dim invalidCharacters As String
    Dim i As Integer
    
    If Purpose = "Tag" Then
        invalidCharacters = " ~!@#$%^&*()+,{}[]|\;:'""<>/?="
        For i = 1 To Len(inputString)
            If InStr(invalidCharacters, Mid(inputString, i, 1)) > 0 Then
                ' If an invalid character is found, return True
                BadCharacterFilter = True
                Exit Function
            End If
        Next i
    ElseIf Purpose = "Commit" Then
        invalidCharacters = " ""#$^:;'<>[]{}@"
        For i = 1 To Len(inputString)
            If InStr(invalidCharacters, Mid(inputString, i, 1)) > 0 Then
                ' If an invalid character is found, return True
                BadCharacterFilter = True
                Exit Function
            End If
        Next i
    Else
        BadCharacterFilter = False
    End If
    
End Function
