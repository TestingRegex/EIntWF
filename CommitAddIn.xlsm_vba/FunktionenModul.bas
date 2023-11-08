'''
'   Eine Sammlung von vielleicht nützlichen Funktionene die in verschiedenen
'   Makros wieder verwendet wird
'
'''

Option Explicit

Function GetUser()

 GetUser = Environ("username")

End Function

' Funktion die ein Ordner-Auswahl-Fenster öffnet
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
Function UserPromptYesNo(ByVal message As String)
    
    UserPromptYesNo = MsgBox(message, vbYesNo)
    
End Function

' Präformatiertes Benutzereingabe Fenster.
Function UserInputText(ByVal message As String, ByVal titleText As String, ByVal fillText As String)

    UserInputText = InputBox(message, titleText, fillText)

End Function


