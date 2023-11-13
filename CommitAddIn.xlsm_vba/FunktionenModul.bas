'''
'   Eine Sammlung von vielleicht n�tzlichen Funktionen die in verschiedenen
'   Makros wieder verwendet werden
'
'''

Option Explicit

Function GetUser()

 GetUser = Environ("username")

End Function

' Funktion die ein Ordner-Auswahl-Fenster �ffnet
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

' Pr�formatiertes Benutzereingabe Fenster, weil ich mir InputBox nicht merken konnte...
Function UserInputText(ByVal message As String, ByVal titleText As String, ByVal fillText As String)

    UserInputText = InputBox(message, titleText, fillText)

End Function

' Eine Funktion die pr�ft ob ein Modul mit dem gegebenen Namen bereits existiert.
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