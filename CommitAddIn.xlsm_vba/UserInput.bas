' Test um mit UserInput umgehen zu können

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