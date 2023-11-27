Option Explicit
'Global Variable!!
Public retrievalType As String

Private Sub CommandButton1_Click()
    
    ' Retrieve a single file
    RetrievalForm.retrievalType = CommandButton1.Caption
    
    ' Close the UserForm
    Unload Me

    GitVersionCheckForm.Show

End Sub

Private Sub CommandButton2_Click()

    ' Retrieve a single file
    RetrievalForm.retrievalType = CommandButton2.Caption
    'Debug.Print "In RetrievalForm value of retrievalType: " & RetrievalForm.retrievalType
    
    ' Close the UserForm
    Unload Me

    GitVersionCheckForm.Show

End Sub