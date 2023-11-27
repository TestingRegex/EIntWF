Option Explicit

Sub UserForm_Initialize()
    ' Initialize the UserForm
    
    ' Example array of choices
    Dim choices() As String
    choices = FindTags()
    'Debug.Print UBound(choices)

    
    Me.Width = 250
    Me.Frame1.Height = UBound(choices) * 38 + 10
    
    Me.WeiterButton.Top = Me.Frame1.Top + Me.Frame1.Height + 10
    Me.CancelButton.Top = Me.Frame1.Top + Me.Frame1.Height + 10
    
    Me.Height = Me.Frame1.Height + Me.WeiterButton.Height + 75
    
    ' Generate option buttons based on the array
    Dim i As Integer
    For i = LBound(choices) To UBound(choices)
        ' Create an option button
        Dim optButton As MSForms.OptionButton
        Set optButton = Frame1.Controls.Add("Forms.OptionButton.1", "OptionButton" & i)
        
        ' Set properties for the option button
        optButton.Caption = choices(i)
        optButton.Top = 20 + (i * 20) ' Adjust the top position based on your layout
        optButton.Left = 20 ' Adjust the left position based on your layout
    Next i
End Sub



Private Sub WeiterButton_Click()
    ' Handle the OK button click event
    Debug.Print "RetrievalForm.retrievalType: " & RetrievalForm.retrievalType
    ' Loop through the option buttons to find the selected one
    Dim i As Integer
    For i = 0 To Frame1.Controls.Count - 1
        If TypeOf Frame1.Controls(i) Is MSForms.OptionButton Then
            If Frame1.Controls(i).Value = True Then
                ' The option button is selected
                'MsgBox "Selected option: " & Frame1.Controls(i).Caption
                If RetrievalForm.retrievalType = "Individuelle Datei" Then
                   'MsgBox "Want to call TagFileRetriaval"
                    TagFileRetrieval (Frame1.Controls(i).Caption)
                    
                    ' Close the UserForm
                    Unload Me
                ElseIf RetrievalForm.retrievalType = "Gesamtes Repository" Then
                    'MsgBox "Want to call TagFullRetriaval"
                    TagFullRetrieval (Frame1.Controls(i).Caption)
                    
                    ' Close the UserForm
                    Unload Me
                Else
                    MsgBox "Etwas ist schief gegangen."
                    Unload Me
                    Exit Sub
                End If
            End If
        End If
    Next i
    
    
End Sub


Private Sub CancelButton_Click()
    ' Handle the Cancel button click
    MsgBox "Vorgang abgebrochen."
    Unload Me ' Close the UserForm
End Sub