VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GitVersionCheckForm 
   Caption         =   "Versionswahl"
   ClientHeight    =   3410
   ClientLeft      =   -960
   ClientTop       =   -3750
   ClientWidth     =   8160
   OleObjectBlob   =   "GitVersionCheckForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GitVersionCheckForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
' A Userform used in the 'Version Laden' button that allows users to select which tag/version
' they would like to load.
'
' The Userform's size/layout is dynamically sized in order to accomodate a varying number of
' existing tags/versions within a given repository.
'
'
'
'
'
'
'
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit

Private Sub UserForm_Initialize()
    ' Initialize the UserForm
    
    ' Example array of choices
    Dim choices() As String
    choices = FindTags()

    
    Me.Width = 250
    Me.FrameTags.Height = UBound(choices) * 38 + 10
    
    Me.WeiterButton.Top = Me.FrameTags.Top + Me.FrameTags.Height + 10
    Me.CancelButton.Top = Me.FrameTags.Top + Me.FrameTags.Height + 10
    
    Me.Height = Me.FrameTags.Height + Me.WeiterButton.Height + 75
    
    ' Generate option buttons based on the array
    Dim i As Long
    For i = LBound(choices) To UBound(choices)
        ' Create an option button
        Dim optButton As MSForms.OptionButton
        Set optButton = FrameTags.Controls.Add("Forms.OptionButton.1", "OptionButton" & i)
        
        ' Set properties for the option button
        optButton.Caption = choices(i)
        optButton.Top = 20 + (i * 20) ' Adjust the top position based on your layout
        optButton.Left = 20 ' Adjust the left position based on your layout
    Next i
End Sub



Private Sub WeiterButton_Click()
    ' Handle the OK button click event
    'Debug.Print "RetrievalForm.retrievalType: " & RetrievalForm.retrievalType
    ' Loop through the option buttons to find the selected one
    Dim i As Long
    
    For i = 0 To FrameTags.Controls.Count - 1
        If TypeOf FrameTags.Controls(i) Is MSForms.OptionButton Then
            If FrameTags.Controls(i).Value = True Then
                ' The option button is selected
                If RetrievalForm.retrievalType = "Individuelle Datei" Then
                    ' Close the UserForm
                    Unload Me
                    TagFileRetrieval (FrameTags.Controls(i).Caption)
                    
                ElseIf RetrievalForm.retrievalType = "Gesamtes Repository" Then
                    ' Close the UserForm
                    Unload Me
                    TagFullRetrieval (FrameTags.Controls(i).Caption)
                    
                Else
                    Unload Me
                    MsgBox "Die variable RetrievalForm.retrievalType hat einen unerwarteten Wert."
                    
                    Exit Sub
                End If
            End If
        End If
    Next i
    
    
End Sub


Private Sub CancelButton_Click()
    ' Handle the Cancel button click
    Unload Me ' Close the UserForm
    MsgBox "Der Vorgang wurde abgebrochen.", vbApplicationModal, "Vorgang Beendet"
    
End Sub
