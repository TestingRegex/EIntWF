VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RetrievalForm 
   Caption         =   "Tag Laden"
   ClientHeight    =   1340
   ClientLeft      =   -230
   ClientTop       =   -900
   ClientWidth     =   1420
   OleObjectBlob   =   "RetrievalForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RetrievalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Global Variable!!
Public retrievalType As String

Private Sub UserForm_Initialize()

    Me.Width = Me.Label1.Width + 20
    Me.Height = Me.Label1.Height + Me.CommandButton1.Height + 50

End Sub

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
