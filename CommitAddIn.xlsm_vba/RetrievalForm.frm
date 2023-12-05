VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RetrievalForm 
   Caption         =   "Tag Laden"
   ClientHeight    =   4670
   ClientLeft      =   -230
   ClientTop       =   -900
   ClientWidth     =   11940
   OleObjectBlob   =   "RetrievalForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RetrievalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
' A Userform used in the 'Version Laden' button that allows users to select whether they
' would like to load a version of the entire repository or just a single file is sufficient.
'
' Given that this userform is not dynamic in nature the size is also static.
'
'
'
'
'
'
'
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
'Global Variable!!
Public retrievalType As String

Private Sub UserForm_Initialize()
    
    Me.Label1.Width = 250
    Me.Label1.Left = 15
    
    Me.Width = Me.Label1.Width + 30
    Me.Height = Me.Label1.Height + Me.CommandButton1.Height + 50
    
    Me.CommandButton1.Top = Me.Label1.Top + Me.Label1.Height + 10
    Me.CommandButton2.Top = Me.CommandButton1.Top
    Me.CommandButton1.Left = 15
    Me.CommandButton1.Width = Me.Label1.Width / 2 - 5
    Me.CommandButton2.Width = Me.CommandButton1.Width
    Me.CommandButton2.Left = Me.CommandButton1.Left + Me.Label1.Width - Me.CommandButton2.Width
    

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
