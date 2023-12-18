VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RetrievalForm 
   Caption         =   "Tag Laden"
   ClientHeight    =   590
   ClientLeft      =   -350
   ClientTop       =   -1350
   ClientWidth     =   1380
   OleObjectBlob   =   "RetrievalForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RetrievalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule SuspiciousPredeclaredInstanceAccess
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

    Me.LabelRetrievalForm.Width = 250
    Me.LabelRetrievalForm.Left = 15
    
    Me.Width = Me.LabelRetrievalForm.Width + 30
    Me.Height = Me.LabelRetrievalForm.Height + Me.SingleFile.Height + 50
    
    Me.SingleFile.Top = Me.LabelRetrievalForm.Top + Me.LabelRetrievalForm.Height + 10
    Me.CompleteRepository.Top = Me.SingleFile.Top
    Me.SingleFile.Left = 15
    Me.SingleFile.Width = Me.LabelRetrievalForm.Width / 2 - 5
    Me.CompleteRepository.Width = Me.SingleFile.Width
    Me.CompleteRepository.Left = Me.SingleFile.Left + Me.LabelRetrievalForm.Width - Me.CompleteRepository.Width
    

End Sub

Private Sub SingleFile_Click()
    
    ' Retrieve a single file
    RetrievalForm.retrievalType = SingleFile.Caption
    
    ' Close the UserForm
    Unload Me

    GitVersionCheckForm.Show

End Sub

Private Sub CompleteRepository_Click()

    ' Retrieve a single file
    RetrievalForm.retrievalType = CompleteRepository.Caption
    'Debug.Print "In RetrievalForm value of retrievalType: " & RetrievalForm.retrievalType
    
    ' Close the UserForm
    Unload Me

    GitVersionCheckForm.Show

End Sub
