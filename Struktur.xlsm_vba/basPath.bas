' AUTHOR: Christian Plattner

' TABLE OF CONTENTS
    ' exe_choose
    ' exe_create
    ' exe_delete
    ' exe_delete_if_empty
    ' exe_open
    ' info_counted_subfolders
    ' info_exists
    
Option Explicit

Function exe_choose(Optional sP As String = "C:") As String
' LAST CHANGE: 23/05/2016

' ALGORITHM
    If Len(sP) > 2 Then
        ChDrive Left(sP, 3)
    ElseIf Len(sP) = 2 Then
        ChDrive sP & "\"
    Else
        ChDrive "C:\"
    End If
    ChDir sP & "\"
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Show
        On Error Resume Next
        exe_choose = .SelectedItems(1)
        Err.Clear
        On Error GoTo 0
    End With
End Function

Function exe_create(sP As String, Optional ByVal idepth As Integer = 1) As Boolean
' LAST CHANGE: 05/03/2016

' VARIABLES
    Dim i As Integer
    Dim iordner_tot As Integer
    Dim sP_neu As String
    
' ALGORITHM
    On Error GoTo Error
    iordner_tot = info_counted_subfolders(sP)
    If iordner_tot < idepth Then idepth = iordner_tot
    If idepth = 0 Then Exit Function
    While idepth > 0
        i = idepth
        sP_neu = sP
        While i > 0
            i = i - 1
            If Not i = 0 Then sP_neu = info_P(sP_neu)
        Wend
        Call exe_create2(sP_neu)
        idepth = idepth - 1
    Wend
    exe_create = True
Error:
End Function

Private Sub exe_create2(sP As String)
' LAST CHANGE: 05/03/2016

' ALGORITHM
    On Error Resume Next
    MkDir sP
End Sub

Function exe_delete(sP As String) As Boolean
' LAST CHANGE: 05/03/2016

' VARIABLES
    Dim objFSO As Object
    
' ALGORITHM
    On Error GoTo Error
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.DeleteFolder sP
    exe_delete = True
Error:
    Set objFSO = Nothing
End Function

Function exe_delete_if_empty(sP As String) As Boolean
' LAST CHANGE: 05/03/2016

' ALGORITHM
    On Error GoTo Error
    exe_delete_if_empty = info_exists(sP)
    If exe_delete_if_empty Then RmDir sP
Exit Function
Error:
    exe_delete_if_empty = False
End Function

Function exe_open(sP As String) As Boolean
' LAST CHANGE: 05/03/2016

' ALGORITHM
    On Error GoTo Error
    Call Shell("Explorer.exe " & sP, vbNormalFocus)
    exe_open = True
Error:
End Function

Function info_counted_subfolders(sP As String) As Integer
' LAST CHANGE: 05/03/2016

' ALGORITHM
    info_counted_subfolders = Len(sP) - Len(Replace(sP, "\", ""))
End Function

Function info_exists(sP As String) As Boolean
' LAST CHANGE: 05/03/2016

' ALGORITHM
    If Not Dir(sP, vbDirectory) = vbNullString Then info_exists = True
End Function

