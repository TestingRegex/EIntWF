' AUTHOR: Christian Plattner

' TABLE OF CONTENTS
    ' exe_choose
    ' exe_copy
    ' exe_delete
    ' exe_link
    ' info_changed
    ' info_exists
    ' info_F
    ' info_generated
    ' info_newest
    ' info_P
    ' info_size

Option Explicit

Function exe_choose(ByRef sP As String, ByRef sF As String, Optional styp As String = "Alle-Dateien (*.*),*.*,", Optional stitel As String = "Wähle eine Datei") As Boolean
' LAST CHANGE: 23/11/2016

' EXAMPLE INCLUDING SPECIAL CASES
    ' Files on "M:\test": "test1.txt", "test2.txt"
    ' Executing: exe_choose("M:",sF,stitel:="Wähle eine Testdatei")
        ' sF is here an empty variable (can also have a value)
        ' User: Has to choose a file starting on folder "M:"
        ' User: Chooses i.e. "test1.txt" in folder "M:\test"
        ' Output: sP = "M:\test", sF = "test1.txt", exe_choose = True
    ' Executing: exe_choose("M:",sF)
        ' User: Has to choose a file starting on folder "M:"
        ' User: Chooses not to pick a file
        ' Output: exe_choose = False

' VARIABLES
    Dim v As Variant
    Dim i As Integer
    Dim i_tot As Integer

' ALGORITHM
    If stitel = "Wähle eine Datei" And Not sF = vbNullString Then _
            stitel = stitel & "; Vorschlag: " & sF
    
    If info_exists(sP) Then
        ChDrive Left(sP, 3)
        ChDir sP
    End If
    
    v = Application.GetOpenFilename(styp, title:=stitel, MultiSelect:=False)
    
    If v = False Then
        Exit Function
    Else
        exe_choose = True
    End If
    
    sF = CStr(v)
    i_tot = Len(sF)
    sP = info_P(sF)
    sF = info_F(sF)
    Set v = Nothing
End Function

Function exe_copy(ByVal sPnF_old As String, sPnF_new As String, Optional boverwrite As Boolean = False) As Boolean
' LAST CHANGE: 23/11/2016

' EXAMPLE INCLUDING SPECIAL CASES
    ' Files on "M:": "test10.txt", "teest1.txt"
    ' Folders on "M:": "test" (empty)
    ' Executing: exe_copy("M:\t?st*.txt","M:\test\", False)
        ' sPnF_old = "M:\t?st*.txt" --> sPnF_old = "M:\test10.txt"
        ' sPnF_new = "M:\test\" --> sPnF_new = "M:\test\test10.txt"
        ' sPnF_new does not exist --> Proceed
        ' Copies "test10.txt" from "M:" to "M:\test" named "test10.txt"
        ' exe_copy = True
    ' Executing: exe_copy("M:\test10.txt","M:\test\test10.txt", True)
        ' Copies "test10.txt" from "M:" to "M:\test" named "test10.txt" and overwrites the existing file
        ' exe_copy = True
    ' Executing: exe_copy("M:\test10.txt","M:\test\test10.txt", False)
        ' sPnF_new = "M:\test\test10.txt" does exist --> Ask user
        ' "Die Datei 'M:\test10.txt' existiert bereits. Soll diese überschrieben werden?" --> User: "No"
        ' exe_copy = False
    ' Executing: exe_copy("M:\test10.txt","M:\test\test\test10.txt", False)
        ' Error: Folder does not exist
        ' exe_copy = False

' VARIABLES
    Dim b As Byte
    Dim sF As String
    Dim sP As String
    Dim bDA As Boolean
    
' ALORITHMUS
    On Error GoTo Error
    
    ' sPnF_old: Completing undefined parts
    If InStr(sPnF_old, "*") > 0 Or InStr(sPnF_old, "?") > 0 Then
        sF = info_newest(sPnF_old)
        If sF = vbNullString Then GoTo Error
        sP = info_P(sPnF_old)
        sPnF_old = sP & "\" & sF
    End If
    
    ' sPnF_new: If sF_new is not defined then use sF_new = sF_old
    If Right(sPnF_new, 1) = "\" Then
        sF = info_F(sPnF_old)
        sPnF_new = sPnF_new & sF
    End If
    
    ' Overwrite: Yes/No
    If info_exists(sPnF_new) And Not boverwrite Then
        b = MsgBox("Die Datei '" & info_F(sPnF_new) & "' existiert bereits. Soll diese überschrieben werden?", vbYesNo)
        If b = vbNo Then Exit Function
    End If
    
    ' DisplayAlerts before copy = DisplayAlerts after copy
    bDA = Application.DisplayAlerts
    Application.DisplayAlerts = False
    FileCopy sPnF_old, sPnF_new
    Application.DisplayAlerts = bDA
    
    exe_copy = True
Error:
End Function

Function exe_delete(sPnF As String) As Boolean
' LAST CHANGE: 23/11/2016

' ALGORITHM
    On Error GoTo Error
    
    exe_delete = info_exists(sPnF)
    If exe_delete Then _
            Kill sPnF
    
    Exit Function
Error:
    exe_delete = False
End Function

Function exe_link(sPnF As String, sP_lnk As String, sF_lnk As String) As Boolean
' LAST CHANGE: 23/11/2016

' VARIABLES
    Dim olink As Object
    
' ALGORITHM
    On Error GoTo Error
    
    Set olink = CreateObject("WScript.Shell")
    
    With olink.CreateShortcut(sP_lnk & "\" & sF_lnk & ".lnk")
        .TargetPath = sPnF
        .Arguments = ""
        .Save
    End With
    
    exe_link = True
Error:
    Set olink = Nothing
End Function

Function info_changed(sPnF As String) As Date
' LAST CHANGE: 23/11/2016

' VARIABLES
    Dim oFSO As Object
    Dim oFile As Object
    
' ALGORITHM
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    If oFSO.FileExists(sPnF) Then
        Set oFile = oFSO.GetFile(sPnF)
        info_changed = oFile.DateLastModified
    End If
    
    Set oFSO = Nothing
    Set oFile = Nothing
End Function

Function info_exists(sPnF As String) As Boolean
' LAST CHANGE: 23/11/2016

' ALGORITHM
    If Not Dir(sPnF, vbDirectory) = vbNullString Then _
            info_exists = True
End Function

Function info_F(sPnF As String) As String
' LAST CHANGE: 23/11/2016

' ALGORITHM
    If InStr(1, sPnF, "\") = 0 Then
        info_F = sPnF
    Else
        info_F = Right(sPnF, Len(sPnF) - Len(info_P(sPnF)) - 1)
    End If
End Function

Function info_generated(sPnF As String) As Date
' LAST CHANGE: 23/11/2016

' VARIABLES
    Dim oFSO As Object
    Dim oFile As Object
    
' ALGORITHM
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    If oFSO.FileExists(sPnF) Then
        Set oFile = oFSO.GetFile(sPnF)
        info_generated = oFile.DateCreated
    End If
    
    Set oFSO = Nothing
    Set oFile = Nothing
End Function

Function info_newest(sPnF As String, Optional ddate As Date) As String
' LAST CHANGE: 23/11/2016

' VARIABLES
    Dim sF As String
    Dim sP As String
    Dim d1 As Date
    Dim d2 As Date
    
' ALGORITHM
    sP = info_P(sPnF)
    
    If Not info_exists(sP) Then _
            Exit Function
    
    sF = Dir(sPnF)
    If ddate = 0 Then
        ' ALGORITHM without given date
        While sF <> ""
            d2 = info_changed(sP & "\" & sF)
            If d2 > d1 Then
                d1 = d2
                info_newest = sF
            End If
            sF = Dir
        Wend
    Else
        ' ALGORITHM with given date
        While sF <> ""
            d2 = info_changed(sP & "\" & sF)
            If d2 > d1 And Int(d2) = Int(ddate) Then
                d1 = d2
                info_newest = sF
            End If
            sF = Dir
        Wend
    End If
End Function

Function info_P(sPnF As String) As String
' LAST CHANGE: 23/11/2016

' ALGORITHM
    If InStr(1, sPnF, "\") = 0 Then Exit Function
    
    info_P = sPnF
    Do
        If Right(info_P, 1) = "\" Then Exit Do
        info_P = Left(info_P, Len(info_P) - 1)
    Loop
    
    info_P = Left(info_P, Len(info_P) - 1)
End Function

Function info_size(sPnF As String) As Long
' LAST CHANGE: 23/11/2016

' VARIABLES
    Dim oFSO As Object
    Dim oF As Object
    
' ALGORITHM
    On Error GoTo Error
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oF = oFSO.GetFile(sPnF)
    
    info_size = oF.Size
Error:
    Set oFSO = Nothing
    Set oF = Nothing
End Function