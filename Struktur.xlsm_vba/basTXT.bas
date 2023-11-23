' AUTHOR Christian Plattner

' TABLE OF CONTENTS
    ' exe_generate
    ' info_as_string
    ' info_as_array
    
Option Explicit

Function exe_generate(sPnF As String, Optional boverwrite As Boolean = False, Optional stext As String = vbNullString) As Boolean
' LAST CHANGE: 04/03/2016
' ALGORITHM
    On Error GoTo Error
    If (Not boverwrite) And (Not Dir(sPnF, vbDirectory) = vbNullString) Then _
            GoTo Error
    Open sPnF For Output As #1
        Print #1, stext
    Close #1
    exe_generate = True
    Exit Function
Error:
    Close #1
End Function

Function info_as_array(sPnF As String, Optional sdelimiter As String = vbNullString) As Variant
' LAST CHANGE: 05/03/2016

' VARIABLES
    Dim stext As String
    Dim stext_after_wrap As String
    Dim z As Integer
    Dim z_max As String
    Dim s1 As Integer
    Dim s2 As Integer
    Dim k As Integer: Dim k_10 As Integer: Dim k_13 As Integer
    Dim srange1() As String
    Dim srange2() As String

' ALGORITHM
    On Error GoTo Error
    z = 0
    z_max = 0
    s1 = 0
    Open sPnF For Input As #1
        Do While Not EOF(1)
            Line Input #1, stext
            z_max = z_max + 1 + 2 * Len(stext) - Len(Replace(stext, Chr(10), "")) - Len(Replace(stext, Chr(13), ""))
        Loop
    Close #1
    Open sPnF For Input As #1
        z_max = z_max - 1
        ReDim srange1(z_max, 0) As String
        If sdelimiter = vbNullString Then
            Do While Not EOF(1)
                Line Input #1, stext
                k_10 = InStr(1, stext, Chr(10))
                k_13 = InStr(1, stext, Chr(13))
                While k_10 > 0 Or k_13 > 0
                    If (k_10 > 0 And k_10 < k_13) Or k_13 = 0 Then
                        stext_after_wrap = Right(stext, Len(stext) - k_10)
                        stext = Left(stext, k_10 - 1)
                    Else
                        stext_after_wrap = Right(stext, Len(stext) - k_13)
                        stext = Left(stext, k_13 - 1)
                    End If
                    srange1(z, 0) = stext
                    z = z + 1
                    stext = stext_after_wrap
                    k_10 = InStr(1, stext, Chr(10))
                    k_13 = InStr(1, stext, Chr(13))
                Wend
                srange1(z, 0) = stext
                z = z + 1
            Loop
        Else
            Do While Not EOF(1)
                Line Input #1, stext
                k_10 = InStr(1, stext, Chr(10))
                k_13 = InStr(1, stext, Chr(13))
                While k_10 > 0 Or k_13 > 0
                    If (k_10 > 0 And k_10 < k_13) Or k_13 = 0 Then
                        stext_after_wrap = Right(stext, Len(stext) - k_10)
                        stext = Left(stext, k_10 - 1)
                    Else
                        stext_after_wrap = Right(stext, Len(stext) - k_13)
                        stext = Left(stext, k_13 - 1)
                    End If
                    s2 = Len(stext) - Len(Replace(stext, sdelimiter, ""))
                    If s2 > s1 Then _
                            s1 = s2
                    ReDim Preserve srange1(z_max, s1) As String
                    ReDim srange2(s2) As String
                    srange2 = Split(stext, sdelimiter)
                    For k = 0 To s2
                        srange1(z, k) = srange2(k)
                    Next k
                    z = z + 1
                    stext = stext_after_wrap
                    k_10 = InStr(1, stext, Chr(10))
                    k_13 = InStr(1, stext, Chr(13))
                Wend
                s2 = Len(stext) - Len(Replace(stext, sdelimiter, ""))
                If s2 > s1 Then _
                        s1 = s2
                ReDim Preserve srange1(z_max, s1) As String
                ReDim srange2(s2) As String
                srange2 = Split(stext, sdelimiter)
                For k = 0 To s2
                    srange1(z, k) = srange2(k)
                Next k
                z = z + 1
            Loop
        End If
Error:
    Close #1
    info_as_array = srange1
End Function

Function info_as_string(sPnF As String) As String
' LAST CHANGE: 28/06/2015

' ALGORITHM
    Open sPnF For Input As #1
    info_as_string = Input$(LOF(1), 1)
    Close #1
End Function