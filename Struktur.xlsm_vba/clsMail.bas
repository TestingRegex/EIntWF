' AUTHOR: Christian Plattner
' https://msdn.microsoft.com/de-de/library/microsoft.office.interop.outlook.mailitem_members.aspx
    
Option Explicit

' GLOBAL VARIABLES
    Dim mSubject As String
    Dim mBody As String
    Dim mTo As String
    Dim mCC As String
    Dim mBCC As String
    Dim mAttachment As String

' Subject
Public Property Let nSubject(ByVal sSubject As String): mSubject = sSubject: End Property
Public Property Get nSubject() As String: nSubject = mSubject: End Property

' Body
Public Property Let nBody(ByVal sBody As String): mBody = sBody: End Property
Public Property Get nBody() As String: nBody = mBody: End Property

' To
Public Property Let nTo(ByVal sTo As String): mTo = sTo: End Property
Public Property Get nTo() As String: nTo = mTo: End Property

' CC
Public Property Let nCC(ByVal sCC As String): mCC = sCC: End Property
Public Property Get nCC() As String: nCC = mCC: End Property

' BCC
Public Property Let nBCC(ByVal sBCC As String): mBCC = sBCC: End Property
Public Property Get nBCC() As String: nBCC = mBCC: End Property

' Anhang
Public Property Let nAttachment(ByVal sAttachment As String): mAttachment = sAttachment: End Property
Public Property Get nAttachment() As String: nAttachment = mAttachment: End Property

Private Sub Class_Initialize()
' LAST CHANGE: 04/02/2016

' ALGORITHM
    Call exe_reset
End Sub

Sub exe_prepare()
' LAST CHANGE: 04/02/2016

' VARIABLES
    Dim sPnF1 As String
    Dim sPnF2 As String
    Dim i As Integer
    Dim oOutlook As Object
    
' ALGORITHM
    Set oOutlook = CreateObject("Outlook.Application")
    With oOutlook.CreateItem(0)
        .To = mTo
        If Not mCC = vbNullString Then _
                .CC = mCC
        If Not mBCC = vbNullString Then _
                .BCC = mBCC
        
        .Subject = mSubject
        .Body = mBody
        If (Not mAttachment = vbNullString) And InStr(1, mAttachment, ".") > 0 Then
            If InStr(1, mAttachment, ";") > 0 Then
                sPnF1 = mAttachment
                i = InStr(1, sPnF1, ";")
                While i > 0
                    sPnF2 = Left(sPnF1, i - 1)
                    If Not Dir(sPnF2, vbDirectory) = vbNullString And InStr(1, sPnF2, ".") > 0 Then _
                            .Attachments.Add sPnF2
                    sPnF1 = Right(sPnF1, Len(sPnF1) - i)
                    i = InStr(1, sPnF1, ";")
                    If i = 0 And InStr(1, sPnF1, ".") > 0 Then
                        If Not Dir(sPnF1) = vbNullString Then _
                                .Attachments.Add sPnF1
                    End If
                Wend
            Else
                If Not Dir(mAttachment, vbDirectory) = vbNullString Then _
                        .Attachments.Add mAttachment
            End If
        End If
        .Display
    End With
    Set oOutlook = Nothing
End Sub

Sub exe_reset()
' LAST CHANGE: 04/02/2016

' ALGORITHM
    mSubject = vbNullString
    mBody = vbNullString
    mTo = vbNullString
    mCC = vbNullString
    mBCC = vbNullString
    mAttachment = vbNullString
End Sub

