' AUTHOR: Christian Plattner

Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
' LAST CHANGE: 26.09.2007
' VARIABLES
    Dim oWB As Workbook
    Dim i As Integer

' ALGORITHM
    For Each oWB In Workbooks
        i = i + 1
        If i > 1 Then Exit For
    Next oWB
    
    If i = 1 Then
        Application.DisplayAlerts = False
        Application.Quit
    Else
        ThisWorkbook.Close Savechanges:=False
        Application.DisplayAlerts = True                    'Fragefenster wieder einschalten
        ActiveWorkbook.UpdateLinks = xlUpdateLinksAlways    'Akutalisiere Verknuepfungen
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
' LAST CHANGE: 29/04/2016
' ALGORITHM
    With tab_GUI
        .Range("nUser").Cells.ClearContents
        .Range("nDate").Cells.ClearContents
        .Range("nProcess").Value = "WVP"
'        .Range("nDatei_Transfer").Cells.ClearContents
        .Activate
        .Range("nProcess").Select
        .Range("nBoolean_DUMMY").FormulaR1C1 = "=RC[1]"
        .box_log.Value = False
        Call .Worksheet_Activate
    End With
    Call tabSTRUCTURE.Worksheet_Activate
End Sub

Sub Workbook_Open()
' LAST CHANGE: 15/04/2016
' VARIABLES
    Dim onetwork As Object
    Dim sPnF As String
    Dim sprocess As String
    
' ALGORITHM
    Set onetwork = CreateObject("wscript.network")
    With tab_GUI
        .Range("nUser").Value = onetwork.UserName
        .Range("nDate").Value = DateSerial(Year(Now - 5 / 24), Month(Now - 5 / 24), Day(Now - 5 / 24))
        .Range("nOperativ").Value = ThisWorkbook.Path
        sPnF = .Range("nConfig").Value
        If basFile.info_exists(sPnF) Then
            sprocess = basTXT.info_as_string(sPnF)
            sprocess = Left(sprocess, InStr(1, sprocess, Chr(80)))
            .Range("nProcess").Value = sprocess
        End If
        
        sprocess = "WVP"
        .Range("nProcess").Value = sprocess
        
        'Call .exe_reset_zonentransfer
        .Activate
    End With
End Sub