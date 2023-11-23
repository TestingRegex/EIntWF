Option Explicit

Private Sub Workbook_Open()



If ThisWorkbook.Name = "Vorlage_Protokoll_WVP.xlsm" Then
    Prozesswahl.Show
    
    
End If


End Sub

Sub Datum()



Dim d_Datum As Date, n_KW As Integer, s_MM As String, n_MM As Integer, s_Datum As String

'd_Datum = Date
'
'If Weekday(d_Datum) <> 2 Then
'    Do While Weekday(d_Datum) <> 2
'    d_Datum = d_Datum + 1
'
'    Loop
'n_KW = WorksheetFunction.WeekNum(d_Datum)
'
'
'
'End If

d_Datum = Date
s_Datum = d_Datum
s_MM = Mid(s_Datum, 4, 2)
n_MM = s_MM
s_YY = Right(s_Datum, 4)

    If n_MM < 10 Then
        s_MM = "0" & M
    Else
        n_MM = s_MM
    End If

s_Datum = "01." & s_MM & "." & s_YY

MsgBox s_Datum

End Sub

Sub test()



ThisWorkbook.Worksheets("NSA Ergebnisse").btn_WVP_NSA.Visible = True
ThisWorkbook.Worksheets("NSA Ergebnisse").btn_MVP_NSA.Visible = True

End Sub

Sub test2()
ThisWorkbook.Worksheets("NSA Ergebnisse").btn_WVP_NSA.Visible = True
ThisWorkbook.Worksheets("NSA Ergebnisse").btn_MVP_NSA.Visible = False
End Sub