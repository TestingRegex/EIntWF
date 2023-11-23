Option Explicit

Public Sub NTC_Einlesen()

Dim w_Me As Workbook, w_Bestimmung As Workbook, ws_Daten As Worksheet, ws_NTCTool As Worksheet, ws_ADF As Worksheet
Dim s_Pfad As String, s_Name As String, s_Datum As String, r_Datum As Range, z_NTC As Integer, i As Integer, z_NTCTool As Integer, n_ADF_CH As Integer, n_Version As Integer
Dim s_GF As String, nZ_GF As Integer, Iperm As String, s_Grenze As String
Set w_Me = ThisWorkbook
Set ws_Daten = w_Me.Sheets("Einstellungen")
Set ws_NTCTool = w_Me.Sheets("NTC ADF-CH und CH-FR")
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
'ws_NTC.Range("F3:AC4").ClearContents
'ws_NTCTool.Range("A2:J25").ClearContents

Iperm = ThisWorkbook.Worksheets("Einstellungen").Range("N12")


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%% NTC für W-1 einlesen %%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

ws_NTCTool.Range("B3:F400").ClearContents

s_Pfad = ws_Daten.Range("B3")


n_Version = 10

Do While n_Version > -1
    s_Name = ws_Daten.Range("C3") & n_Version & ".xlsx"
    
    If Dir(s_Pfad & s_Name) <> "" Then
        Application.EnableEvents = False
        Set w_Bestimmung = Workbooks.Open(s_Pfad & s_Name, ReadOnly:=True)
        Application.EnableEvents = True
        Exit Do
    Else
        n_Version = n_Version - 1
        If n_Version = -1 Then
            MsgBox "Die CH-FR Werte für W-1 von der WVP wurden nicht gefunden! Werte manuel in den Reiter <<NTC ADF-CH und CH-FR>> kopieren.(Pfad: O:\MA-SO-D-2_D-1\_Planung_YYYY\NTC\NTC_Tool)"
'            Exit Sub
            GoTo Weiter
        End If
        
    End If
    
Loop

'*************************
'FR Einlesen
'**************************

s_Grenze = ws_Daten.Range("D3")

Set ws_ADF = Workbooks(s_Name).Sheets(s_Grenze)

' Zeile in ADF Bestimungshilfe Blatt NTC_ADF-CH_CHFR

z_NTCTool = 171 ' Zeile im Protokoll WVP
z_NTC = 1     ' Zeile von Datenspeicher

'Auslesen der Daten
For i = 1 To 169

If Iperm = "Iperm10" Or Iperm = "Iperm20" Then
    ws_NTCTool.Cells(z_NTCTool, 2) = ws_ADF.Cells(z_NTC, 2) ' NTC ADF->CH
End If

    ws_NTCTool.Cells(z_NTCTool, 4) = ws_ADF.Cells(z_NTC, 2) ' NTC CH-FR
    ws_NTCTool.Cells(z_NTCTool, 5) = ws_ADF.Cells(z_NTC, 6) ' NTC FR->CH
    ws_NTCTool.Cells(z_NTCTool, 7) = ws_ADF.Cells(z_NTC, 10) 'Grundfalltag
    z_NTC = z_NTC + 1
    z_NTCTool = z_NTCTool + 1
Next

'*************************
'DE Einlesen
'**************************

s_Grenze = ws_Daten.Range("E3")

Set ws_ADF = Workbooks(s_Name).Sheets(s_Grenze)

' Zeile in ADF Bestimungshilfe Blatt NTC_ADF-CH_CHFR

z_NTCTool = 171 ' Zeile im Protokoll WVP
z_NTC = 1     ' Zeile von Datenspeicher

'Auslesen der Daten
For i = 1 To 169
    ws_NTCTool.Cells(z_NTCTool, 2) = ws_ADF.Cells(z_NTC, 2) ' NTC CH-DE Full Export
    ws_NTCTool.Cells(z_NTCTool, 3) = ws_ADF.Cells(z_NTC, 6) ' NTC CH-DE Transit
    z_NTC = z_NTC + 1
    z_NTCTool = z_NTCTool + 1
Next

'*************************
'ADF Einlesen
'**************************

s_Grenze = ws_Daten.Range("F3")

Set ws_ADF = Workbooks(s_Name).Sheets(s_Grenze)

' Zeile in ADF Bestimungshilfe Blatt NTC_ADF-CH_CHFR

z_NTCTool = 171 ' Zeile im Protokoll WVP
z_NTC = 1     ' Zeile von Datenspeicher

'Auslesen der Daten
For i = 1 To 169
    ws_NTCTool.Cells(z_NTCTool, 6) = ws_ADF.Cells(z_NTC, 2) ' NTC CH-DE Full Export
    
    z_NTC = z_NTC + 1
    z_NTCTool = z_NTCTool + 1
Next

w_Bestimmung.Close False
Weiter:

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%% NTC für W-0 einlesen %%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

s_Pfad = ws_Daten.Range("B3")

n_Version = 10

Do While n_Version > -1
    s_Name = ws_Daten.Range("C4") & n_Version & ".xlsx"

    If Dir(s_Pfad & s_Name) <> "" Then
        Application.EnableEvents = False
        Set w_Bestimmung = Workbooks.Open(s_Pfad & s_Name, ReadOnly:=True) 'muss noch geändert werden
        Application.EnableEvents = True
        Exit Do
    Else
        n_Version = n_Version - 1

        If n_Version = -1 Then
            MsgBox "Die CH-FR Werte für W-0 von der WVP wurden nicht gefunden! Werte manuel in den Reiter <<NTC ADF-CH und CH-FR>> kopieren.(Pfad: O:\MA-SO-D-2_D-1\_Planung_YYYY\NTC\NTC_Tool)"
'            Exit Sub
            GoTo Weiter2
        End If

    End If

Loop

'*************************
'FR Einlesen
'**************************

s_Grenze = ws_Daten.Range("D3")

Set ws_ADF = Workbooks(s_Name).Sheets(s_Grenze)

' Zeile in ADF Bestimungshilfe Blatt NTC_ADF-CH_CHFR

z_NTCTool = 2 ' Zeile im Protokoll WVP
z_NTC = 1     ' Zeile von Datenspeicher

'Auslesen der Daten
For i = 1 To 169
    ws_NTCTool.Cells(z_NTCTool, 4) = ws_ADF.Cells(z_NTC, 2) ' NTC CH-FR
    ws_NTCTool.Cells(z_NTCTool, 5) = ws_ADF.Cells(z_NTC, 6) ' NTC FR->CH
    ws_NTCTool.Cells(z_NTCTool, 7) = ws_ADF.Cells(z_NTC, 10) 'Grundfalltag
    z_NTC = z_NTC + 1
    z_NTCTool = z_NTCTool + 1
Next


'*************************
'DE Einlesen
'**************************

s_Grenze = ws_Daten.Range("E3")

Set ws_ADF = Workbooks(s_Name).Sheets(s_Grenze)

' Zeile in ADF Bestimungshilfe Blatt NTC_ADF-CH_CHFR

z_NTCTool = 2 ' Zeile im Protokoll WVP
z_NTC = 1     ' Zeile von Datenspeicher


'Auslesen der Daten
For i = 1 To 169

    ws_NTCTool.Cells(z_NTCTool, 2) = ws_ADF.Cells(z_NTC, 2) ' NTC CH-DE Full Export
    ws_NTCTool.Cells(z_NTCTool, 3) = ws_ADF.Cells(z_NTC, 6) ' NTC CH-DE Transit
    z_NTC = z_NTC + 1
    z_NTCTool = z_NTCTool + 1
Next

'*************************
'ADF Einlesen
'**************************

s_Grenze = ws_Daten.Range("F3")

Set ws_ADF = Workbooks(s_Name).Sheets(s_Grenze)

' Zeile in ADF Bestimungshilfe Blatt NTC_ADF-CH_CHFR

z_NTCTool = 2 ' Zeile im Protokoll WVP
z_NTC = 1     ' Zeile von Datenspeicher

'Auslesen der Daten
For i = 1 To 169

    ws_NTCTool.Cells(z_NTCTool, 6) = ws_ADF.Cells(z_NTC, 2) ' NTC CH-DE Full Export
    
    z_NTC = z_NTC + 1
    z_NTCTool = z_NTCTool + 1
Next



w_Bestimmung.Close False

Weiter2:


ws_NTCTool.Range("A171:G171").ClearContents


    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    
End Sub
