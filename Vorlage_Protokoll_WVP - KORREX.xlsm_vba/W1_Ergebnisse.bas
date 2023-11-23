Option Explicit

Public Sub NSA_Einlesen()

With Application
    .DisplayAlerts = False
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .ScreenUpdating = False
End With

'Dim arrayCount As String
Dim AM As String, BB As String, MSN As String, DD As String, MM As String, Ordner As String, Dateiname As String, Verletzung As String
Dim D As Integer, M As Integer, YY As Integer, k As Integer, n As Integer, KK As Integer, i As Integer, zähle As Integer, Version As Integer, Z As Integer, nSpNSA As Integer
Dim n_ZUeb As Integer, n_SpZeitst As Integer, n_ZZeitst As Integer, n_Kontrolle As Integer, n_Eingetragen As Integer
Dim Datum As Date, hilf As Range, n_Tage As Integer, s_Prozess As String, s_ZeitSt As String, d_GrunfallTag As Date

Dim w_Me As Workbook, ws_Ergebnisse As Worksheet, ws_Einst As Worksheet, ws_Ueb As Worksheet, w_NSA As Workbook, ws_NSA As Worksheet

Set w_Me = ThisWorkbook
Set ws_Ergebnisse = w_Me.Sheets("NSA Ergebnisse")
Set ws_Einst = w_Me.Sheets("Einstellungen")
Set ws_Ueb = w_Me.Sheets("Übersicht")


ws_Ergebnisse.Range("A4:AJ74").ClearContents
n_Tage = ws_Einst.Range("B9")
Datum = ws_Einst.Range("M3")
s_Prozess = ws_Einst.Range("B11")


n_ZUeb = 27
i = 2 ' Spalte W-1 NSA Ergebnisse
'' Tagesscharf (Montg bis Sonntag)
For k = 1 To n_Tage

n_Eingetragen = 0

    YY = Year(Datum)
    M = Month(Datum)
    D = Day(Datum)
    
    If M < 10 Then
        MM = "0" & M
    Else
        MM = M
    End If
    
    
    If D < 10 Then
        DD = "0" & D
    Else
        DD = D
    End If
    
    Ordner = ws_Einst.Range("B5") & MM
    
Version = 10

Do While Version > -2
    
    Dateiname = Ordner & "\" & YY & MM & DD & "_VP_" & s_Prozess & "_" & Version & ".xlsm"
    If Dir(Dateiname) <> "" Then
        'MsgBox ("Datei gefunden")
        n_Eingetragen = n_Eingetragen + 1
        GoTo Öffnen
    Else
    Version = Version - 1
    End If
    
    If Version = -1 Then
'        MsgBox ("Datei nicht gefunden")
        ws_Ueb.Range("B" & n_ZUeb) = "Nein"
        GoTo Nächster_Tag
    End If
Loop


Öffnen:
    
n_Kontrolle = 1

'    Set appExcel = New Application 'Excel wird im Hintergrund geöffnet und ist somit nicht sichtbar
'    appExcel.Visible = False
    Application.EnableEvents = False
    Set w_NSA = Workbooks.Open(Dateiname) 'appExcel.Workbooks.Open(Dateiname)
    Set ws_NSA = w_NSA.Sheets("Übersicht")
    Application.EnableEvents = True
    zähle = 1
    n = 28
    Z = 4
    ws_Ueb.Range("B" & n_ZUeb) = "Ja"
    
    
    Do 'Schlaufe zum Auslesen und Eintragen der AM BB und Massnahmen
        
        AM = ws_NSA.Range("B" & n) ' Liest AM aus DACF Excel aus
        
        
        If AM = "Gerechnete Stunden:" Then
            n = n + 1
            AM = ws_NSA.Range("B" & n)
        End If
        
        BB = ws_NSA.Range("Q" & n) ' Liest BB aus DACF Excel aus
        MSN = ws_NSA.Range("AF" & n) ' Liest Massnahme aus DACF Excel aus
        'AM = Trim(AM)
        
        nSpNSA = 57
        
        
        
        Do
            If ws_NSA.Cells(n, nSpNSA) <> "" Then
                Verletzung = ws_NSA.Cells(n, nSpNSA)
                
                If n_Kontrolle = 1 Then
                ' nach dem Zeitstempel suchen und dann in das Übersichtsblatt kopieren
                Set hilf = ws_Einst.Range("U1:Z1100").Find(nSpNSA)
                n_SpZeitst = hilf.Column
                n_ZZeitst = hilf.Row
                ws_Ueb.Range("B" & n_ZUeb) = ws_Ueb.Range("B" & n_ZUeb) & "  (" & ws_Einst.Cells(n_ZZeitst, n_SpZeitst + 1).Text & ")"
                s_ZeitSt = ws_Einst.Cells(n_ZZeitst, n_SpZeitst + 1).Text
                n_Kontrolle = n_Kontrolle + 1
                End If
                
                Exit Do
            Else
                nSpNSA = nSpNSA + 1
            End If
            
            If nSpNSA = 81 Then
                Exit Do
            End If
        Loop
        
        
        
        If MSN = "Es wurden keine Verletzungen gefunden" Then
            Exit Do
        End If
        
        If AM = "" Then
            Exit Do
        End If
                
        ' Schreibt die AM BB und Massnahmen in entsprechende Zeilen
        ws_Ergebnisse.Cells(Z, i) = AM
        ws_Ergebnisse.Cells(Z, i + 1) = BB
        ws_Ergebnisse.Cells(Z, i + 2) = MSN
        ws_Ergebnisse.Cells(Z, i + 3) = Verletzung
        n = n + 1
        Z = Z + 1
        zähle = zähle + 1
        
    Loop


    d_GrunfallTag = ws_NSA.Range("AQ5")
    


 w_NSA.Close True
Nächster_Tag:
    
    If s_Prozess = "WVP" Then
    n_Eingetragen = 1
    ws_Ergebnisse.Cells(2, i) = Format(Datum, "DDDD, DD.MM.YYYY") & ":  (GF: " & d_GrunfallTag & " )"
    End If
    
    If s_Prozess = "MVP" Then
    ws_Ergebnisse.Cells(2, i) = Format(Datum, "DDDD, DD.MM.YYYY") & " (" & s_ZeitSt & ")"
    End If
    

n_ZUeb = n_ZUeb + 1
Datum = Datum + 1



' für MVP wenn nichts eingetragen wird nicht hochzählen
If n_Eingetragen <> 0 Then
    i = i + 3
    ' Überspringen der Überprüfung ob Massnahme vorhanden
    
    If zähle = 4 Or 8 Or 12 Or 16 Or 20 Or 24 Or 28 Then
        i = i + 2
    End If
End If

Next
        
        With Application
            .DisplayAlerts = True
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .ScreenUpdating = True
        End With

ws_Ergebnisse.Columns("A:AB").AutoFit
'ThisWorkbook.Worksheets("DACF_Masnahmen").Range("A1:U15").Rows.EntireRow.AutoFit
ws_Ergebnisse.Select

End Sub

Public Sub Ergebnisse_kopieren() ' Für WVP Übersicht
Dim w_Me As Workbook, ws_Ergebnisse As Worksheet, ws_Einst As Worksheet, ws_Ueb As Worksheet, w_NSA As Workbook, ws_NSA As Worksheet

Set w_Me = ThisWorkbook
Set ws_Ergebnisse = w_Me.Sheets("NSA Ergebnisse")
Set ws_Einst = w_Me.Sheets("Einstellungen")
Set ws_Ueb = w_Me.Sheets("Übersicht")


Dim n_ZErgeb As Integer, n_SpErgeb As Integer, n_ZUeb As Integer, n_Zahl As Integer
Dim s_Ergebnisse As String

n_ZErgeb = 4
n_SpErgeb = 2
n_ZUeb = 27

Do ' Schlaufe für Tage
    s_Ergebnisse = ""
    n_ZErgeb = 4

    Do ' Schlaufe für einträge pro Tag
        If ws_Ergebnisse.Cells(n_ZErgeb, n_SpErgeb).Interior.ColorIndex <> xlNone Or ws_Ergebnisse.Cells(n_ZErgeb, n_SpErgeb - 1) = "x" Then
            'Sting erstellen für Übersicht
            
            s_Ergebnisse = s_Ergebnisse & ws_Ergebnisse.Cells(n_ZErgeb, n_SpErgeb + 3).Text & "  " & ws_Ergebnisse.Cells(n_ZErgeb, n_SpErgeb) & " bei Ausfall " & ws_Ergebnisse.Cells(n_ZErgeb, n_SpErgeb + 1) & vbCrLf & "Massnahme: " & ws_Ergebnisse.Cells(n_ZErgeb, n_SpErgeb + 2) & vbCrLf
            'kopieren der ergebnisse auf die Übersichtsseite
            
            ws_Ueb.Range("C" & n_ZUeb) = s_Ergebnisse
        End If
    
        ' Schlaufe Verlasen wenn keine Einträge
        If ws_Ergebnisse.Cells(n_ZErgeb, n_SpErgeb) = "" Then
            Exit Do
        End If
    
    n_ZErgeb = n_ZErgeb + 1
    
    Loop

n_ZUeb = n_ZUeb + 1

If n_SpErgeb > 36 Then
    Exit Do
End If

n_SpErgeb = n_SpErgeb + 5
Loop

'Einfärben der Ergebnisse und Masnahmen in farbe
'For n_ZUeb = 27 To 33
'
'ws_Ueb.Range("C" & n_ZUeb).Characters(Start:=1, Length:=4).Font.ColorIndex = 3 'Rot
'
'n_Zahl = InStr(1, ws_Ueb.Range("C" & n_ZUeb), "Massnahme:")
'
'ws_Ueb.Range("C" & n_ZUeb).Characters(Start:=n_Zahl, Length:=10).Font.ColorIndex = 10 'grün
'
'
'Next



ThisWorkbook.Worksheets("Übersicht").Select

End Sub


Public Sub Ergebnisübertragen() ' Für MVP Übersicht
Dim w_Me As Workbook, ws_Ergebnisse As Worksheet, ws_Einst As Worksheet, ws_Ueb As Worksheet, w_NSA As Workbook, ws_NSA As Worksheet

Set w_Me = ThisWorkbook
Set ws_Ergebnisse = w_Me.Sheets("NSA Ergebnisse")
Set ws_Einst = w_Me.Sheets("Einstellungen")
Set ws_Ueb = w_Me.Sheets("MVP Übersicht")


ws_Ueb.Range("A4:G200").ClearContents

Dim n_ZErgeb As Integer, n_SpErgeb As Integer, n_ZUeb As Integer, n_Zahl As Integer, n_Zählen As Integer
Dim s_Zeitstempel As String, s_AM As String, s_BB As String, s_Verletzung As String, s_Massnahmen As String


n_ZErgeb = 4
n_SpErgeb = 2
n_ZUeb = 4
n_Zählen = 1

Do While n_Zählen < ws_Einst.Range("B9") ' Schlaufe für Tage

    Do ' Schlaufe für einträge pro Tag
        If ws_Ergebnisse.Cells(n_ZErgeb, n_SpErgeb).Interior.ColorIndex <> xlNone Or ws_Ergebnisse.Cells(n_ZErgeb, n_SpErgeb - 1) = "x" Then
            'Sting erstellen für Übersicht
            
            ws_Ueb.Cells(n_ZUeb, 2) = ws_Ergebnisse.Cells(2, n_SpErgeb).Text 'ws_Ergebnisse.Cells(n_ZErgeb, 2).Text ' Zeitstempel kopieren
            ws_Ueb.Cells(n_ZUeb, 3) = ws_Ergebnisse.Cells(n_ZErgeb, n_SpErgeb).Text ' s_AM
            ws_Ueb.Cells(n_ZUeb, 4) = ws_Ergebnisse.Cells(n_ZErgeb, n_SpErgeb + 1).Text ' BB
            ws_Ueb.Cells(n_ZUeb, 5) = ws_Ergebnisse.Cells(n_ZErgeb, n_SpErgeb + 3).Text ' Verletzung
            ws_Ueb.Cells(n_ZUeb, 6) = ws_Ergebnisse.Cells(n_ZErgeb, n_SpErgeb + 2).Text ' Massnahmen
            
            'kopieren der ergebnisse auf die Übersichtsseite
            n_ZUeb = n_ZUeb + 1
        End If
    
        ' Schlaufe Verlasen wenn keine Einträge
        If ws_Ergebnisse.Cells(n_ZErgeb, n_SpErgeb) = "" Then
            Exit Do
        End If
    
    n_ZErgeb = n_ZErgeb + 1
    
    Loop
    n_ZErgeb = 4

n_SpErgeb = n_SpErgeb + 5
n_Zählen = n_Zählen + 1

Loop

ws_Ueb.Select

End Sub


















