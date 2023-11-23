Option Explicit

Private Sub btn_MVP_Click()
Dim s_Pfad As String, s_Name As String, WB As Workbook, d_Datum As Date, s_MM As String, n_MM As Integer, s_Datum As String, s_YY As String

ThisWorkbook.Worksheets("Einstellungen").Range("B11") = "MVP"

ThisWorkbook.Worksheets("NSA Ergebnisse").btn_WVP_NSA.Visible = False
ThisWorkbook.Worksheets("NSA Ergebnisse").btn_MVP_NSA.Visible = True

ThisWorkbook.Worksheets("Übersicht").Visible = False
ThisWorkbook.Worksheets("NTC ADF-CH und CH-FR").Visible = False
ThisWorkbook.Worksheets("MVP Übersicht").Visible = True


d_Datum = Date
s_Datum = d_Datum
s_MM = Mid(s_Datum, 4, 2)
n_MM = s_MM
s_YY = Right(s_Datum, 4)

    If n_MM + 2 < 10 Then
        s_MM = "0" & n_MM + 2
    ElseIf n_MM + 2 = 12 Then
        s_MM = "01"
    ElseIf n_MM + 2 = 13 Then
        s_MM = "02"
    Else
        s_MM = n_MM + 2
    End If
    

s_Datum = "01." & s_MM & "." & s_YY

ThisWorkbook.Worksheets("Einstellungen").Range("M3") = s_Datum
ThisWorkbook.Worksheets("Einstellungen").Range("B9") = 31


s_Pfad = ThisWorkbook.Worksheets("Einstellungen").Range("B10").Text
s_Name = ThisWorkbook.Worksheets("Einstellungen").Range("C10").Text


'Wenn das File für nächste Woche schon vorhanden ist öffne es und schliesse die vorlage
'Wenn es nicht existiert speichere die Vorlage mit dem richtigen Namen.
If Dir(s_Pfad & s_Name & ".xlsm") <> "" Then
    Set WB = Workbooks.Open(s_Pfad & s_Name & ".xlsm")
    ThisWorkbook.Close False
Else
    ThisWorkbook.SaveAs (s_Pfad & s_Name & ".xlsm")
End If



Prozesswahl.Hide
End Sub

Private Sub btn_WVP_Click()
Dim s_Pfad As String, s_Name As String, WB As Workbook, d_Datum As Date, BY As Integer, hilf As Range, zeile As Integer, spalte As Integer, NTCDE As Integer
Dim wert_start As Date, wert_ende As Date

ThisWorkbook.Worksheets("Einstellungen").Range("B11") = "WVP"

ThisWorkbook.Worksheets("NSA Ergebnisse").btn_WVP_NSA.Visible = True
ThisWorkbook.Worksheets("NSA Ergebnisse").btn_MVP_NSA.Visible = False
ThisWorkbook.Worksheets("Übersicht").Visible = True
ThisWorkbook.Worksheets("NTC ADF-CH und CH-FR").Visible = True
ThisWorkbook.Worksheets("MVP Übersicht").Visible = False


'Datums definition, Wählt den Montag von nächster Woche. 2 = vb_Monday
d_Datum = Date

If Weekday(d_Datum) <> 2 Then
    Do While Weekday(d_Datum) <> 2
        d_Datum = d_Datum + 1
    Loop
End If

'Datum in den Einstellungen speichern, wird für das einlesen der NSA Rechnungen benötigt!
ThisWorkbook.Worksheets("Einstellungen").Range("M3") = d_Datum
ThisWorkbook.Worksheets("Einstellungen").Range("B9") = 7


s_Pfad = ThisWorkbook.Worksheets("Einstellungen").Range("B6").Text
s_Name = ThisWorkbook.Worksheets("Einstellungen").Range("C6").Text

' Iperm Max und Min Werte korrekt in die Gesamtsteuerung kopieren
    BY = Format(d_Datum, "yyyy") ' Jahr des businessday
    Set hilf = ThisWorkbook.Worksheets("Einstellungen").Range("A:BB").Find("Iperm-Perioden") ' Suchen wo das Steht.
    zeile = hilf.Row + 1 ' Zeile eins nach unten
    spalte = hilf.Column + 1 ' zwei Spalten weiter
    
    Do
        NTCDE = ThisWorkbook.Worksheets("Einstellungen").Cells(zeile, spalte + 2) ' NTC Wert von ADF test ob was drin damit schlaufe nicht unendlich
        
        If NTCDE = 0 Then
            Exit Do
    End If
        
        
        wert_start = CDate(ThisWorkbook.Worksheets("Einstellungen").Cells(zeile, spalte) & BY)
        wert_ende = CDate(ThisWorkbook.Worksheets("Einstellungen").Cells(zeile, spalte + 1) & BY)
    
        If Format(d_Datum, "dd.mm.yyyy") >= CDate(wert_start) And Format(d_Datum, "dd.mm.yyyy") <= CDate(wert_ende) Then
            ThisWorkbook.Worksheets("Einstellungen").Range("P12") = ThisWorkbook.Worksheets("Einstellungen").Cells(zeile, spalte + 2) ' NTC CH-DE max
            ThisWorkbook.Worksheets("Einstellungen").Range("Q12") = ThisWorkbook.Worksheets("Einstellungen").Cells(zeile, spalte + 3) ' NTC CH-DE min
            ThisWorkbook.Worksheets("Einstellungen").Range("R12") = ThisWorkbook.Worksheets("Einstellungen").Cells(zeile, spalte + 4) ' NTC CH-FR max
            ThisWorkbook.Worksheets("Einstellungen").Range("S12") = ThisWorkbook.Worksheets("Einstellungen").Cells(zeile, spalte + 5) ' NTC CH-FR min
            ThisWorkbook.Worksheets("Einstellungen").Range("T12") = ThisWorkbook.Worksheets("Einstellungen").Cells(zeile, spalte + 6) ' NTC ADF-CH max
            ThisWorkbook.Worksheets("Einstellungen").Range("U12") = ThisWorkbook.Worksheets("Einstellungen").Cells(zeile, spalte + 7) ' NTC ADF-CH min
            ThisWorkbook.Worksheets("Einstellungen").Range("R16") = ThisWorkbook.Worksheets("Einstellungen").Cells(zeile, spalte + 8) ' NTC FR-CH max
            ThisWorkbook.Worksheets("Einstellungen").Range("S16") = ThisWorkbook.Worksheets("Einstellungen").Cells(zeile, spalte + 9) ' NTC FR-CH min
            ThisWorkbook.Worksheets("Einstellungen").Range("N12") = ThisWorkbook.Worksheets("Einstellungen").Cells(zeile, spalte - 1) 'Iperm eintragen
            Exit Do
        End If
    
        zeile = zeile + 1
    Loop

'Diagramm ADF und DE Visible anhand des Iperm Wertes
If ThisWorkbook.Worksheets("Einstellungen").Range("N12") <> "Iperm10" Then
    ThisWorkbook.Worksheets("Übersicht").Shapes("Gruppieren 17").Visible = False
End If
If ThisWorkbook.Worksheets("Einstellungen").Range("N12") = "Iperm10" Or ThisWorkbook.Worksheets("Einstellungen").Range("N12") = "Iperm20" Then
    ThisWorkbook.Worksheets("Übersicht").Shapes("Gruppieren 17").Visible = True
End If


'Wenn das File für nächste Woche schon vorhanden ist öffne es und schliesse die vorlage
'Wenn es nicht existiert speichere die Vorlage mit dem richtigen Namen.
If Dir(s_Pfad & s_Name & ".xlsm") <> "" Then
    Set WB = Workbooks.Open(s_Pfad & s_Name & ".xlsm")
    ThisWorkbook.Close False
Else
    ThisWorkbook.SaveAs (s_Pfad & s_Name & ".xlsm")
End If

Prozesswahl.Hide

End Sub




Private Sub UserForm_Click()

End Sub