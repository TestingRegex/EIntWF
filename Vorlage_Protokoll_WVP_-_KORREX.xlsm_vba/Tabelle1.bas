Option Explicit

Private Sub btn_Diagramm_ADF_Click()
ThisWorkbook.Worksheets("Übersicht").Shapes("Gruppieren 17").Visible = True
End Sub

Private Sub btn_Speichern_Click()
Dim s_Pfad As String, s_Name As String

s_Pfad = ThisWorkbook.Worksheets("Einstellungen").Range("B6").Text
s_Name = ThisWorkbook.Worksheets("Einstellungen").Range("C6").Text
    
ThisWorkbook.Save

ThisWorkbook.Worksheets("Übersicht").ExportAsFixedFormat Type:=xlTypePDF, Filename:=s_Pfad & s_Name & ".pdf", Quality:=xlQualityStandard, _
IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False


End Sub

Private Sub btn_Start_Einlesen_Click()

Call NTC_Werte.NTC_Einlesen

Call W1_Ergebnisse.NSA_Einlesen

ThisWorkbook.Worksheets("NSA Ergebnisse").btn_WVP_NSA.Visible = True

ThisWorkbook.Worksheets("NSA Ergebnisse").Select

End Sub
