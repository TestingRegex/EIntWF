Option Explicit


Private Sub bnt_NSA_MVP_Click()

Call W1_Ergebnisse.NSA_Einlesen

ThisWorkbook.Worksheets("NSA Ergebnisse").btn_MVP_NSA.Visible = True

ThisWorkbook.Worksheets("NSA Ergebnisse").Select

End Sub