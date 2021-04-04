VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepListaBajas 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ListadoBajas.dsx":0000
End
Attribute VB_Name = "ArepListaBajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public TotalNetoPagar As Double
 
Private Sub ActiveReport_ReportStart()


  Me.lblTitulo.Caption = Titulo
  Me.LblSubtitulo.Caption = SubTitulo
End Sub

Private Sub Detail_Format()
 Dim TotalIngresos As Double, TotalEgresos As Double, TotalVacaciones As Double, TotalAguinaldo As Double, TotalAntiguedad As Double, MontoHRSExtras As Double, TotalOtrosSalarios As Double
 Dim TotalInss As Double, MontoIr As Double, Prestamo As Double, Deducciones As Double

  TotalVacaciones = Me.FldVacaciones.Text
  TotalAguinaldo = Me.FldAguinaldo.Text
  TotalAntiguedad = Me.FldAntiguedad.Text
  TotalOtrosSalarios = Me.FldOtrosIngresos.Text
  TotalInss = Me.FldMontoInss.Text
  MontoIr = Me.FldMontoIr.Text
  Prestamo = Me.FldPrestamos.Text
  Deducciones = Me.FldDeducciones.Text

    TotalIngresos = (TotalVacaciones + TotalAguinaldo + TotalAntiguedad + TotalOtrosSalarios)
    TotalEgresos = TotalInss + MontoIr + Prestamo + Deducciones


  Me.LblNetoPagar.Caption = Format(TotalIngresos - TotalEgresos, "##,##0.00")
  TotalNetoPagar = TotalNetoPagar + TotalIngresos - TotalEgresos
    
End Sub

Private Sub GroupFooter1_Format()
  Me.LblTotalNetoPagar.Caption = Format(TotalNetoPagar, "##,##0.00")
End Sub

Private Sub GroupHeader1_Format()
  TotalNetoPagar = 0
End Sub

