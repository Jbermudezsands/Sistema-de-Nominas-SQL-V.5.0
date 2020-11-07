VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepColillaProduccionLegal 
   Caption         =   "ArepColillas"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ArepColillaProduccionLegal.dsx":0000
End
Attribute VB_Name = "ArepColillaProduccionLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_ReportStart()
                        Me.LblPeriodo.Caption = FrmNominaActiva.AdoBusca.Recordset("FechaNominaINI") & " al " & FrmNominaActiva.AdoBusca.Recordset("FechaNomina")
                        Me.LblTitulo.Caption = Titulo
End Sub

Private Sub Detail_Format()

Me.LblTotalDeduccion.Caption = Format(Val(Me.FldAdelanto.Text) + Val(Me.FldDeducciones.Text) + Val(Me.FldDescuento.Text), "#,##0.00")
Me.LblDeducciones.Caption = Format(Val(Me.FldAdelanto.Text) + Val(Me.FldDeducciones.Text) + Val(Me.FldDescuento.Text), "#,##0.00")
End Sub

