VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepColillaVaca 
   Caption         =   "ArepColillas"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19368
   SectionData     =   "ArepColillaVaca.dsx":0000
End
Attribute VB_Name = "ArepColillaVaca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
Dim CodEmpleado As String, NumNomina As Double
Dim Dias As Double, DiasDescuento As Double, DiasNeto As Double

If Me.Field17.Text = "" Then
  NumNomina = 0
Else
  NumNomina = Me.Field17.Text
End If





CodEmpleado = Me.Field14.Text
FrmReportes.AdoConsulta.RecordSource = "SELECT * From Reembolso WHERE (NumNomina = " & NumNomina & ") AND (CodEmpleado = '" & CodEmpleado & "')"
FrmReportes.AdoConsulta.Refresh
If Not FrmReportes.AdoConsulta.Recordset.EOF Then
   Me.LblAjustes.Caption = FrmReportes.AdoConsulta.Recordset("Monto")
Else
   Me.LblAjustes.Caption = "0.00"
End If

If Me.FldDiasPagar.Text = "" Then
  Dias = 0
Else
  Dias = Me.FldDiasPagar.Text
End If

If Me.FldDiasDescuento.Text = "" Then
  DiasDescuento = 0
Else
  DiasDescuento = Me.FldDiasDescuento.Text
End If
DiasNeto = Dias - DiasDescuento
Me.FldDiasNeto.Text = Format(DiasNeto, "##,##0.00")

'Me.LblTotalDeduccion.Caption = Format(Val(Me.FldAdelanto.Text) + Val(Me.FldDeducciones.Text) + Val(Me.FldDescuento.Text), "#,##0.00")
'Me.LblDeducciones.Caption = Format(Val(Me.FldAdelanto.Text) + Val(Me.FldDeducciones.Text) + Val(Me.FldDescuento.Text), "#,##0.00")
End Sub

