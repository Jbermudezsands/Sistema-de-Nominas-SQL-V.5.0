VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepColilla13vo 
   Caption         =   "ArepColillas"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepColilla13vo.dsx":0000
End
Attribute VB_Name = "ArepColilla13vo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ActiveReport_ReportStart()
  Me.LblTipoColilla.Caption = Quien
End Sub

Private Sub Detail_BeforePrint()

If CDbl(Me.txtDiasAcum.Text) >= 30 Then

'   Me.txtTotalPagar.Text = Me.txtMontoPagar.Text
   
Else

'   Me.txtTotalPagar.Text = Format((CDbl(Me.txtMontoPagar.Text) / 30.4167) * CDbl(Me.txtDiasAcum.Text), "##,###.#0")

End If


End Sub

Private Sub Detail_Format()

'Me.LblTotalDeduccion.Caption = Format(Val(Me.FldAdelanto.Text) + Val(Me.FldDeducciones.Text) + Val(Me.FldDescuento.Text), "#,##0.00")
'Me.LblDeducciones.Caption = Format(Val(Me.FldAdelanto.Text) + Val(Me.FldDeducciones.Text) + Val(Me.FldDescuento.Text), "#,##0.00")
End Sub

