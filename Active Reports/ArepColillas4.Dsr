VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepColillas 
   Caption         =   "ArepColillas"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "ArepColillas.dsx":0000
End
Attribute VB_Name = "ArepColillas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()

Me.LblTotalDeduccion.Caption = Format(Val(Me.FldAdelanto.Text) + Val(Me.FldDeducciones.Text) + Val(Me.FldDescuento.Text), "#,##0.00")
Me.LblDeducciones.Caption = Format(Val(Me.FldAdelanto.Text) + Val(Me.FldDeducciones.Text) + Val(Me.FldDescuento.Text), "#,##0.00")
End Sub

