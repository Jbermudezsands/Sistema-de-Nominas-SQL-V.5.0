VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepColillas 
   Caption         =   "ArepColillas"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19368
   SectionData     =   "ArepColillas.dsx":0000
End
Attribute VB_Name = "ArepColillas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()

Me.LblTotalDeduccion.Caption = Format(val(Me.FldAdelanto.Text) + val(Me.FldDeducciones.Text) + val(Me.FldDescuento.Text), "#,##0.00")
Me.LblDeducciones.Caption = Format(val(Me.FldAdelanto.Text) + val(Me.FldDeducciones.Text) + val(Me.FldDescuento.Text), "#,##0.00")
End Sub

