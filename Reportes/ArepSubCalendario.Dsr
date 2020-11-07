VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepSub 
   Caption         =   "SubReporte"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   16642
   SectionData     =   "ArepSubCalendario.dsx":0000
End
Attribute VB_Name = "ArepSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
Dim Fecha1 As String, Fecha2 As String

Fecha1 = Format(Me.Field1.Text, "dd/mm/yyyy")
Fecha2 = Format(Me.Field2.Text, "dd/mm/yyyy")

Me.lblFecha.Caption = Fecha1 & " al " & Fecha2

End Sub
