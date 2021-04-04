VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} SubSolicitud 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19420
   SectionData     =   "SubSolicitud.dsx":0000
End
Attribute VB_Name = "SubSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
    Me.Field4.Text = "No. Solicitud: " & Field4.Text
    Me.Field9.Text = "Desde: " & Format(Field9.Text, "dd/MM/yyyy")
    Me.Field10.Text = "Hasta: " & Format(Field10.Text, "dd/MM/yyyy")
    Me.Field6.Text = "Solicitado: " & Format(Field6.Text, "##,##0.00")
    Me.Field11.Text = "Observaciones: " & Field11.Text
End Sub

