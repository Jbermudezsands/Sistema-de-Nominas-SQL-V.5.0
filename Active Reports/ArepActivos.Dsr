VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepActivos 
   Caption         =   "Reporte de Empleados"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ArepActivos.dsx":0000
End
Attribute VB_Name = "ArepActivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
       Me.LblTitulo.Caption = Titulo
       Me.LblSubtitulo.Caption = SubTitulo
       If Dir(RutaLogo) <> "" Then
         Me.ImgLogo.Picture = LoadPicture(RutaLogo)
       End If
       Me.LblImpreso.Caption = "Impreso: " & Format(Now, "dd/mm/yyyy")
End Sub

