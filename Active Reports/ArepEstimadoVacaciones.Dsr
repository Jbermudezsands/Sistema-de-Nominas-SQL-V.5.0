VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepEstimadoVacaciones 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20340
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35878
   _ExtentY        =   19420
   SectionData     =   "ArepEstimadoVacaciones.dsx":0000
End
Attribute VB_Name = "ArepEstimadoVacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PageHeader_Format()
       Me.LblTitulo.Caption = Titulo
       Me.LblSubtitulo.Caption = SubTitulo
       If Dir(RutaLogo) <> "" Then
         Me.ImgLogo.Picture = LoadPicture(RutaLogo)
       End If
       Me.LblDesde.Caption = "Desde " & FrmReportes.TxtFecha1.Value & " Hasta " & FrmReportes.TxtFecha2.Value
End Sub
