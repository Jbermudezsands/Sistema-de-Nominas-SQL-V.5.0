VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepBajas 
   Caption         =   "Reporte de las Bajas"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ArepBajas.dsx":0000
End
Attribute VB_Name = "ArepBajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PageHeader_Format()
    FrmBajas.AdoDatosEmpresa.RecordSource = "DatosEmpresa"
    FrmBajas.AdoDatosEmpresa.Refresh
    ArepBajas.LblTitulo.Caption = FrmBajas.AdoDatosEmpresa.Recordset("NombreEmpresa")
    ArepBajas.LblSubtitulo.Caption = Subtitulo
    ArepBajas.ImgLogo.Picture = LoadPicture(RutaLogo)
End Sub
