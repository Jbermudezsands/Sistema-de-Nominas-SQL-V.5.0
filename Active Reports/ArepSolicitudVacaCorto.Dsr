VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepSolicitudVacaCorto 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12960
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   22860
   _ExtentY        =   19368
   SectionData     =   "ArepSolicitudVacaCorto.dsx":0000
End
Attribute VB_Name = "ArepSolicitudVacaCorto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PageHeader_Format()
Me.LblFechaHoy.Caption = Format(Now, "Long Date")
Me.LblTitulo.Caption = Titulo
Me.LblSubtitulo.Caption = SubTitulo

    If Dir(RutaLogo, vbDirectory) <> "" Then
        Me.ImgLogo.Picture = LoadPicture(RutaLogo)
    End If
    
End Sub
