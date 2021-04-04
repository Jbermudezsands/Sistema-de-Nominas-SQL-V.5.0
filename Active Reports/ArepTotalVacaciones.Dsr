VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepTotalVacaciones 
   Caption         =   "ArepTotalVacaciones"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19420
   SectionData     =   "ArepTotalVacaciones.dsx":0000
End
Attribute VB_Name = "ArepTotalVacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
    Me.txtAcumuladas.Text = Format(Me.txtAcumuladas.Text, "##,##0.00")
    Me.txtDisponibles.Text = Format(Me.txtDisponibles.Text, "##,##0.00")
    Me.txtDisfrutadas.Text = Format(Me.txtDisfrutadas.Text, "##,##0.00")
End Sub

Private Sub PageHeader_Format()
 Me.LblDesde.Caption = DateTime.Now

 Me.LblTitulo.Caption = Titulo
       Me.LblSubtitulo.Caption = SubTitulo
       If Dir(RutaLogo) <> "" Then
         Me.ImgLogo.Picture = LoadPicture(RutaLogo)
       End If
       
End Sub
