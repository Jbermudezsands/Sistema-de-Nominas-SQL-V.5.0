VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepConsolidadoVacaciones 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19420
   SectionData     =   "ArepConsolidadoVacaciones.dsx":0000
End
Attribute VB_Name = "ArepConsolidadoVacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
Me.txtSaldo.Text = Format(txtSaldo.Text, "##,##0.00")
End Sub

Private Sub PageHeader_Format()

       Me.txtFechaContratoVac.Text = Format(txtFechaContratoVac.Text, "dd/MM/yyyy")

       Me.LblTitulo.Caption = Titulo
       Me.LblSubtitulo.Caption = SubTitulo
       
       If Dir(RutaLogo) <> "" Then
         Me.ImgLogo.Picture = LoadPicture(RutaLogo)
       End If
       
      

End Sub
