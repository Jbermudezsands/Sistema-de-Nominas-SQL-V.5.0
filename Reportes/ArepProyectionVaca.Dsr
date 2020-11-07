VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepProyectionVaca 
   Caption         =   "Reporte de Proyeccion de Vacaciones"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepProyectionVaca.dsx":0000
End
Attribute VB_Name = "ArepProyectionVaca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
     Me.LblTitulo3.Caption = "Reporte de Proyecciones de Vacaciones"
     Me.DataControl1.ConnectionString = ConexionReporte
     Me.LblTitulo.Caption = Titulo
     Me.LblSubtitulo.Caption = SubTitulo
     Me.LblDesde.Caption = "Desde: " & Format(FrmReportes.TxtFecha1.Value, "dd/mm/yyyy") & " Hasta: " & Format(FrmReportes.TxtFecha2.Value, "dd/mm/yyyy")
     If Dir(RutaLogo) <> "" Then
       Me.ImgLogo.Picture = LoadPicture(RutaLogo)
     End If
     
End Sub

