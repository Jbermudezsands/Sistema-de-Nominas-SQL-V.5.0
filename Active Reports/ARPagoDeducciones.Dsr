VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARPagoDeducciones 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "ARPagoDeducciones.dsx":0000
End
Attribute VB_Name = "ARPagoDeducciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Dim SqlPagoDeduccion As String
DaoDtaPagoDeducciones.DatabaseName = ruta
DaoDtaPagoDeducciones.ConnectionString = Conexion

SqlPagoDeduccion = "SELECT DetalleDeduccion.NumDeduccion, TipoDeduccion.Deduccion, Deduccion.CodTipoDeduccion, Deduccion.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.Pagado, DetalleDeduccion.NumNomina FROM TipoDeduccion INNER JOIN (Empleado INNER JOIN (Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion) ON Empleado.CodEmpleado = Deduccion.CodEmpleado) ON (TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion) AND (TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion) WHERE DetalleDeduccion.Pagado=True AND DetalleDeduccion.NumNomina= " & NumNomina & ""
DaoDtaPagoDeducciones.RecordSource = SqlPagoDeduccion
DaoDtaPagoDeducciones.Refresh
lbltitulo.Caption = Titulo
LblSubtitulo.Caption = Subtitulo
ImgLogo.Picture = LoadPicture(RutaLogo)
End Sub

