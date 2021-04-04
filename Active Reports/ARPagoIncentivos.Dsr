VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARPagoIncentivos 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "ARPagoIncentivos.dsx":0000
End
Attribute VB_Name = "ARPagoIncentivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Dim SqlPagoIncentivo As String
DaoDtaPagoIncentivos.DatabaseName = ruta
DaoDtaPagoIncentivos.ConnectionString = Conexion

SqlPagoIncentivo = "SELECT DetalleIncentivo.NumIncentivo, Incentivo.CodTipoIncentivo, TipoIncentivo.Incentivo, Incentivo.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleIncentivo.Valor, DetalleIncentivo.NumVez, DetalleIncentivo.Pagado, DetalleIncentivo.NumNomina FROM Empleado INNER JOIN (TipoIncentivo INNER JOIN (Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo) ON TipoIncentivo.CodTipoIncentivo = Incentivo.CodTipoIncentivo) ON Empleado.CodEmpleado = Incentivo.CodEmpleado WHERE DetalleIncentivo.Pagado=True AND DetalleIncentivo.NumNomina= " & NumNomina & ""
DaoDtaPagoIncentivos.RecordSource = SqlPagoIncentivo
DaoDtaPagoIncentivos.Refresh
lbltitulo.Caption = Titulo
LblSubtitulo.Caption = Subtitulo
ImgLogo.Picture = LoadPicture(RutaLogo)

End Sub

