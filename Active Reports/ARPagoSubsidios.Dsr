VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ARPagoSubsidios 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "ARPagoSubsidios.dsx":0000
End
Attribute VB_Name = "ARPagoSubsidios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Dim SqlPagoSubsidios As String
DaoDtaPagoSubsidios.DatabaseName = ruta
DaoDtaPagoSubsidios.ConnectionString = Conexion

SqlPagoSubsidios = "SELECT DetalleSubsidio.NumSubsidio, TipoSubsidio.Subsidio, Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleSubsidio.Valor, DetalleSubsidio.NumVez, DetalleSubsidio.Pagado, DetalleSubsidio.NumNominaSubsidio FROM TipoSubsidio INNER JOIN ((Empleado INNER JOIN Subsidio ON Empleado.CodEmpleado = Subsidio.CodEmpleado) INNER JOIN DetalleSubsidio ON Subsidio.NumSubsidio = DetalleSubsidio.NumSubsidio) ON TipoSubsidio.CodTipoSubsidio = Subsidio.CodTipoSubsidio WHERE DetalleSubsidio.Pagado=True AND DetalleSubsidio.NumNominaSubsidio= " & NumNominaSubsidio & ""
DaoDtaPagoSubsidios.RecordSource = SqlPagoSubsidios
DaoDtaPagoSubsidios.Refresh
lbltitulo.Caption = Titulo
LblSubtitulo.Caption = Subtitulo
ImgLogo.Picture = LoadPicture(RutaLogo)
End Sub

