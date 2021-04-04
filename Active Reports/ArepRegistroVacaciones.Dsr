VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepRegistroVacaciones 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19420
   SectionData     =   "ArepRegistroVacaciones.dsx":0000
End
Attribute VB_Name = "ArepRegistroVacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
       Dim sql As String
       sql = "SELECT     NumeroSolicitud, FechaInicio, FechaFin, TipoSolicitud, DiasDisfrutar, Observaciones   FROM         SolicitudVacaciones   WHERE not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'   and  (CodigoEmpleado = '" & txtCodigo.Text & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(LblFechaIni, "dd/MM/yyyy") & " 00:00') AND (fechaInicio <= '" & Format(LblFechaFin, "dd/MM/yyyy") & " 23:59')"
       Set Me.SubReport1.object = New SubSolicitud
       Me.SubReport1.object.AdoSolicitud.ConnectionString = ConexionReporte
       Me.SubReport1.object.AdoSolicitud.Source = sql
End Sub

Private Sub PageHeader_Format()
        Me.LblTitulo.Caption = Titulo
       Me.LblSubtitulo.Caption = SubTitulo
       If Dir(RutaLogo) <> "" Then
         Me.ImgLogo.Picture = LoadPicture(RutaLogo)
       End If
       Me.LblDesde.Caption = "Desde " & FrmReportes.TxtFecha1.Value & " Hasta " & FrmReportes.TxtFecha2.Value
End Sub
