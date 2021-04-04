VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arepAsistenciaDepto2 
   Caption         =   "Reporte de Asistencia por Departamento"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20340
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35878
   _ExtentY        =   19420
   SectionData     =   "arepAsistenciaDepto2.dsx":0000
End
Attribute VB_Name = "arepAsistenciaDepto2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalEmpleados As Double, TotalEmpleadosSinMarcar As Double, TotalEmpleadosConMarcas As Double


Private Sub ActiveReport_ReportStart()
   TotalEmpleados = 0
   TotalEmpleadosSinMarcar = 0
   TotalEmpleadosConMarcas = 0
   
End Sub

Private Sub Detail_Format()
  Dim SqlString As String, FechaEntrada As Date, CodEmpleado As String
  
  If Me.txtFechaEntrada.Text <> "" Then
   FechaEntrada = Me.txtFechaEntrada.Text
  End If
  
  If Me.TxtCodEmpleado.Text <> "" Then
    CodEmpleado = Me.TxtCodEmpleado.Text
  End If
  
  
  
  SqlString = "SELECT SUM(DiasDisfrutados) AS DiasDisfrutados From SolicitudVacaciones WHERE (TipoSolicitud = 'Permiso Programado') AND (FechaSolicitud = CONVERT(DATETIME, '" & Format(FechaEntrada, "yyyy-mm-dd") & "', 102)) AND (CodigoEmpleado = '" & CodEmpleado & "')"
  MDIPrimero.AdoConsulta.ConnectionString = Conexion
  MDIPrimero.AdoConsulta.RecordSource = SqlString
  MDIPrimero.AdoConsulta.Refresh
  Me.TxtPermisos.Caption = "0.00"
  If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
   If Not IsNull(MDIPrimero.AdoConsulta.Recordset("DiasDisfrutados")) Then
      Me.TxtPermisos.Caption = MDIPrimero.AdoConsulta.Recordset("DiasDisfrutados")
   End If
  Else
      Me.TxtPermisos.Caption = "0.00"
  End If
  

End Sub

Private Sub GroupFooter1_Format()
  
  TotalEmpleados = Me.FldTotalEmpleado.Text
  Me.LblTotalEmpleadosSinMarcacion.Caption = TotalEmpleadosSinMarcar
  Me.LblTotalEmpleadosConMarcacion.Caption = TotalEmpleados - TotalEmpleadosSinMarcar
  

  
End Sub

Private Sub GroupFooter2_Format()
   Dim TotalHoras As Double

TotalHoras = Me.FldTotalHoras.Text
  If TotalHoras = 0 Then
    TotalEmpleadosSinMarcar = TotalEmpleadosSinMarcar + 1
  End If
End Sub

Private Sub GroupHeader1_Format()
    TotalEmpleados = 0
   TotalEmpleadosSinMarcar = 0
   TotalEmpleadosConMarcas = 0
End Sub
