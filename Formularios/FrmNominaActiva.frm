VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmNominaActiva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de la Nomina Activa"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoDepartamentos 
      Height          =   375
      Left            =   360
      Top             =   6240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoDepartamentos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   375
      Left            =   240
      Top             =   4200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoConsulta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   615
      Left            =   3240
      Picture         =   "FrmNominaActiva.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "FrmNominaActiva.frx":1E72
      Top             =   5160
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdExportaCSV 
      Caption         =   "Exp.CSV"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3240
      Picture         =   "FrmNominaActiva.frx":1E78
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   840
      Top             =   4320
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.ListBox CmbReportes 
      Height          =   2400
      ItemData        =   "FrmNominaActiva.frx":2182
      Left            =   120
      List            =   "FrmNominaActiva.frx":2184
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton CmdExportar 
      Caption         =   "Exportar"
      Height          =   615
      Left            =   3240
      Picture         =   "FrmNominaActiva.frx":2186
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   3240
      Picture         =   "FrmNominaActiva.frx":24C8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin MSAdodcLib.Adodc AdoBusca 
      Height          =   375
      Left            =   360
      Top             =   4920
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoBusca"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin XtremeSuiteControls.ProgressBar Barra 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   3015
      _Version        =   786432
      _ExtentX        =   5318
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
End
Attribute VB_Name = "FrmNominaActiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbReportes_Click()

Me.cmdImprimir.Enabled = True

 Select Case Me.CmbReportes.Text
   Case "Listado Horas Extra"
     Me.CmdExportaCsv.Enabled = True
   Case "Reporte x Produccion"
     Me.CmdExportaCsv.Enabled = True
   Case "Reporte Percepsiones"
     Me.CmdExportaCsv.Enabled = True
   Case "Nomina x Cargo"
      Me.CmdExportar.Enabled = True
   Case "Nomina x Cargo Resumen"
      Me.CmdExportar.Enabled = True
   Case "Exportar Nomina sin Formato"
      Me.cmdImprimir.Enabled = False
      Me.CmdExportaCsv.Enabled = False
   Case Else
     Me.CmdExportaCsv.Enabled = False
     Me.CmdExportar.Enabled = False
 End Select
End Sub

Private Sub CmdExportaCSV_Click()
On Error GoTo TipoErrs
Dim SQlReportes As String, Longitud As Integer, Respuesta As Integer
Dim Cadena As String, Mes As String, Dia As String, Ano As String
Dim TextoMonto As String, TipoMovimiento As String, j As Integer
Dim Codigo As String
salir = False
Me.Barra.Visible = True
Me.CommonDialog1.ShowSave
Directorio = ""
Directorio = Me.CommonDialog1.FileName + ".csv"
'Fecha1 = Year(Me.DTFecha1.Value) & "-" & Month(Me.DTFecha1.Value) & "-" & Day(Me.DTFecha1.Value)
'Fecha2 = Year(Me.DTFecha2.Value) & "-" & Month(Me.DTFecha2.Value) & "-" & Day(Me.DTFecha2.Value)

Select Case CmbReportes.Text
  Case "Listado Horas Extra"
     SQlReportes = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, HorasExtras.CantHoras, HorasExtras.NumNomina,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres FROM Empleado INNER JOIN HorasExtras ON Empleado.CodEmpleado = HorasExtras.CodEmpleado Where (((HorasExtras.CantHoras) <> 0) And ((HorasExtras.NumNomina) = " & NumNomina & ")) ORDER BY Empleado.CodEmpleado1"
  Case "Reporte x Produccion"
     SQlReportes = "SELECT DetalleProduccion.CodEmpleado, DetalleProduccion.NumNomina, DetalleProduccion.CodReferencia, DetalleProduccion.CodProceso,DetalleProduccion.Ref, DetalleProduccion.Lunes, DetalleProduccion.Martes, DetalleProduccion.Miercoles, DetalleProduccion.Jueves,DetalleProduccion.Viernes, DetalleProduccion.Sabado, DetalleProduccion.Domingo, DetalleProduccion.TotalUnidades, DetalleProduccion.SalarioPieza,DetalleProduccion.Precio , DetalleProduccion.unidad, DetalleProduccion.Pagado, Empleado.CodEmpleado1 FROM DetalleProduccion INNER JOIN Empleado ON DetalleProduccion.CodEmpleado = Empleado.CodEmpleado Where (DetalleProduccion.NumNomina = " & NumNomina & ")"

   Case "Reporte Percepsiones"
     SQlReportes = "SELECT SUM(SalarioBasico) AS SalarioBasico, SUM(Destajo) AS Produccion, SUM(HorasExtras) AS HorasExtra, SUM(Comisiones) AS Puntualidad, SUM(VacacionesPagadas) AS Vacaciones, SUM(SeptimoDia) AS SeptimoDia, SUM(IncetivoProduccion) AS IncentivosProduccion, SUM(Incentivos)AS Antiguedad, SUM(OtrosIngresos) AS OtrosIngresos, SUM(SalarioBasico) + SUM(Destajo) + SUM(HorasExtras) + SUM(Comisiones)+ SUM(SeptimoDia) + SUM(IncetivoProduccion + Incentivos) + SUM(OtrosIngresos) + SUM(VacacionesPagadas) AS TotalDevengado,SUM(Deducciones) AS Deducciones, SUM(Prestamo) AS Prestamo, SUM(MontoINSS) AS MontoInss, SUM(MontoIR) AS MontoIr, SUM(DiasDescuento) AS DiasDescuento, SUM(Adelantos) AS Adelantos, SUM(Deducciones) + SUM(Prestamo) + SUM(MontoINSS) + SUM(MontoIR) + SUM(DiasDescuento)+ SUM(Adelantos) AS TotalDeduccines, (SUM(SalarioBasico) + SUM(Destajo) + SUM(HorasExtras) + SUM(Comisiones) + SUM(SeptimoDia)" & vbLf
     SQlReportes = SQlReportes & "+ SUM(IncetivoProduccion + Incentivos) + SUM(OtrosIngresos) + SUM(VacacionesPagadas)) - (SUM(Deducciones) + SUM(Prestamo) + SUM(MontoINSS)" & vbLf
     SQlReportes = SQlReportes & "+ SUM(MontoIR) + SUM(DiasDescuento) + SUM(Adelantos)) AS Neto, SUM(INSSPatronal) AS InssPatronal, SUM(IRPatronal) AS IrPatronal, SUM(INATEC)" & vbLf
     SQlReportes = SQlReportes & "AS Inatec, NumNomina, SUM(HE) AS HE, SUM(HTrabajada) AS HTrabajada, SUM(INSSPatronal) + SUM(INATEC)" & vbLf
     SQlReportes = SQlReportes & "AS TotalObligaciones" & vbLf
     SQlReportes = SQlReportes & "From dbo.DetalleNomina" & vbLf
     SQlReportes = SQlReportes & "GROUP BY NumNomina" & vbLf
     SQlReportes = SQlReportes & "Having (NumNomina = " & NumNomina & ")" & vbLf
   
End Select
Me.AdoBusca.RecordSource = SQlReportes
AdoBusca.Refresh
Me.AdoBusca.Recordset.MoveLast
Maximo = AdoBusca.Recordset.RecordCount
If (Dir(Directorio) <> "") Then
  Respuesta = MsgBox("Reescribir el Archivo?", vbYesNo, "Enlace Pacioli")
  If Respuesta = 6 Then
     Kill (Directorio)
               Open Directorio For Output As #1
                     
                AdoBusca.Recordset.MoveFirst
                With Barra
                   .Min = 0
                   .Value = 0
                   .Max = Maximo
                   j = 0
                   
                Select Case CmbReportes.Text
                   Case "Listado Horas Extra"
                      Cadena = "CodEmpleado" & "," & "Nombres" & "," & "CantHoras" & "," & "NumNomina"
                   Case "Reporte x Produccion"
                      Cadena = "CodEmpleado" & "," & "CodReferencia" & "," & "CodProceso" & "," & "Precio" & "," & "Ref" & "," & "Lunes" & "," & "Martes" & "," & "Miercoles" & "," & "Jueves" & "," & "Viernes" & "," & "Sabado" & "," & "Domingo" & "," & "TotalUnidades" & "," & "SalarioPieza" & "," & "Unidad"
                   Case "Reporte Percepsiones"
                      Cadena = "SalarioBasico" & "," & "HTrabajada" & "," & "HE" & "," & "Produccion" & "," & "HorasExtra" & "," & "Puntualidad" & "," & "Vacaciones" & "," & "SeptimoDia" & "," & "IncentivosProduccion" & "," & "Antiguedad" & "," & "OtrosIngresos" & "," & "TotalDevengado" & "," & "Deducciones" & "," & "Prestamo" & "," & "MontoInss" & "," & "MontoIr" & "," & "DiasDescuento" & "," & "Adelantos" & "," & "TotalDeduccines" & "," & "Neto" & "," & "InssPatronal" & "," & "Inatec" & "," & "NumNomina"

                 End Select
                    Print #1, Cadena
                   
                 Do While Not AdoBusca.Recordset.EOF
                 '////////Inicialiso las variables/////////////////
                 Select Case CmbReportes.Text
                   Case "Listado Horas Extra"
                      Cadena = AdoBusca.Recordset("CodEmpleado1") & "," & AdoBusca.Recordset("Nombres") & "," & AdoBusca.Recordset("CantHoras") & "," & AdoBusca.Recordset("NumNomina")
                   Case "Reporte x Produccion"
                      Cadena = AdoBusca.Recordset("CodEmpleado1") & "," & AdoBusca.Recordset("CodReferencia") & "," & AdoBusca.Recordset("CodProceso") & "," & AdoBusca.Recordset("Precio") & "," & AdoBusca.Recordset("Ref") & "," & AdoBusca.Recordset("Lunes") & "," & AdoBusca.Recordset("Martes") & "," & AdoBusca.Recordset("Miercoles") & "," & AdoBusca.Recordset("Jueves") & "," & AdoBusca.Recordset("Viernes") & "," & AdoBusca.Recordset("Sabado") & "," & AdoBusca.Recordset("Domingo") & "," & AdoBusca.Recordset("TotalUnidades") & "," & AdoBusca.Recordset("SalarioPieza") & "," & AdoBusca.Recordset("Unidad")
                   Case "Reporte Percepsiones"
                      Cadena = AdoBusca.Recordset("SalarioBasico") & "," & AdoBusca.Recordset("HTrabajada") & "," & AdoBusca.Recordset("HE") & "," & AdoBusca.Recordset("Produccion") & "," & AdoBusca.Recordset("HorasExtra") & "," & AdoBusca.Recordset("Puntualidad") & "," & AdoBusca.Recordset("Vacaciones") & "," & AdoBusca.Recordset("SeptimoDia") & "," & AdoBusca.Recordset("IncentivosProduccion") & "," & AdoBusca.Recordset("Antiguedad") & "," & AdoBusca.Recordset("OtrosIngresos") & "," & AdoBusca.Recordset("TotalDevengado") & "," & AdoBusca.Recordset("Deducciones") & "," & AdoBusca.Recordset("Prestamo") & "," & AdoBusca.Recordset("MontoInss") & "," & AdoBusca.Recordset("MontoIr") & "," & AdoBusca.Recordset("DiasDescuento") & "," & AdoBusca.Recordset("Adelantos") & "," & AdoBusca.Recordset("TotalDeduccines") & "," & AdoBusca.Recordset("Neto") & "," & AdoBusca.Recordset("InssPatronal") & "," & AdoBusca.Recordset("Inatec") & "," & AdoBusca.Recordset("NumNomina")

                 End Select
                    Print #1, Cadena
                                    
                    
                    
                  AdoBusca.Recordset.MoveNext
                  j = j + 1
                  Me.Caption = "Procesando:  " & j & " de " & Maximo & " Registros "
                  DoEvents
                  .Value = j
                  Cadena = ""
                  Loop
                  End With
                  
                 Close #1

                MsgBox "La Exportacion, fue Creada con Exito", vbExclamation, "Sistema de Enlace"
                salir = True
  End If
Else '//////En caso que no exista el Archivo///////////
                
                Open Directorio For Output As #1
                'SQLExporta = "SELECT Empleado.CodEmpleado, Empleado.CodDepartamento, Historico.CodCuenta, Historico.CuentaCredito, DetalleNomina.NumNomina, Nomina.Fecha, [DetalleNomina]![SalarioBasico]+[DetalleNomina]![Destajo]+[DetalleNomina]![HorasExtras]+[DetalleNomina]![Comisiones]+[DetalleNomina]![Incentivos]-[DetalleNomina]![Deducciones]-[DetalleNomina]![Prestamo]-[DetalleNomina]![MontoINSS]-[DetalleNomina]![MontoIR]+[DetalleNomina]![TotalSubsidio] AS GranTotal FROM Nomina INNER JOIN ((Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado) ON Nomina.NumNomina = DetalleNomina.NumNomina Where DetalleNomina.NumNomina = " & NumNomina & " ORDER BY Empleado.CodEmpleado"
                
                AdoBusca.Recordset.MoveFirst
                With Barra
                   .Min = 0
                   .Value = 0
                   .Max = Maximo
                   j = 0
                 Do While Not AdoBusca.Recordset.EOF
                 Select Case CmbReportes.Text
                   Case "Listado Horas Extra"
                      Cadena = AdoBusca.Recordset("CodEmpleado1") & "," & AdoBusca.Recordset("Nombres") & "," & AdoBusca.Recordset("CantHoras") & "," & AdoBusca.Recordset("NumNomina")
                   Case "Reporte x Produccion"
                      Cadena = AdoBusca.Recordset("CodEmpleado1") & "," & AdoBusca.Recordset("CodReferencia") & "," & AdoBusca.Recordset("CodProceso") & "," & AdoBusca.Recordset("Precio") & "," & AdoBusca.Recordset("Ref") & "," & AdoBusca.Recordset("Lunes") & "," & AdoBusca.Recordset("Martes") & "," & AdoBusca.Recordset("Miercoles") & "," & AdoBusca.Recordset("Jueves") & "," & AdoBusca.Recordset("Viernes") & "," & AdoBusca.Recordset("Sabado") & "," & AdoBusca.Recordset("Domingo") & "," & AdoBusca.Recordset("TotalUnidades") & "," & AdoBusca.Recordset("SalarioPieza") & "," & AdoBusca.Recordset("Unidad")
                   Case "Reporte Percepsiones"
                      Cadena = AdoBusca.Recordset("SalarioBasico") & "," & AdoBusca.Recordset("HTrabajada") & "," & AdoBusca.Recordset("HE") & "," & AdoBusca.Recordset("Produccion") & "," & AdoBusca.Recordset("HorasExtra") & "," & AdoBusca.Recordset("Puntualidad") & "," & AdoBusca.Recordset("Vacaciones") & "," & AdoBusca.Recordset("SeptimoDia") & "," & AdoBusca.Recordset("IncentivosProduccion") & "," & AdoBusca.Recordset("Antiguedad") & "," & AdoBusca.Recordset("OtrosIngresos") & "," & AdoBusca.Recordset("TotalDevengado") & "," & AdoBusca.Recordset("Deducciones") & "," & AdoBusca.Recordset("Prestamo") & "," & AdoBusca.Recordset("MontoInss") & "," & AdoBusca.Recordset("MontoIr") & "," & AdoBusca.Recordset("DiasDescuento") & "," & AdoBusca.Recordset("Adelantos") & "," & AdoBusca.Recordset("TotalDeduccines") & "," & AdoBusca.Recordset("Neto") & "," & AdoBusca.Recordset("InssPatronal") & "," & AdoBusca.Recordset("Inatec") & "," & AdoBusca.Recordset("NumNomina")

                 End Select

                    Print #1, Cadena
                                    
                    
                    
                  AdoBusca.Recordset.MoveNext
                  j = j + 1
                  .Value = j
                  Me.Caption = "Procesando:  " & j & " de " & Maximo & " Registros "
                  DoEvents
                  Cadena = ""
                  Loop
                  End With
                  
                 Close #1

                MsgBox "La Exportacion, fue Creada con Exito", vbExclamation, "Sistema de Enlace"
                Me.Barra.Visible = False
  End If
Exit Sub
TipoErrs:
  MsgBox Err.Description

End Sub

Private Sub CmdExportar_Click()
Dim SQlReportes As String
Me.CommonDialog1.ShowSave
Directorio = ""
Directorio = Me.CommonDialog1.FileName + ".xls"

 Exportar = True
 
 Dim FechaIni As Date, FechaFin As Date
Select Case CmbReportes.Text

Case "Exportar Nomina sin Formato"

     

        SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
        SQlReportes = SQlReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
        SQlReportes = SQlReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
        SQlReportes = SQlReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.FechaNominaINI, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
        SQlReportes = SQlReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
        SQlReportes = SQlReportes & "                      Cargo.Cargo, DetalleNomina.BonoProduccion, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
        SQlReportes = SQlReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
        SQlReportes = SQlReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
        SQlReportes = SQlReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
        SQlReportes = SQlReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
        SQlReportes = SQlReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
        SQlReportes = SQlReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
        SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion AS TotalDevengado," & vbLf
        SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
        SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
        SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion)" & vbLf
        SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
        SQlReportes = SQlReportes & "                      DetalleNomina.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion,Empleado.CodEmpleado1" & vbLf
        SQlReportes = SQlReportes & " FROM         Nomina INNER JOIN" & vbLf
        SQlReportes = SQlReportes & "                       Grupo INNER JOIN" & vbLf
        SQlReportes = SQlReportes & "                       Cargo INNER JOIN" & vbLf
        SQlReportes = SQlReportes & "                       TipoNomina INNER JOIN" & vbLf
        SQlReportes = SQlReportes & "                       Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
        SQlReportes = SQlReportes & "                       DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
        SQlReportes = SQlReportes & "                       TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina" & vbLf
        SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ") AND((dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
        SQlReportes = SQlReportes & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia)" & vbLf
        SQlReportes = SQlReportes & "                      - (dbo.DetalleNomina.Prestamo + dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.MontoIR + dbo.DetalleNomina.Deducciones) <> 0)" & vbLf
        SQlReportes = SQlReportes & " ORDER BY Empleado.CodGrupo, Empleado.CodEmpleado1" & vbLf

   Me.AdoConsulta.RecordSource = SQlReportes
   Me.AdoConsulta.Refresh

   Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
    'Call Formato_Excel(8, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
V = 4
H = 0
i = 1
           objExcel.ActiveSheet.Cells(1, 1) = "METRO GARMENTS S.A."
           objExcel.ActiveSheet.Cells(2, 1) = "SEMANA DEL  " & Format(AdoConsulta.Recordset("FechaNominaINI"), "Long Date") & " AL " & Format(AdoConsulta.Recordset("FechaNomina"), "Long Date")
           
           objExcel.ActiveSheet.Cells(3, 1) = "Codigo"  'A
           objExcel.ActiveSheet.Cells(3, 2) = "Nombre"  'B
            objExcel.ActiveSheet.Cells(3, 3) = "Cargo"  'C
            objExcel.ActiveSheet.Cells(3, 4) = "Tarifa Horaria"  'D
            objExcel.ActiveSheet.Cells(3, 5) = "Horas Trabajadas" 'E
            objExcel.ActiveSheet.Cells(3, 6) = "Salario Basico" 'F
            objExcel.ActiveSheet.Cells(3, 7) = "Monto Destajo" 'G
            objExcel.ActiveSheet.Cells(3, 8) = "Septimo Dia" 'H
            objExcel.ActiveSheet.Cells(3, 9) = "Num Horas" 'I
            objExcel.ActiveSheet.Cells(3, 10) = "Horas Extra" 'J
            objExcel.ActiveSheet.Cells(3, 11) = "Incentivo Produccion" 'K
            objExcel.ActiveSheet.Cells(3, 12) = "Otros Ingresos" 'L
            objExcel.ActiveSheet.Cells(3, 13) = "Total Devengado" 'M
            objExcel.ActiveSheet.Cells(3, 14) = "INSS" 'N
            objExcel.ActiveSheet.Cells(3, 15) = "IR" 'O
            objExcel.ActiveSheet.Cells(3, 16) = "Total Deduc" 'P
            objExcel.ActiveSheet.Cells(3, 17) = "NETO PAGAR" 'Q
            
            objExcel.ActiveSheet.Columns("A").ColumnWidth = 7
            objExcel.ActiveSheet.Columns("B").ColumnWidth = 7
            objExcel.ActiveSheet.Columns("C").ColumnWidth = 6
            objExcel.ActiveSheet.Columns("D").ColumnWidth = 14
            objExcel.ActiveSheet.Columns("E").ColumnWidth = 16
            objExcel.ActiveSheet.Columns("F").ColumnWidth = 14
            objExcel.ActiveSheet.Columns("G").ColumnWidth = 12
            objExcel.ActiveSheet.Columns("H").ColumnWidth = 10
            objExcel.ActiveSheet.Columns("I").ColumnWidth = 11
            objExcel.ActiveSheet.Columns("J").ColumnWidth = 20
            objExcel.ActiveSheet.Columns("K").ColumnWidth = 14
            objExcel.ActiveSheet.Columns("L").ColumnWidth = 15
            objExcel.ActiveSheet.Columns("M").ColumnWidth = 14
            objExcel.ActiveSheet.Columns("N").ColumnWidth = 5
            objExcel.ActiveSheet.Columns("O").ColumnWidth = 3
            objExcel.ActiveSheet.Columns("P").ColumnWidth = 12
            objExcel.ActiveSheet.Columns("Q").ColumnWidth = 10

            
            
     Do While Not Me.AdoConsulta.Recordset.EOF 'est
            
            objExcel.ActiveSheet.Cells(V, H + 1) = Me.AdoConsulta.Recordset("CodEmpleado1")
            objExcel.ActiveSheet.Cells(V, H + 2) = Me.AdoConsulta.Recordset("Nombres")
            objExcel.ActiveSheet.Cells(V, H + 3) = Me.AdoConsulta.Recordset("Cargo")
            objExcel.ActiveSheet.Cells(V, H + 4) = Me.AdoConsulta.Recordset("TarifaHoraria")
            objExcel.ActiveSheet.Cells(V, H + 5) = Me.AdoConsulta.Recordset("HTrabajada")
            objExcel.ActiveSheet.Cells(V, H + 6) = Me.AdoConsulta.Recordset("SalarioBasico")
            objExcel.ActiveSheet.Cells(V, H + 7) = Me.AdoConsulta.Recordset("Destajo")
            objExcel.ActiveSheet.Cells(V, H + 8) = Me.AdoConsulta.Recordset("SeptimoDia")
            objExcel.ActiveSheet.Cells(V, H + 9) = Me.AdoConsulta.Recordset("HE")
            objExcel.ActiveSheet.Cells(V, H + 10) = Me.AdoConsulta.Recordset("HorasExtras")
            objExcel.ActiveSheet.Cells(V, H + 11) = Me.AdoConsulta.Recordset("IncetivoProduccion")
            objExcel.ActiveSheet.Cells(V, H + 12) = Me.AdoConsulta.Recordset("OtrosIngresos")
            objExcel.ActiveSheet.Cells(V, H + 13) = Me.AdoConsulta.Recordset("TotalDevengado")
            objExcel.ActiveSheet.Cells(V, H + 14) = Me.AdoConsulta.Recordset("MontoINSS")
            objExcel.ActiveSheet.Cells(V, H + 15) = Me.AdoConsulta.Recordset("MontoIR")
            objExcel.ActiveSheet.Cells(V, H + 16) = Me.AdoConsulta.Recordset("TotalDeducir")
            objExcel.ActiveSheet.Cells(V, H + 17) = Me.AdoConsulta.Recordset("NetoPagar")


            V = V + 1
            i = i + 1
            
        Me.AdoConsulta.Recordset.MoveNext
     Loop


Case "Nomina x Cargo"
       ArepNominaDpto.AdoNomina.ConnectionString = ConexionReporte
       ArepNominaDpto.LblFecha.Caption = Format(Now, "dd/mm/yyyy ")
'       ArepNominaDpto.LblDesde = FrmCalcularNomina.lblFecha1.Caption
'       ArepNominaDpto.LblHasta = FrmCalcularNomina.lblFecha2.Caption
       
    '///////////////////////////INTRUCCION SQL SERVER
    SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
    SQlReportes = SQlReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
    SQlReportes = SQlReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
    SQlReportes = SQlReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
    SQlReportes = SQlReportes & "                      Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
    SQlReportes = SQlReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
    SQlReportes = SQlReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
    SQlReportes = SQlReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
    SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion AS TotalDevengado," & vbLf
    SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
    SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
    SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion)" & vbLf
    SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
    SQlReportes = SQlReportes & "                      Empleado.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1,DetalleNomina.produjo" & vbLf
    SQlReportes = SQlReportes & " FROM         Nomina INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       Grupo INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       Cargo INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       TipoNomina INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
    SQlReportes = SQlReportes & "                       TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina" & vbLf
    SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ")AND " & vbLf
    SQlReportes = SQlReportes & " (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
    SQlReportes = SQlReportes & "DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion <> 0)" & vbLf
    SQlReportes = SQlReportes & " ORDER BY Cargo.Cargo, DetalleNomina.Produjo DESC, Empleado.CodEmpleado1" & vbLf


       ArepNominaDpto.AdoNomina.Source = SQlReportes
       ArepNominaDpto.LblTitulo.Caption = Titulo
       ArepNominaDpto.LblSubtitulo.Caption = SubTitulo
       ArepNominaDpto.ImgLogo.Picture = LoadPicture(RutaLogo)
       ArepNominaDpto.Show 1
       
Case "Nomina x Cargo Resumen"
       ArepNominaDptoResumen.AdoNomina.ConnectionString = ConexionReporte
       ArepNominaDptoResumen.LblFecha.Caption = Format(Now, "dd/mm/yyyy ")
'       ArepNominaDpto.LblDesde = FrmCalcularNomina.lblFecha1.Caption
'       ArepNominaDpto.LblHasta = FrmCalcularNomina.lblFecha2.Caption
       
    '///////////////////////////INTRUCCION SQL SERVER
    SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
    SQlReportes = SQlReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
    SQlReportes = SQlReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
    SQlReportes = SQlReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
    SQlReportes = SQlReportes & "                      Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
    SQlReportes = SQlReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
    SQlReportes = SQlReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
    SQlReportes = SQlReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
    SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion AS TotalDevengado," & vbLf
    SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
    SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
    SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion)" & vbLf
    SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
    SQlReportes = SQlReportes & "                      Empleado.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1,DetalleNomina.produjo" & vbLf
    SQlReportes = SQlReportes & " FROM         Nomina INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       Grupo INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       Cargo INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       TipoNomina INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
    SQlReportes = SQlReportes & "                       TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina" & vbLf
    SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ")AND " & vbLf
    SQlReportes = SQlReportes & " (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
    SQlReportes = SQlReportes & "DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion <> 0)" & vbLf
    SQlReportes = SQlReportes & " ORDER BY Cargo.Cargo, DetalleNomina.Produjo DESC, Empleado.CodEmpleado1" & vbLf


       ArepNominaDptoResumen.AdoNomina.Source = SQlReportes
       ArepNominaDptoResumen.LblTitulo.Caption = Titulo
       ArepNominaDptoResumen.LblSubtitulo.Caption = SubTitulo
       ArepNominaDptoResumen.ImgLogo.Picture = LoadPicture(RutaLogo)
       ArepNominaDptoResumen.Show 1
  End Select
End Sub

Private Sub cmdImprimir_Click()
Dim FechaIni As Date, FechaFin As Date, CodDepartamentos As String, DescripcionDpto As String
Dim SqlString As String, FechaNomina As String, Respuesta As Double
Dim rpt As Object
Dim fPreview As New FrmPreview

Select Case CmbReportes.Text
Case "Colilla para Cada Departamento"
     Me.AdoDepartamentos.Refresh
     Do While Not Me.AdoDepartamentos.Recordset.EOF
     
          CodDepartamento = Me.AdoDepartamentos.Recordset("CodDepartamento")
          DescripcionDpto = Me.AdoDepartamentos.Recordset("Departamento")
     
          SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal,  Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo,  Cargo.Cargo, DetalleNomina.BonoProduccion, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones,  " & _
                      "DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal,  DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo,  Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE, DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion  AS TotalDevengado, DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir,  " & _
                      "(DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +  DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar, DetalleNomina.TarifaHoraria, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1, departamento.departamento , departamento.CodDepartamento,Nomina.FechaNominaINI  " & _
                      "FROM  Nomina INNER JOIN  Grupo INNER JOIN  Cargo INNER JOIN  TipoNomina INNER JOIN  Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN  DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON  TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN  Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento " & _
                      "WHERE     (Nomina.NumNomina = " & NumNomina & ") AND ((DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) <> 0) AND (Departamento.CodDepartamento = '" & CodDepartamento & "') " & _
                      "ORDER BY Empleado.CodGrupo, Empleado.Apellido1, Empleado.Apellido2 "
                      
'                      SQlReportes = "SELECT Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo, Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas, DetalleNomina.BonoProduccion, DetalleNomina.Prestamo, " & _
'                                    "DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre,DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres,DetalleNomina.HE, DetalleNomina.SalarioBasico DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasTurno + DetalleNomina.Antiguedad AS TotalDevengado, DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir,  " & _
'                                    "(DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasTurno + DetalleNomina.Antiguedad) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar, DetalleNomina.TarifaHoraria, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1, Empleado.NumeroInss, Empleado.NumCedula, Departamento.Departamento, DetalleNomina.HorasTurno, DetalleNomina.HTurno, DetalleNomina.Antiguedad, DetalleNomina.AñoAntiguedad , departamento.CodDepartamento, Nomina.FechaNominaINI FROM  Nomina INNER JOIN Grupo INNER JOIN Cargo INNER JOIN  TipoNomina INNER JOIN  " & _
'                                    "Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento  " & _
'                                    "WHERE (Nomina.NumNomina = " & NumNomina & ") AND (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasTurno + DetalleNomina.Antiguedad <> 0) AND (Departamento.CodDepartamento = '" & CodDepartamento & "') ORDER BY Nomina.NumNomina, Empleado.Nombre1"
                      
                    
                      Me.AdoBusca.RecordSource = SQlReportes
                      Me.AdoBusca.Refresh
                      If Not Me.AdoBusca.Recordset.EOF Then
                      
                       Respuesta = MsgBox("Departamento Procesado " & DescripcionDpto, vbYesNo, "Zeus Nominas")
                       If Respuesta = 7 Then
                        Exit Sub
                       End If
                        
                      
'                        ArepColillaProduccionLegal.AdoColillas.Source = SQlReportes
'                        ArepColillaProduccionLegal.LblPeriodo.Caption = Me.AdoBusca.Recordset("FechaNominaINI") & " al " & Me.AdoBusca.Recordset("FechaNomina")
'                        ArepColillaProduccionLegal.LblTitulo.Caption = Titulo
'                        ArepColillaProduccionLegal.AdoColillas.ConnectionString = ConexionReporte

                        ArepColillasPago.AdoColillas.Source = SQlReportes
                        PeriodoReporte = "Desde " & Me.AdoBusca.Recordset("FechaNominaINI") & " Hasta " & Me.AdoBusca.Recordset("FechaNomina")
                        ArepColillasPago.LblTitulo.Caption = Titulo
                        ArepColillasPago.AdoColillas.ConnectionString = ConexionReporte
                        ArepColillasPago.LblPeriodo.Caption = "   Colilla de Pago # " & NumNomina & ", Corespondiente del " & FrmCalcularNomina.LblFecha1.Caption & " al " & FrmCalcularNomina.LblFecha2.Caption

     Set rpt = New ArepColillasPago
     rpt.AdoColillas.ConnectionString = ConexionReporte
     rpt.AdoColillas.Source = SQlReportes
     fPreview.RunReport rpt

     fPreview.Show 1
'           fPreview.arv.ReportSource = ArepColillaProduccionLegal
'           fPreview.Show 1
                        
'                        ArepNominaProduccionLegal.LblDesde.Caption = Me.AdoBusca.Recordset("FechaNominaINI")
'                        ArepNominaProduccionLegal.LblHasta.Caption = Me.AdoBusca.Recordset("FechaNomina")
'                        ArepNominaProduccionLegal.AdoNomina.ConnectionString = ConexionReporte
'                        ArepNominaProduccionLegal.LblFecha.Caption = Format(Now, "dd/mm/yyyy ")
'                        ArepNominaProduccionLegal.AdoNomina.Source = SQlReportes
'                        ArepNominaProduccionLegal.lbltitulo.Caption = Titulo
'                        ArepNominaProduccionLegal.LblSubtitulo.Caption = Subtitulo
'                        ArepNominaProduccionLegal.ImgLogo.Picture = LoadPicture(RutaLogo)
'                        ArepNominaProduccionLegal.Show 1
                      
                      End If
         
         Me.AdoDepartamentos.Recordset.MoveNext
      Loop




Case "Nomina para Cada Departamento"

     Me.AdoDepartamentos.Refresh
     Do While Not Me.AdoDepartamentos.Recordset.EOF
     
          CodDepartamento = Me.AdoDepartamentos.Recordset("CodDepartamento")
          DescripcionDpto = Me.AdoDepartamentos.Recordset("Departamento")
     
          SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal,  Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo,  Cargo.Cargo, DetalleNomina.BonoProduccion, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion As Incentivos, DetalleNomina.Deducciones,  " & _
                      "DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal,  DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo,  Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE, DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion  AS TotalDevengado, DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir,  " & _
                      "(DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +  DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar, DetalleNomina.TarifaHoraria, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1, departamento.departamento , departamento.CodDepartamento,Nomina.FechaNominaINI  " & _
                      "FROM  Nomina INNER JOIN  Grupo INNER JOIN  Cargo INNER JOIN  TipoNomina INNER JOIN  Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN  DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON  TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN  Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento " & _
                      "WHERE     (Nomina.NumNomina = " & NumNomina & ") AND ((DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) <> 0) AND (Departamento.CodDepartamento = '" & CodDepartamento & "') " & _
                      "ORDER BY Empleado.CodGrupo, Empleado.CodEmpleado1 "
                      
                      Me.AdoBusca.RecordSource = SQlReportes
                      Me.AdoBusca.Refresh
                      If Not Me.AdoBusca.Recordset.EOF Then
                      
                      Respuesta = MsgBox("Departamento Procesado " & DescripcionDpto, vbYesNo, "Zeus Nominas")
                       If Respuesta = 7 Then
                        Exit Sub
                       End If
                        
                        ArepNominaProduccionLegal.LblDesde.Caption = Me.AdoBusca.Recordset("FechaNominaINI")
                        ArepNominaProduccionLegal.LblHasta.Caption = Me.AdoBusca.Recordset("FechaNomina")
                        ArepNominaProduccionLegal.AdoNomina.ConnectionString = ConexionReporte
                        ArepNominaProduccionLegal.LblFecha.Caption = Format(Now, "dd/mm/yyyy ")
                        ArepNominaProduccionLegal.AdoNomina.Source = SQlReportes
                        ArepNominaProduccionLegal.LblTitulo.Caption = Titulo
                        ArepNominaProduccionLegal.LblSubtitulo.Caption = SubTitulo
                        ArepNominaProduccionLegal.ImgLogo.Picture = LoadPicture(RutaLogo)
                        ArepNominaProduccionLegal.Show 1
'           fPreview.arv.ReportSource = ArepNominaProduccionLegal
'           fPreview.Show 1

'     Set rpt = New ArepColillaProduccionLegal
'     rpt.DataControl1.ConnectionString = ConexionReporte
'     rpt.DataControl1.Source = SQlReportes
'     fPreview.RunReport rpt
'
'     fPreview.Show 1
                      
                      End If
         
         Me.AdoDepartamentos.Recordset.MoveNext
      Loop

Case "Nomina x Departamento"

       SQlReportes = " SELECT  TOP 100 PERCENT Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico,Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo, Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones,DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos+ DetalleNomina.IncetivoProduccion As Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones, " & _
                      "DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones,DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13,DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE,DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas AS TotalDevengado,DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir,(DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + " & _
                      "DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+DetalleNomina.IncetivoProduccion) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar,DetalleNomina.TarifaHoraria, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1,departamento.departamento " & _
                      "FROM Nomina INNER JOIN Grupo INNER JOIN Cargo INNER JOIN  TipoNomina INNER JOIN Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN  DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento " & _
                      " WHERE     (Nomina.NumNomina = " & NumNomina & ")AND " & _
                      "(DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + " & _
                      "DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion <> 0) " & _
                      " ORDER BY Departamento.Departamento, DetalleNomina.Produjo DESC, Empleado.CodEmpleado1 "

       ArepNominaDepartamento.AdoNomina.ConnectionString = ConexionReporte
       ArepNominaDepartamento.LblFecha.Caption = Format(Now, "dd/mm/yyyy ")
       ArepNominaDepartamento.AdoNomina.Source = SQlReportes
       ArepNominaDepartamento.LblTitulo.Caption = Titulo
       ArepNominaDepartamento.LblSubtitulo.Caption = SubTitulo
       ArepNominaDepartamento.ImgLogo.Picture = LoadPicture(RutaLogo)
'       ArepNominaDepartamento.Show 1
           fPreview.arv.ReportSource = ArepNominaDepartamento
           fPreview.Show 1

Case "Nomina x Cargo"
       ArepNominaDpto.AdoNomina.ConnectionString = ConexionReporte
       ArepNominaDpto.LblFecha.Caption = Format(Now, "dd/mm/yyyy ")
'       ArepNominaDpto.LblDesde = FrmCalcularNomina.lblFecha1.Caption
'       ArepNominaDpto.LblHasta = FrmCalcularNomina.lblFecha2.Caption
       
    '///////////////////////////INTRUCCION SQL SERVER
    SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
    SQlReportes = SQlReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
    SQlReportes = SQlReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
    SQlReportes = SQlReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
    SQlReportes = SQlReportes & "                      Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
    SQlReportes = SQlReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
    SQlReportes = SQlReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
    SQlReportes = SQlReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
    SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion AS TotalDevengado," & vbLf
    SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
    SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
    SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion)" & vbLf
    SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
    SQlReportes = SQlReportes & "                      Empleado.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1,DetalleNomina.produjo" & vbLf
    SQlReportes = SQlReportes & " FROM         Nomina INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       Grupo INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       Cargo INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       TipoNomina INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
    SQlReportes = SQlReportes & "                       TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina" & vbLf
    SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ")AND " & vbLf
    SQlReportes = SQlReportes & " (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
    SQlReportes = SQlReportes & "DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion <> 0)" & vbLf
    SQlReportes = SQlReportes & " ORDER BY Cargo.Cargo, DetalleNomina.Produjo DESC, Empleado.CodEmpleado1" & vbLf


       ArepNominaDpto.AdoNomina.Source = SQlReportes
       ArepNominaDpto.LblTitulo.Caption = Titulo
       ArepNominaDpto.LblSubtitulo.Caption = SubTitulo
       ArepNominaDpto.ImgLogo.Picture = LoadPicture(RutaLogo)
'       ArepNominaDpto.Show 1
           fPreview.arv.ReportSource = ArepNominaDpto
           fPreview.Show 1
       
Case "Nomina x Cargo Resumen"
       ArepNominaDptoResumen.AdoNomina.ConnectionString = ConexionReporte
       ArepNominaDptoResumen.LblFecha.Caption = Format(Now, "dd/mm/yyyy ")
'       ArepNominaDpto.LblDesde = FrmCalcularNomina.lblFecha1.Caption
'       ArepNominaDpto.LblHasta = FrmCalcularNomina.lblFecha2.Caption
       
    '///////////////////////////INTRUCCION SQL SERVER
    SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
    SQlReportes = SQlReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
    SQlReportes = SQlReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
    SQlReportes = SQlReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
    SQlReportes = SQlReportes & "                      Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
    SQlReportes = SQlReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
    SQlReportes = SQlReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
    SQlReportes = SQlReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
    SQlReportes = SQlReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
    SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion AS TotalDevengado," & vbLf
    SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
    SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
    SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion)" & vbLf
    SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
    SQlReportes = SQlReportes & "                      Empleado.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1,DetalleNomina.produjo" & vbLf
    SQlReportes = SQlReportes & " FROM         Nomina INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       Grupo INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       Cargo INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       TipoNomina INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
    SQlReportes = SQlReportes & "                       DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
    SQlReportes = SQlReportes & "                       TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina" & vbLf
    SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ")AND " & vbLf
    SQlReportes = SQlReportes & " (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
    SQlReportes = SQlReportes & "DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion <> 0)" & vbLf
    SQlReportes = SQlReportes & " ORDER BY Cargo.Cargo, DetalleNomina.Produjo DESC, Empleado.CodEmpleado1" & vbLf


       ArepNominaDptoResumen.AdoNomina.Source = SQlReportes
       ArepNominaDptoResumen.LblTitulo.Caption = Titulo
       ArepNominaDptoResumen.LblSubtitulo.Caption = SubTitulo
       ArepNominaDptoResumen.ImgLogo.Picture = LoadPicture(RutaLogo)
'       ArepNominaDptoResumen.Show 1
           fPreview.arv.ReportSource = ArepNominaDptoResumen
           fPreview.Show 1



Case "Reporte Percepsiones"


SQlReportes = "SELECT     SUM(SalarioBasico) AS SalarioBasico, SUM(Destajo) AS Produccion, SUM(HorasExtras) AS HorasExtra, SUM(Comisiones) AS Puntualidad, " & _
                      "SUM(VacacionesPagadas) AS Vacaciones, SUM(SeptimoDia) AS SeptimoDia, SUM(IncetivoProduccion) AS IncentivosProduccion, SUM(Incentivos) " & _
                      "AS Antiguedad, SUM(OtrosIngresos) AS OtrosIngresos, SUM(SalarioBasico) + SUM(Destajo) + SUM(HorasExtras) + SUM(Comisiones) " & _
                      "+ SUM(SeptimoDia) + SUM(IncetivoProduccion + Incentivos) + SUM(OtrosIngresos) + SUM(VacacionesPagadas) + SUM(BonoProduccion) " & _
                      "AS TotalDevengado, SUM(Deducciones) AS Deducciones, SUM(Prestamo) AS Prestamo, SUM(MontoINSS) AS MontoInss, SUM(MontoIR) AS MontoIr, " & _
                      "SUM(DiasDescuento) AS DiasDescuento, SUM(Adelantos) AS Adelantos, SUM(Deducciones) + SUM(Prestamo) + SUM(MontoINSS) + SUM(MontoIR) " & _
                      "+ SUM(DiasDescuento) + SUM(Adelantos) AS TotalDeduccines, (SUM(SalarioBasico) + SUM(Destajo) + SUM(HorasExtras) + SUM(Comisiones) " & _
                      "+ SUM(SeptimoDia) + SUM(IncetivoProduccion + Incentivos) + SUM(OtrosIngresos) + SUM(BonoProduccion) + SUM(VacacionesPagadas)) " & _
                      "- (SUM(Deducciones) + SUM(Prestamo) + SUM(MontoINSS) + SUM(MontoIR) + SUM(DiasDescuento) + SUM(Adelantos)) AS Neto, SUM(INSSPatronal) " & _
                      "AS InssPatronal, SUM(IRPatronal) AS IrPatronal, SUM(INATEC) AS Inatec, NumNomina, SUM(HE) AS HE, SUM(HTrabajada) AS HTrabajada, " & _
                      "SUM(INSSPatronal) + SUM(INATEC) AS TotalObligaciones, SUM(BonoProduccion) AS BonoProduccion, SUM(AjusteINSS) AS AjusteINSS " & _
             "From DetalleNomina " & _
             "GROUP BY NumNomina " & _
             "Having (NumNomina = " & NumNomina & ") "

 
 SqlString = "SELECT  * From Nomina Where (NumNomina = " & NumNomina & ")"
                      
 Me.AdoBusca.RecordSource = SqlString
 Me.AdoBusca.Refresh
 If Not FrmNominaActiva.AdoBusca.Recordset.EOF Then
  FechaNomina = "Desde: " & Format(Me.AdoBusca.Recordset("FechaNominaINI"), "Long Date") & "        Hasta: " & Format(Me.AdoBusca.Recordset("FechaNomina"), "Long Date")
 
 End If


 MDIPrimero.DtaEmpresa.Refresh
 If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
   FormatoNomina = MDIPrimero.DtaEmpresa.Recordset("FormatoNomina")
 End If
 
 Select Case FormatoNomina
   Case "Nomina Comercial2"
  
  
  Case "Nomina Comercial"

  Case "Nomina Produccion"

     ArepPersepciones.DataControl1.ConnectionString = ConexionReporte
     ArepPersepciones.LblTitulo.Caption = Titulo
     ArepPersepciones.LblSubtitulo.Caption = SubTitulo
     ArepPersepciones.LblTitulo3.Caption = "Reporte de Persepciones de la Nomina No " & NumNomina
     ArepPersepciones.ImgLogo.Picture = LoadPicture(RutaLogo)
     ArepPersepciones.LblFechaNomina.Caption = FechaNomina
     
     ArepPersepciones.DataControl1.Source = SQlReportes
     
     Me.Text1.Text = ArepPersepciones.DataControl1.Source
'     ArepPersepciones.Show 1

           fPreview.arv.ReportSource = ArepPersepciones
           fPreview.Show 1
     
     
   Case "Nomina Bono Produccion"

     ArepPersepcionesBonoProduccion.DataControl1.ConnectionString = ConexionReporte
     ArepPersepcionesBonoProduccion.LblTitulo.Caption = Titulo
     ArepPersepcionesBonoProduccion.LblSubtitulo.Caption = SubTitulo
     ArepPersepcionesBonoProduccion.LblTitulo3.Caption = "Reporte de Persepciones de la Nomina No " & NumNomina
     If Dir(RutaLogo) <> "" Then
      ArepPersepcionesBonoProduccion.ImgLogo.Picture = LoadPicture(RutaLogo)
     End If
     
     ArepPersepcionesBonoProduccion.LblFechaNomina.Caption = FechaNomina
     ArepPersepcionesBonoProduccion.DataControl1.Source = SQlReportes
     
     Me.Text1.Text = ArepPersepcionesBonoProduccion.DataControl1.Source
'     ArepPersepcionesBonoProduccion.Show 1
           fPreview.arv.ReportSource = ArepPersepcionesBonoProduccion
           fPreview.Show 1
     
   
  End Select




Case "Reporte x Produccion"
'      Fechaini = Format(FrmCalcularNomina.LblFecha1.Caption.Caption, "dd/mm/yyyy")
'      Fechafin = Format(FrmCalcularNomina.LblFecha2.Caption, "dd/mm/yyyy")
'     Numero = Me.TxtNNomina.Text

'      Fechaini = Format(FrmCalcularNomina.LblFecha1.Caption.Caption, "dd/mm/yyyy")
'      Fechafin = Format(FrmCalcularNomina.LblFecha2.Caption, "dd/mm/yyyy")
'     Numero = Me.TxtNNomina.Text
     ArepProduccion.DataControl1.ConnectionString = ConexionReporte
     ArepProduccion.DataControl1.UserID = "metro"
     ArepProduccion.DataControl1.Password = "metro"
     ArepProduccion.LblTitulo.Caption = Titulo
     ArepProduccion.LblSubtitulo.Caption = SubTitulo
     ArepProduccion.LblTitulo3.Caption = "Reporte de Produccion de la Nomina No " & NumNomina
     ArepProduccion.ImgLogo.Picture = LoadPicture(RutaLogo)
     ArepProduccion.DataControl1.Source = "SELECT DetalleNomina.NumNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HE, " & _
                      "DetalleNomina.DD, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, " & _
                      "DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, " & _
                      "DetalleNomina.Vacaciones, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.INATEC, DetalleNomina.Mes13, " & _
                      "DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.TotalSubsidio, DetalleNomina.VacacionesPagadas, " & _
                      "DetalleNomina.DiasVacaciones, DetalleNomina.AdelantosVacaciones, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, " & _
                      "DetalleNomina.IncetivoProduccion, DetalleNomina.TarifaHoraria, Nomina.FechaNomina, Nomina.FechaNominaINI, Empleado.CodEmpleado1, " & _
                      "DetalleProduccion.CodReferencia, DetalleProduccion.CodReferencia1 , DetalleProduccion.CodProceso, DetalleProduccion.Ref, DetalleProduccion.Lunes, DetalleProduccion.Martes, " & _
                      "DetalleProduccion.Miercoles, DetalleProduccion.Jueves, DetalleProduccion.Viernes, DetalleProduccion.Sabado, DetalleProduccion.Domingo, " & _
                      "DetalleProduccion.TotalUnidades, DetalleProduccion.SalarioPieza, DetalleProduccion.Precio, DetalleProduccion.Unidad, " & _
                      "DetalleProduccion.Pagado FROM DetalleNomina INNER JOIN " & _
                      "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN " & _
                      "Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado INNER JOIN " & _
                      "DetalleProduccion ON Empleado.CodEmpleado = DetalleProduccion.CodEmpleado " & _
                      "WHERE (DetalleNomina.NumNomina = " & NumNomina & ")AND (DetalleProduccion.NumNomina = " & NumNomina & ") ORDER BY Empleado.CodEmpleado1, DetalleProduccion.CodReferencia, DetalleProduccion.CodProceso"
                       ' ORDER BY Empleado.CodEmpleado1"
     
'     ArepProduccion.Show 1
           fPreview.arv.ReportSource = ArepProduccion
           fPreview.Show 1

 Case "Listado Horas Extra"
       ArepHorasExtras.DataControl1.ConnectionString = ConexionReporte
       ArepHorasExtras.DataControl1.Source = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, HorasExtras.CantHoras, HorasExtras.NumNomina FROM Empleado INNER JOIN HorasExtras ON Empleado.CodEmpleado = HorasExtras.CodEmpleado Where (((HorasExtras.CantHoras) <> 0) And ((HorasExtras.NumNomina) = " & NumNomina & ")) ORDER BY Empleado.CodEmpleado1"
       ArepHorasExtras.LblTitulo.Caption = Titulo
       ArepHorasExtras.LblSubtitulo.Caption = SubTitulo
       ArepHorasExtras.ImgLogo.Picture = LoadPicture(RutaLogo)
       ArepHorasExtras.Show 1
'           fPreview.arv.ReportSource = ArepHorasExtras
'           fPreview.Show 1
       
  Case "Nomina x Codigo Empleado"
'       CodTipoNomina = FrmCalcularNomina.DtaTipoNomina.Recordset("CodTipoNomina")
'       NumNomina = FrmCalcularNomina.DtaNomina.Recordset("NumNomina")

       ArepNomina.AdoNomina.ConnectionString = ConexionReporte
       ArepNomina.LblFecha.Caption = Format(Now, "dd/mm/yyyy ")
'       ArepNomina.LblDesde = FrmCalcularNomina.lblFecha1.Caption
'       ArepNomina.LblHasta = FrmCalcularNomina.lblFecha2.Caption
'
 '//////////PARA ACCESS 97 SE UTILIZA EN LAS CONSULTAS "].["  POR    "].["
'NumNomina = DtaNomina.Recordset("NumNomina")
'SQLReportes = "SELECT  Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre," & vbLf
'SQLReportes = SQLReportes & "Cargo.CodCargo , Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas,DetalleNomina.Prestamo, DetalleNomina.MontoInss, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal," & vbLf
'SQLReportes = SQLReportes & "DetalleNomina.IRPatronal, DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo,Empleado.DescripOtrIngre, Grupo.Grupo, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, DetalleNomina.HE,[DetalleNomina].[SalarioBasico]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]+[DetalleNomina].[HorasExtras]+[DetalleNomina].[OtrosIngresos]+[DetalleNomina].[Destajo]+ [DetalleNomina].[VacacionesPagadas] AS TotalDevengado, [DetalleNomina].[Prestamo]+[DetalleNomina].[MontoINSS]+[DetalleNomina].[MontoIR]+[DetalleNomina].[Deducciones] AS TotalDeducir, [TotalDevengado]-[TotalDeducir] AS NetoPagar" & vbLf
'SQLReportes = SQLReportes & "FROM Nomina INNER JOIN (Grupo INNER JOIN ((Cargo INNER JOIN (TipoNomina INNER JOIN Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina) ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) ON Grupo.CodGrupo = Empleado.CodGrupo) ON (TipoNomina.CodTipoNomina = Nomina.CodTipoNomina) AND (Nomina.NumNomina = DetalleNomina.NumNomina)" & vbLf
'SQLReportes = SQLReportes & "Where (((Nomina.NumNomina) = " & NumNomina & ")) ORDER BY Nomina.NumNomina, DetalleNomina.CodEmpleado" & vbLf

'///////////////////////////INTRUCCION SQL SERVER
SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
SQlReportes = SQlReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
SQlReportes = SQlReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
SQlReportes = SQlReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
SQlReportes = SQlReportes & "                      Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
SQlReportes = SQlReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion AS TotalDevengado," & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion)" & vbLf
SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
SQlReportes = SQlReportes & "                      Empleado.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1" & vbLf
SQlReportes = SQlReportes & " FROM         Nomina INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       Grupo INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       Cargo INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       TipoNomina INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
SQlReportes = SQlReportes & "                       TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina" & vbLf
SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ")" & vbLf
SQlReportes = SQlReportes & " ORDER BY Nomina.NumNomina, Empleado.CodEmpleado1" & vbLf


       ArepNomina.AdoNomina.Source = SQlReportes
       ArepNomina.LblTitulo.Caption = Titulo
       ArepNomina.LblSubtitulo.Caption = SubTitulo
       ArepNomina.ImgLogo.Picture = LoadPicture(RutaLogo)
'       ArepNomina.Show 1
           fPreview.arv.ReportSource = ArepNomina
           fPreview.Show 1
  
End Select
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
 With Me.AdoBusca
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
 End With
 
  With Me.AdoConsulta
   .ConnectionString = Conexion
 End With
 
 With Me.AdoDepartamentos
   .ConnectionString = Conexion
   .RecordSource = "SELECT  * From departamento"
   .Refresh
 End With
 
 
CmbReportes.AddItem "Listado Horas Extra"
CmbReportes.AddItem "Nomina x Codigo Empleado"
CmbReportes.AddItem "Nomina x Departamento"
CmbReportes.AddItem "Nomina para Cada Departamento"
CmbReportes.AddItem "Colilla para Cada Departamento"
CmbReportes.AddItem "Nomina x Cargo"
CmbReportes.AddItem "Nomina x Cargo Resumen"
CmbReportes.AddItem "Reporte x Produccion"
CmbReportes.AddItem "Reporte Percepsiones"
CmbReportes.AddItem "Exportar Nomina sin Formato"

End Sub
