VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form FrmExportaBac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportacion Planilla al BAC"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkExportaDolares 
      Caption         =   "Exportar Salarios en Dolares"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   360
      Top             =   5400
      Width           =   3135
      _ExtentX        =   5530
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
      Caption         =   "DtaConsulta"
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
   Begin MSAdodcLib.Adodc DtaExporta 
      Height          =   375
      Left            =   240
      Top             =   4680
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "DtaExporta"
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
   Begin vbskfree.Skinner Skinner1 
      Left            =   720
      Top             =   3360
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin MSComCtl2.DTPicker DTFecha 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      Format          =   78643201
      CurrentDate     =   38237
   End
   Begin VB.TextBox TxtCod 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin SmartButtonProject.SmartButton CmdExel 
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ForeColor       =   8388608
      Caption         =   "Exportar Exel"
      Picture         =   "FrmExportaBac.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SmartButtonProject.SmartButton CmdSalir 
      Height          =   855
      Left            =   2400
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      Caption         =   "Salir"
      Picture         =   "FrmExportaBac.frx":0B76
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Dia a Pagar"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmExportaBac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExel_Click()
On Error GoTo TipoErrs
Dim SQlReportes As String, V As Integer, H As Integer, i As Integer
Dim Año As String, MesLetra As String, Neto As String, Dias As String
Dim CanDias As String, QuinLetra As String, Nombres As String, Espacio As String
Dim TotalNomina As Double, Neto1 As Double, Cod As String, NetoT As String, Longitud As Integer
Dim TasaCambio As Double, FechaNomina As Date

Espacio = " "
Select Case Quien
 Case "CalcularNomina"
       '//////////////////////Cargo la Consulta de la Nomina///////////////////////
       NumNomina = FrmCalcularNomina.DtaNomina.Recordset("NumNomina")
'//////////PARA ACCESS 97 SE UTILIZA EN LAS CONSULTAS "].["  POR    "].["
'NumNomina = DtaNomina.Recordset("NumNomina")
'SQLReportes = "SELECT  Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre," & vbLf
'SQLReportes = SQLReportes & "Cargo.CodCargo , Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas,DetalleNomina.Prestamo, DetalleNomina.MontoInss, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal," & vbLf
'SQLReportes = SQLReportes & "DetalleNomina.IRPatronal, DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo,Empleado.DescripOtrIngre, Grupo.Grupo, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, DetalleNomina.HE,[DetalleNomina].[SalarioBasico]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]+[DetalleNomina].[HorasExtras]+[DetalleNomina].[OtrosIngresos]+[DetalleNomina].[Destajo]+ [DetalleNomina].[VacacionesPagadas] AS TotalDevengado, [DetalleNomina].[Prestamo]+[DetalleNomina].[MontoINSS]+[DetalleNomina].[MontoIR]+[DetalleNomina].[Deducciones] AS TotalDeducir, [TotalDevengado]-[TotalDeducir] AS NetoPagar" & vbLf
'SQLReportes = SQLReportes & "FROM Nomina INNER JOIN (Grupo INNER JOIN ((Cargo INNER JOIN (TipoNomina INNER JOIN Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina) ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) ON Grupo.CodGrupo = Empleado.CodGrupo) ON (TipoNomina.CodTipoNomina = Nomina.CodTipoNomina) AND (Nomina.NumNomina = DetalleNomina.NumNomina)" & vbLf
'SQLReportes = SQLReportes & "Where (((Nomina.NumNomina) = " & NumNomina & ")) ORDER BY Nomina.NumNomina, DetalleNomina.CodEmpleado" & vbLf

'///////////////////////////INTRUCCION SQL SERVER
'SQLReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
'SQLReportes = SQLReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
'SQLReportes = SQLReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
'SQLReportes = SQLReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
'SQLReportes = SQLReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
'SQLReportes = SQLReportes & "                      Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
'SQLReportes = SQLReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
'SQLReportes = SQLReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
'SQLReportes = SQLReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
'SQLReportes = SQLReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
'SQLReportes = SQLReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
'SQLReportes = SQLReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
'SQLReportes = SQLReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas AS TotalDevengado," & vbLf
'SQLReportes = SQLReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
'SQLReportes = SQLReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
'SQLReportes = SQLReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas)" & vbLf
'SQLReportes = SQLReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar" & vbLf
'SQLReportes = SQLReportes & " FROM         Nomina INNER JOIN" & vbLf
'SQLReportes = SQLReportes & "                       Grupo INNER JOIN" & vbLf
'SQLReportes = SQLReportes & "                       Cargo INNER JOIN" & vbLf
'SQLReportes = SQLReportes & "                       TipoNomina INNER JOIN" & vbLf
'SQLReportes = SQLReportes & "                       Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
'SQLReportes = SQLReportes & "                       DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
'SQLReportes = SQLReportes & "                       TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina" & vbLf
'SQLReportes = SQLReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ")" & vbLf
'SQLReportes = SQLReportes & " ORDER BY Nomina.NumNomina, DetalleNomina.CodEmpleado" & vbLf

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
SQlReportes = SQlReportes & "                      DetalleNomina.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion,Empleado.CodEmpleado1, Empleado.NumCedula" & vbLf
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
SQlReportes = SQlReportes & " ORDER BY Empleado.CodEmpleado1" & vbLf


       Me.DtaConsulta.RecordSource = SQlReportes
       Me.DtaConsulta.Refresh

       Me.DtaExporta.Refresh
        'Me.'Me.DtaExporta.Recordset.Edit
       Me.DtaExporta.Recordset("CodigoBAC") = val(Me.TxtCod.Text)
       Me.DtaExporta.Recordset.Update

       Mes = Month(Me.DtaConsulta.Recordset("FechaNomina"))
       Año = Year(Me.DtaConsulta.Recordset("FechaNomina"))
       CanDias = Day(Me.DtaConsulta.Recordset("FechaNomina"))
       Dias = Day(Me.DTFecha.Value)
       Cod = Me.TxtCod.Text
       FechaNomina = Me.DtaConsulta.Recordset("FechaNomina")

       ConvertirMes (Mes)
      If CanDias > 15 Then
         QuinLetra = "Segunda Quincena de " & Convertir
      Else
         QuinLetra = "Primera Quincena de" & Convertir
      End If
  
   Case "NominaVacaciones"
      NumNomVaca = Frm13Vaca.TxtNumNomVaca.Text
      '///////////////////////////Cargo la Consulta de Vacaciones////////////////////////////////
      SQlReportes = "SELECT NomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, ([DetalleNomVaca].[SalarioMensual]*([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento])/" & DiasMes & ")-[DetalleNomVaca].[AdelantoVacaciones] AS MontoAPagar, [DetalleNomVaca].[SalarioMensual]*([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento])/" & DiasMes & " AS TotalDevengado, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, ([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento]) AS TotalDescuento " & vbLf
      SQlReportes = SQlReportes & "FROM NomVaca INNER JOIN (Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado) ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca Where (((NomVaca.NumNomVaca) = " & NumNomVaca & " )) ORDER BY DetalleNomVaca.CodEmpleado"
      Me.DtaConsulta.RecordSource = SQlReportes
      Me.DtaConsulta.Refresh
       
      Me.DtaExporta.Refresh
        'Me.'Me.DtaExporta.Recordset.Edit
       Me.DtaExporta.Recordset("CodigoBAC") = val(Me.TxtCod.Text)
       Me.DtaExporta.Recordset.Update

       Mes = Month(Me.DtaConsulta.Recordset("FechaNomina"))
       Año = Year(Me.DtaConsulta.Recordset("FechaNomina"))
       CanDias = Day(Me.DtaConsulta.Recordset("FechaNomina"))
       Dias = Day(Me.DTFecha.Value)
       Cod = Me.TxtCod.Text
       FechaNomina = Me.DtaConsulta.Recordset("FechaNomina")
       
       
   Case "Nomina13vo"
      NumNomVaca = Frm13VacaMes.TxtNumNom13.Text
      '///////////////////////////Cargo la Consulta de Vacaciones////////////////////////////////
      SQlReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo AS NetoPagar, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Cargo.Cargo, DetalleNom13Mes.SalarioAPagar AS TotalDevengado, Empleado.CodEmpleado1, Empleado.CuentaBanco,Nom13Mes.FechaAplica FROM  Nom13Mes INNER JOIN  Cargo INNER JOIN  Empleado ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN  DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes  " & _
                    "Where (Nom13Mes.NumNom13Mes = " & NumNomVaca & ") ORDER BY Empleado.CodEmpleado1"
       Me.DtaConsulta.RecordSource = SQlReportes
       Me.DtaConsulta.Refresh
       
       Me.DtaExporta.Refresh
        'Me.'Me.DtaExporta.Recordset.Edit
       Me.DtaExporta.Recordset("CodigoBAC") = val(Me.TxtCod.Text)
       Me.DtaExporta.Recordset.Update

       Mes = Month(Me.DtaConsulta.Recordset("FechaAplica"))
       Año = Year(Me.DtaConsulta.Recordset("FechaAplica"))
       CanDias = Day(Me.DtaConsulta.Recordset("FechaAplica"))
       Dias = Day(Me.DTFecha.Value)
       Cod = Me.TxtCod.Text
       QuinLetra = "13vo Mes"
       FechaNomina = Me.DtaConsulta.Recordset("FechaNomina")
       
End Select
    'Dim Heading(8) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
    'Heading(0) = "Nombre"
    'Heading(1) = "Apellidos"
    'Heading(2) = "Direccion"
    'Heading(3) = "Poblacion"
    'Heading(4) = "Provincia"
    'Heading(5) = "Pais"
    'Heading(6) = "Telefono"
    'Heading(7) = "DNI"
            
   
    Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
    'Call Formato_Excel(8, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
V = 2
H = 1
i = 1
           objExcel.ActiveSheet.Columns("C").NumberFormat = "0000#"
           objExcel.ActiveSheet.Columns("E").NumberFormat = "0000#"
          
           objExcel.ActiveSheet.Cells(1, 1) = "B"
           objExcel.ActiveSheet.Cells(1, 2) = "2248"
           objExcel.ActiveSheet.Cells(1, 3) = Cod
            
           objExcel.ActiveSheet.Cells(1, 5) = "0"
           objExcel.ActiveSheet.Cells(1, 6) = Año
           objExcel.ActiveSheet.Cells(1, 7) = Mes
           objExcel.ActiveSheet.Cells(1, 8) = Dias
            'objExcel.ActiveSheet.Cells(1, 10) = CantEmpleados


      TasaCambio = BuscaTasaCambio(FechaNomina)

 
     Do While Not Me.DtaConsulta.Recordset.EOF 'esto nos sirve pa leer los datos desde
       CodEmpleado = DtaConsulta.Recordset("CodEmpleado")
      'la tabla de access para despues colocarlos en las celdas correspondientes
       Nombre = Me.DtaConsulta.Recordset("Nombres")
      If Me.ChkExportaDolares.Value = 1 Then
        Neto = Format(Me.DtaConsulta.Recordset("NetoPagar") / TasaCambio, "####0.00")
        Neto1 = Format(Me.DtaConsulta.Recordset("NetoPagar") / TasaCambio, "##,##0.00")
      Else
        Neto = Format(Me.DtaConsulta.Recordset("NetoPagar"), "####0.00")
        Neto1 = Format(Me.DtaConsulta.Recordset("NetoPagar"), "##,##0.00")
      End If
      
       Longitud = Len(Neto)
       NetoT = Mid(Neto, Longitud - 1, 3)
       NetoT = (Mid(Neto, 1, Longitud - 3)) & NetoT
       With DtaConsulta.Recordset
        If CodEmpleado <> "195" And CodEmpleado <> "196" And CodEmpleado <> "229" Then
       
           If Not (V = 1) Then
             objExcel.ActiveSheet.Cells(V, H) = "T"
           End If
            'objExcel.Cells(1, 1).Format = Text
            objExcel.ActiveSheet.Cells(V, H + 1) = "2248"
            objExcel.ActiveSheet.Cells(V, H + 2) = Cod
            objExcel.ActiveSheet.Cells(V, H + 3) = DtaConsulta.Recordset("NumCedula")
            objExcel.ActiveSheet.Cells(V, H + 4) = Format(.Fields!CodEmpleado1, "00000#")
'            objExcel.ActiveSheet.Cells(V, H + 5) = i
            objExcel.ActiveSheet.Cells(V, H + 5) = Año
            objExcel.ActiveSheet.Cells(V, H + 6) = Mes
            objExcel.ActiveSheet.Cells(V, H + 7) = Dias
            objExcel.ActiveSheet.Cells(V, H + 8) = NetoT
            objExcel.ActiveSheet.Cells(V, H + 10) = QuinLetra
            Nombres = Mid(Nombre, 1, 25)
             objExcel.ActiveSheet.Cells(V, H + 12) = Nombres
            V = V + 1
            i = i + 1
            TotalNomina = TotalNomina + Neto1
            .MoveNext
         Else
            .MoveNext
         End If
   
        End With
     Loop
     
     
       Neto = Format(TotalNomina, "####0.00")
       Longitud = Len(Neto)
       NetoT = Mid(Neto, Longitud - 1, 3)
       NetoT = (Mid(Neto, 1, Longitud - 3)) & NetoT
   
            objExcel.ActiveSheet.Cells(1, 10) = i - 1
            objExcel.ActiveSheet.Cells(1, 9) = NetoT
       

       objExcel.ActiveSheet.Columns("A").ColumnWidth = 1
       objExcel.ActiveSheet.Columns("B").ColumnWidth = 4
        objExcel.ActiveSheet.Columns("C").ColumnWidth = 5
        objExcel.ActiveSheet.Columns("D").ColumnWidth = 20
        objExcel.ActiveSheet.Columns("E").ColumnWidth = 5
        objExcel.ActiveSheet.Columns("F").ColumnWidth = 4
        objExcel.ActiveSheet.Columns("G").ColumnWidth = 2
        objExcel.ActiveSheet.Columns("H").ColumnWidth = 2
        objExcel.ActiveSheet.Columns("I").ColumnWidth = 13
         objExcel.ActiveSheet.Columns("J").ColumnWidth = 3
         objExcel.ActiveSheet.Columns("K").ColumnWidth = 30
         objExcel.ActiveSheet.Columns("L").ColumnWidth = 1
         objExcel.ActiveSheet.Columns("M").ColumnWidth = 30
         
 
 Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto

Exit Sub
TipoErrs:
ControlErrores

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion

End With

With Me.DtaExporta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Exporta"
   .Refresh
End With
Me.DTFecha.Value = Format(Now, "dd/mm/yyyy")
Me.DtaExporta.Refresh
Me.TxtCod.Text = Me.DtaExporta.Recordset("CodigoBAC")
End Sub
