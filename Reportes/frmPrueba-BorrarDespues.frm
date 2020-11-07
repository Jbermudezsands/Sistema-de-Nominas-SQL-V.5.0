VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPrueba 
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc adoIngresos 
      Height          =   495
      Left            =   4320
      Top             =   6720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   873
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Puntualidad"
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
   Begin MSAdodcLib.Adodc adoAntiguedad 
      Height          =   375
      Left            =   1560
      Top             =   6600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Antiguedad"
      Caption         =   "Antiguedad"
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
   Begin VB.CommandButton cmdDevengadoHora 
      Caption         =   "Devengado x Hora"
      Height          =   495
      Left            =   6840
      TabIndex        =   9
      Top             =   5160
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc adoFechaPlanilla 
      Height          =   375
      Left            =   6840
      Top             =   3120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Fecha_Planilla"
      Caption         =   "Fecha Planilla"
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
   Begin VB.CommandButton cmdTrasladoDevengado 
      Caption         =   "Tarifa Horaria"
      Height          =   495
      Left            =   6840
      TabIndex        =   8
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton cmdAsistencia0708 
      Caption         =   "Asistencia 07 y 08"
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   6240
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc adoAsistencia 
      Height          =   330
      Left            =   1080
      Top             =   6120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"frmPrueba-BorrarDespues.frx":0000
      OLEDBString     =   $"frmPrueba-BorrarDespues.frx":0088
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "AsistenciaEmpleado"
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton cmdAsistencias 
      Caption         =   "&Asistencias Empleaos"
      Height          =   495
      Left            =   4080
      TabIndex        =   6
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdRevisionHorarios 
      Caption         =   "Revision"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   5280
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc adoHistorico 
      Height          =   375
      Left            =   2040
      Top             =   4680
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Empleado"
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton cmdCopiar 
      Caption         =   "Ejecutar"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   5280
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc adoTurno 
      Height          =   375
      Left            =   2040
      Top             =   4200
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"frmPrueba-BorrarDespues.frx":0110
      OLEDBString     =   $"frmPrueba-BorrarDespues.frx":0198
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Turno"
      Caption         =   "Adodc1"
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
   Begin MSAdodcLib.Adodc adoHorarioEmpl 
      Height          =   375
      Left            =   3360
      Top             =   1800
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"frmPrueba-BorrarDespues.frx":0220
      OLEDBString     =   $"frmPrueba-BorrarDespues.frx":02A8
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "HorarioEmpleado"
      Caption         =   "Adodc1"
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
   Begin MSAdodcLib.Adodc adoEmpleadoSQL 
      Height          =   375
      Left            =   1440
      Top             =   3600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"frmPrueba-BorrarDespues.frx":0330
      OLEDBString     =   $"frmPrueba-BorrarDespues.frx":03B8
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Empleado"
      Caption         =   "Adodc1"
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
   Begin MSAdodcLib.Adodc adoEmpleadoViejo 
      Height          =   375
      Left            =   1320
      Top             =   2880
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Empleado"
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtResultado 
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtSalida 
      DataField       =   "LSalida"
      DataSource      =   "adoPrueba"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtEntrada 
      DataField       =   "LEntrada"
      DataSource      =   "adoPrueba"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2160
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc adoPrueba 
      Height          =   375
      Left            =   1200
      Top             =   1200
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Devengado_Hora"
      Caption         =   "Devengado por Hora"
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
End
Attribute VB_Name = "frmPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAsistencia0708_Click()


Dim sFecha1 As Variant


sFecha1 = Mid$("07/08/2005", 7, 4) & "-" & Mid$("07/08/2005", 4, 2) & "-" & Mid$("07/08/2005", 1, 2)


Me.adoEmpleadoSQL.CommandType = adCmdText
Me.adoEmpleadoSQL.RecordSource = "SELECT CodEmpleado FROM Empleado"
Me.adoEmpleadoSQL.Refresh




Me.adoAsistencia.CommandType = adCmdText
Me.adoAsistencia.RecordSource = "SELECT * FROM AsistenciaEmpleado WHERE FechaEntrada = CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102)"
Me.adoAsistencia.Refresh



Do While Not Me.adoAsistencia.Recordset.EOF
 
 Me.adoEmpleadoSQL.CommandType = adCmdText
 Me.adoEmpleadoSQL.RecordSource = "SELECT CodEmpleado, CodTipoNomina FROM Empleado WHERE CodEmpleado ='" & Me.adoAsistencia.Recordset.Fields("CodEmpleado")
 Me.adoEmpleadoSQL.Refresh
 

' adoAsistencia.Recordset.AddNew
' adoAsistencia.Recordset.Fields("CodEmpleado") = saEmpleados(bContaEmpleados)
 adoAsistencia.Recordset.Fields("CodTipoNomina") = Me.adoEmpleadoSQL.Recordset.Fields("CodTipoNomina")
 adoAsistencia.Recordset.Fields("FechaEntrada") = "07/08/2005"
 adoAsistencia.Recordset.Fields("HoraEntrada") = "07:00:00"
 adoAsistencia.Recordset.Fields("FechaSalida") = "07/08/2005"
 adoAsistencia.Recordset.Fields("HoraSalida") = "16:00:00"
 adoAsistencia.Recordset.Fields("bActivo") = 0
 adoAsistencia.Recordset.Fields("CodTurno") = "Diurno"
 adoAsistencia.Recordset.Update

Loop










End Sub

Private Sub cmdAsistencias_Click()

Dim saFechas(8) As String
Dim saEmpleados(31) As String
Dim bContaFechas As Byte
Dim bContaEmpleados As Byte

Me.adoAsistencia.CommandType = adCmdTable
Me.adoAsistencia.RecordSource = "AsistenciaEmpleado"
Me.adoAsistencia.Refresh




saFechas(1) = "01/08/2005"
saFechas(2) = "02/08/2005"
saFechas(3) = "03/08/2005"
saFechas(4) = "04/08/2005"
saFechas(5) = "05/08/2005"
saFechas(6) = "06/08/2005"
saFechas(7) = "07/08/2005"

saEmpleados(1) = "000001"
saEmpleados(2) = "000002"
saEmpleados(3) = "000003"
saEmpleados(4) = "000004"
saEmpleados(5) = "000005"
saEmpleados(6) = "000006"
saEmpleados(7) = "000007"
saEmpleados(8) = "000008"
saEmpleados(9) = "000009"
saEmpleados(10) = "000010"
saEmpleados(11) = "000011"
saEmpleados(12) = "000012"
saEmpleados(13) = "000013"
saEmpleados(14) = "000014"
saEmpleados(15) = "000015"
saEmpleados(16) = "000016"
saEmpleados(17) = "000017"
saEmpleados(18) = "000018"
saEmpleados(19) = "000019"
saEmpleados(20) = "000020"
saEmpleados(21) = "000021"
saEmpleados(22) = "000022"
saEmpleados(23) = "000023"
saEmpleados(24) = "000024"
saEmpleados(25) = "000025"
saEmpleados(26) = "000026"
saEmpleados(27) = "000027"
saEmpleados(28) = "000028"
saEmpleados(29) = "000029"
saEmpleados(30) = "000030"


bContaFechas = 1
bContaEmpleados = 1


Do While bContaFechas <= 7




Do While bContaEmpleados <= 30


 adoAsistencia.Recordset.AddNew
 adoAsistencia.Recordset.Fields("CodEmpleado") = saEmpleados(bContaEmpleados)
 adoAsistencia.Recordset.Fields("CodTipoNomina") = "02"
 adoAsistencia.Recordset.Fields("FechaEntrada") = saFechas(bContaFechas)
 adoAsistencia.Recordset.Fields("HoraEntrada") = "07:00:00"
 adoAsistencia.Recordset.Fields("FechaSalida") = saFechas(bContaFechas)
 adoAsistencia.Recordset.Fields("HoraSalida") = "17:30:00"
 adoAsistencia.Recordset.Fields("bActivo") = 0
 adoAsistencia.Recordset.Fields("CodTurno") = "Diurno"
 adoAsistencia.Recordset.Update
 adoAsistencia.Refresh

 bContaEmpleados = bContaEmpleados + 1

Loop

bContaEmpleados = 1
bContaFechas = bContaFechas + 1


Loop














End Sub

Private Sub cmdCopiar_Click()

Dim iHistorico As Integer
Dim sCodEmpl As String
Dim bLongitudCod As Byte


Me.adoEmpleadoViejo.CommandType = adCmdText
Me.adoEmpleadoViejo.RecordSource = "SELECT * FROM Empleado ORDER BY Cod_Empl ASC"
Me.adoEmpleadoViejo.Refresh

Me.adoEmpleadoSQL.CommandType = adCmdTable
Me.adoEmpleadoSQL.RecordSource = "Empleado"
Me.adoEmpleadoSQL.Refresh

Me.adoHistorico.CommandType = adCmdTable
Me.adoHistorico.RecordSource = "Historico"
Me.adoHistorico.Refresh

Me.adoHorarioEmpl.CommandType = adCmdTable
Me.adoHorarioEmpl.RecordSource = "HorarioEmpleado"
Me.adoHorarioEmpl.Refresh

Me.adoTurno.CommandType = adCmdTable
Me.adoTurno.RecordSource = "Turno"
Me.adoTurno.Refresh

iHistorico = 1

Do While Not Me.adoEmpleadoViejo.Recordset.EOF

   sCodEmpl = Me.adoEmpleadoViejo.Recordset.Fields("Cod_Empl")
   bLongitudCod = 6 - Len(sCodEmpl)
   
   Select Case bLongitudCod
   
   
   Case 1:
      sCodEmpl = "0" & sCodEmpl
      
   Case 2:
      sCodEmpl = "00" & sCodEmpl
      
   Case 3:
      sCodEmpl = "000" & sCodEmpl
   
   Case 4:
      sCodEmpl = "0000" & sCodEmpl

   Case 5:
      sCodEmpl = "00000" & sCodEmpl
    
 End Select
 
 
 Me.adoEmpleadoSQL.Recordset.AddNew
 Me.adoEmpleadoSQL.Recordset.Fields("CodEmpleado") = sCodEmpl
 Me.adoEmpleadoSQL.Recordset.Fields("Nombre1") = Mid$(Me.adoEmpleadoViejo.Recordset.Fields("Nombre"), 1, 20)
 Me.adoEmpleadoSQL.Recordset.Fields("Nombre2") = "."
 Me.adoEmpleadoSQL.Recordset.Fields("Apellido1") = "."
 Me.adoEmpleadoSQL.Recordset.Fields("Apellido2") = "."
 Me.adoEmpleadoSQL.Recordset.Fields("CodDepartamento") = Me.adoEmpleadoViejo.Recordset.Fields("Cod_Depto")
 Me.adoEmpleadoSQL.Recordset.Fields("CodGrupo") = "02"
 Me.adoEmpleadoSQL.Recordset.Fields("CodTipoNomina") = "02"
 Me.adoEmpleadoSQL.Recordset.Fields("TarifaHoraria") = Me.adoEmpleadoViejo.Recordset.Fields("Sal_Hora")
 Me.adoEmpleadoSQL.Recordset.Fields("NumeroInss") = Me.adoEmpleadoViejo.Recordset.Fields("Inss")
 Me.adoEmpleadoSQL.Recordset.Fields("SalarioFijo") = "N"
 Me.adoEmpleadoSQL.Recordset.Fields("Nacionalidad") = "Nicaraguense"
 Me.adoEmpleadoSQL.Recordset.Fields("CodCargo") = "02"
 Me.adoEmpleadoSQL.Recordset.Fields("Activo") = 1
 Me.adoEmpleadoSQL.Recordset.Update
 Me.adoEmpleadoSQL.Refresh
 
 
 Me.adoHistorico.Recordset.AddNew
 Me.adoHistorico.Recordset.Fields("Id") = iHistorico
 Me.adoHistorico.Recordset.Fields("CodEmpleado") = sCodEmpl
 Me.adoHistorico.Recordset.Fields("FechaContrato") = Me.adoEmpleadoViejo.Recordset.Fields("Fech_Ing")
 Me.adoHistorico.Recordset.Fields("SueldoActual") = 0
 Me.adoHistorico.Recordset.Fields("SueldoAnterior") = 0
 Me.adoHistorico.Recordset.Fields("SueldoInicial") = 0
 Me.adoHistorico.Recordset.Update
 
 iHistorico = iHistorico + 1
 
' Me.adoHorarioEmpl.Recordset.AddNew
' Me.adoHorarioEmpl.Recordset.Fields("CodEmpleado") = sCodEmpl
' Me.adoHorarioEmpl.Recordset.Fields("LEntrada") = Me.adoTurno.Recordset.Fields("LEntrada")
' Me.adoHorarioEmpl.Recordset.Fields("LSalida") = Me.adoTurno.Recordset.Fields("LSalida")
' Me.adoHorarioEmpl.Recordset.Fields("MEntrada") = Me.adoTurno.Recordset.Fields("MEntrada")
' Me.adoHorarioEmpl.Recordset.Fields("MSalida") = Me.adoTurno.Recordset.Fields("MSalida")
' Me.adoHorarioEmpl.Recordset.Fields("MCEntrada") = Me.adoTurno.Recordset.Fields("MCEntrada")
' Me.adoHorarioEmpl.Recordset.Fields("MCSalida") = Me.adoTurno.Recordset.Fields("MCSalida")
' Me.adoHorarioEmpl.Recordset.Fields("JEntrada") = Me.adoTurno.Recordset.Fields("JEntrada")
' Me.adoHorarioEmpl.Recordset.Fields("JSalida") = Me.adoTurno.Recordset.Fields("JSalida")
' Me.adoHorarioEmpl.Recordset.Fields("VEntrada") = Me.adoTurno.Recordset.Fields("VEntrada")
' Me.adoHorarioEmpl.Recordset.Fields("VSalida") = Me.adoTurno.Recordset.Fields("VSalida")
' Me.adoHorarioEmpl.Recordset.Fields("SEntrada") = Me.adoTurno.Recordset.Fields("SEntrada")
' Me.adoHorarioEmpl.Recordset.Fields("SSalida") = Me.adoTurno.Recordset.Fields("SSalida")
' Me.adoHorarioEmpl.Recordset.Fields("DEntrada") = Me.adoTurno.Recordset.Fields("DEntrada")
' Me.adoHorarioEmpl.Recordset.Fields("DSalida") = Me.adoTurno.Recordset.Fields("DSalida")
' Me.adoHorarioEmpl.Recordset.Fields("TurnoLunes") = Me.adoTurno.Recordset.Fields("CodTurno")
' Me.adoHorarioEmpl.Recordset.Fields("TurnoMartes") = Me.adoTurno.Recordset.Fields("CodTurno")
' Me.adoHorarioEmpl.Recordset.Fields("TurnoMiercoles") = Me.adoTurno.Recordset.Fields("CodTurno")
' Me.adoHorarioEmpl.Recordset.Fields("TurnoJueves") = Me.adoTurno.Recordset.Fields("CodTurno")
' Me.adoHorarioEmpl.Recordset.Fields("TurnoViernes") = Me.adoTurno.Recordset.Fields("CodTurno")
' Me.adoHorarioEmpl.Recordset.Fields("TurnoSabado") = Me.adoTurno.Recordset.Fields("CodTurno")
' Me.adoHorarioEmpl.Recordset.Fields("TurnoDomingo") = Me.adoTurno.Recordset.Fields("CodTurno")
' Me.adoHorarioEmpl.Recordset.Fields("TComida") = Me.adoTurno.Recordset.Fields("TComida")
' Me.adoHorarioEmpl.Recordset.Update
' Me.adoHorarioEmpl.Refresh
 
' Me.adoEmpleadoViejo.Recordset.MoveNext
 
 

Loop




End Sub

Private Sub cmdDevengadoHora_Click()

Dim sFecha1 As Variant
Dim sCodViejo As String
Dim iConta As Integer
Dim iAciertos As Integer
Dim sFecha2 As String
Dim iPeriodo As Integer
Dim iAno As Integer
Dim sMes As String
Dim sngHorasLaboradas As Single
Dim sngHorasExtras As Single
Dim sngTotalHoras As Single
Dim sngHoraLunes As Single
Dim sngHoraMartes As Single
Dim sngHoraMiercoles As Single
Dim sngHoraJueves As Single
Dim sngHoraViernes As Single
Dim sngHoraSabado As Single
Dim sngHoraDomingo As Single
Dim iDias As Integer
Dim sngInc As Single
Dim sngSeptimo As Single


Me.adoPrueba.CommandType = adCmdTable
Me.adoPrueba.RecordSource = "Devengado_Hora"
Me.adoPrueba.Refresh

sFecha1 = Mid$("01/08/2005", 7, 4) & "-" & Mid$("01/08/2005", 4, 2) & "-" & Mid$("01/08/2005", 1, 2)
sFecha2 = Mid$("07/08/2005", 7, 4) & "-" & Mid$("07/08/2005", 4, 2) & "-" & Mid$("07/08/2005", 1, 2)

Me.adoEmpleadoSQL.CommandType = adCmdText
Me.adoEmpleadoSQL.RecordSource = "SELECT CodEmpleado, TarifaHoraria, CodTipoNomina FROM Empleado WHERE CodTipoNomina ='02'"
Me.adoEmpleadoSQL.Refresh

Me.adoEmpleadoViejo.CommandType = adCmdText
Me.adoEmpleadoViejo.RecordSource = "SELECT * FROM Empleado"
Me.adoEmpleadoViejo.Refresh

Me.adoIngresos.CommandType = adCmdText
Me.adoIngresos.RecordSource = "SELECT * FROM Ingreso_Empl"
Me.adoIngresos.Refresh

Me.adoFechaPlanilla.CommandType = adCmdText
Me.adoFechaPlanilla.RecordSource = "SELECT * FROM Fecha_Planilla WHERE Actual =True"
Me.adoFechaPlanilla.Refresh

Me.adoAsistencia.CommandType = adCmdText
Me.adoAsistencia.RecordSource = "SELECT * FROM AsistenciaEmpleado WHERE FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND FechaSalida < = CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102)"
Me.adoAsistencia.Refresh


iPeriodo = Me.adoFechaPlanilla.Recordset.Fields("Periodo")
iAno = Me.adoFechaPlanilla.Recordset.Fields("año")
sMes = Me.adoFechaPlanilla.Recordset.Fields("mes")



Do While Not Me.adoEmpleadoSQL.Recordset.EOF
   
    

sCodViejo = Mid$(Me.adoEmpleadoSQL.Recordset.Fields("CodEmpleado"), 2, 5)
iConta = 1
iAciertos = 0

Do While iConta <= 5
    
 If InStr(1, Mid$(sCodViejo, iConta, 1), "0", vbTextCompare) = 1 Then
    iAciertos = iAciertos + 1

 Else
    iConta = 6
 End If

 iConta = iConta + 1

Loop
  
sCodViejo = Mid$(sCodViejo, iAciertos + 1, Len(sCodViejo) - 1)

   Me.adoPrueba.CommandType = adCmdText
   Me.adoPrueba.RecordSource = "SELECT * FROM Devengado_Hora WHERE Cod_Empl =" & sCodViejo & " AND Periodo =" & iPeriodo & " AND mes ='" & sMes & "' AND año =" & iAno
   Me.adoPrueba.Refresh

If Me.adoPrueba.Recordset.EOF Then
   
   Me.adoAsistencia.CommandType = adCmdText
   Me.adoAsistencia.RecordSource = "SELECT * FROM AsistenciaEmpleado WHERE CodEmpleado ='" & Me.adoEmpleadoSQL.Recordset.Fields("CodEmpleado") & "' AND FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND FechaSalida < = CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) ORDER BY FechaEntrada ASC"
   Me.adoAsistencia.Refresh
   
   iDias = 1
   sngHoraLunes = 0
   sngHoraMartes = 0
   sngHoraMiercoles = 0
   sngHoraJueves = 0
   sngHoraViernes = 0
   sngHoraSabado = 0
   sngHoraDomingo = 0
   sngHorasLaboradas = 0
   sngHorasExtras = 0
   sngSeptimo = 0
   
   If Not Me.adoAsistencia.Recordset.EOF Then
   
   Do While iDias <= 7 And Not Me.adoAsistencia.Recordset.EOF
      
      Select Case iDias
      
      Case 1:
           If Me.adoAsistencia.Recordset.Fields("Dia") = "Lun" Then
              sngHoraLunes = Me.adoAsistencia.Recordset.Fields("HLaboradas") + Me.adoAsistencia.Recordset.Fields("HExtras")
              sngHorasLaboradas = Me.adoAsistencia.Recordset.Fields("HLaboradas")
              sngHorasExtras = sngHorasExtras + Me.adoAsistencia.Recordset.Fields("HExtras")
              Me.adoAsistencia.Recordset.MoveNext
           Else
              sngHoraLunes = 0
           End If
           
       Case 2:
           If Me.adoAsistencia.Recordset.Fields("Dia") = "Mart" Then
              sngHoraMartes = Me.adoAsistencia.Recordset.Fields("HLaboradas") + Me.adoAsistencia.Recordset.Fields("HExtras")
              sngHorasLaboradas = sngHorasLaboradas + Me.adoAsistencia.Recordset.Fields("HLaboradas")
              sngHorasExtras = sngHorasExtras + Me.adoAsistencia.Recordset.Fields("HExtras")
              Me.adoAsistencia.Recordset.MoveNext
           Else
              sngHoraMartes = 0
           End If
       
       Case 3:
           If Me.adoAsistencia.Recordset.Fields("Dia") = "Mierc" Then
              sngHoraMiercoles = Me.adoAsistencia.Recordset.Fields("HLaboradas") + Me.adoAsistencia.Recordset.Fields("HExtras")
              sngHorasLaboradas = sngHorasLaboradas + Me.adoAsistencia.Recordset.Fields("HLaboradas")
              sngHorasExtras = sngHorasExtras + Me.adoAsistencia.Recordset.Fields("HExtras")
              Me.adoAsistencia.Recordset.MoveNext
           Else
              sngHoraMiercoles = 0
           End If
           
       Case 4:
           If Me.adoAsistencia.Recordset.Fields("Dia") = "Juev" Then
              sngHoraJueves = Me.adoAsistencia.Recordset.Fields("HLaboradas") + Me.adoAsistencia.Recordset.Fields("HExtras")
              sngHorasLaboradas = sngHorasLaboradas + Me.adoAsistencia.Recordset.Fields("HLaboradas")
              sngHorasExtras = sngHorasExtras + Me.adoAsistencia.Recordset.Fields("HExtras")
              Me.adoAsistencia.Recordset.MoveNext
           Else
              sngHoraJueves = 0
           End If
           
       Case 5:
           If Me.adoAsistencia.Recordset.Fields("Dia") = "Viern" Then
              sngHoraViernes = Me.adoAsistencia.Recordset.Fields("HLaboradas") + Me.adoAsistencia.Recordset.Fields("HExtras")
              sngHorasLaboradas = sngHorasLaboradas + Me.adoAsistencia.Recordset.Fields("HLaboradas")
              sngHorasExtras = sngHorasExtras + Me.adoAsistencia.Recordset.Fields("HExtras")
              Me.adoAsistencia.Recordset.MoveNext
           Else
              sngHoraViernes = 0
           End If
           
       Case 6:
           If Me.adoAsistencia.Recordset.Fields("Dia") = "Sab" Then
              sngHoraSabado = Me.adoAsistencia.Recordset.Fields("HLaboradas") + Me.adoAsistencia.Recordset.Fields("HExtras")
              sngHorasLaboradas = sngHorasLaboradas + Me.adoAsistencia.Recordset.Fields("HLaboradas")
              sngHorasExtras = sngHorasExtras + Me.adoAsistencia.Recordset.Fields("HExtras")
              Me.adoAsistencia.Recordset.MoveNext
           Else
              sngHoraSabado = 0
           End If
           
       Case 7:
           If Me.adoAsistencia.Recordset.Fields("Dia") = "Dom" Then
              sngHoraDomingo = Me.adoAsistencia.Recordset.Fields("HLaboradas") + Me.adoAsistencia.Recordset.Fields("HExtras")
              sngHorasLaboradas = sngHorasLaboradas + Me.adoAsistencia.Recordset.Fields("HLaboradas")
              sngHorasExtras = sngHorasExtras + Me.adoAsistencia.Recordset.Fields("HExtras")
              Me.adoAsistencia.Recordset.MoveNext
           Else
              sngHoraDomingo = 0
           End If
           
           
           
      End Select
      
      iDias = iDias + 1
      
      
   Loop
   
   Me.adoEmpleadoViejo.CommandType = adCmdText
   Me.adoEmpleadoViejo.RecordSource = "SELECT * FROM Empleado WHERE Cod_Empl =" & CInt(sCodViejo)
   Me.adoEmpleadoViejo.Refresh
   
   sngInc = BuscarIncentivo((56 * CSng(Me.adoEmpleadoSQL.Recordset.Fields("TarifaHoraria"))), Me.adoEmpleadoViejo.Recordset.Fields("Fech_Ing"))
   
   
   If sngHorasLaboradas >= 47 Then
      sngSeptimo = Me.adoEmpleadoSQL.Recordset.Fields("TarifaHoraria") * 8
   Else
      sngSeptimo = 0
   End If
   
   
   
    Me.adoPrueba.Recordset.AddNew
      Me.adoPrueba.Recordset.Fields(0) = sCodViejo
      Me.adoPrueba.Recordset.Fields(1) = iPeriodo
      Me.adoPrueba.Recordset.Fields(2) = iAno
      Me.adoPrueba.Recordset.Fields(3) = sMes
      Me.adoPrueba.Recordset.Fields(4) = sngHoraLunes
      Me.adoPrueba.Recordset.Fields(5) = sngHoraMartes
      Me.adoPrueba.Recordset.Fields(6) = sngHoraMiercoles
      Me.adoPrueba.Recordset.Fields(7) = sngHoraJueves
      Me.adoPrueba.Recordset.Fields(8) = sngHoraViernes
      Me.adoPrueba.Recordset.Fields(9) = sngHoraSabado
      Me.adoPrueba.Recordset.Fields(10) = sngHoraDomingo
      Me.adoPrueba.Recordset.Fields(11) = CSng(Me.adoEmpleadoSQL.Recordset.Fields("TarifaHoraria")) * sngHorasLaboradas
      Me.adoPrueba.Recordset.Fields("HExtras") = Format(CSng(sngHorasExtras), "##0.##") * Me.adoEmpleadoSQL.Recordset.Fields("TarifaHoraria") * 2
      Me.adoPrueba.Recordset.Fields("Sept") = Format(CSng(sngSeptimo), "##0.##")
      Me.adoPrueba.Recordset.Fields("Antig") = Format(sngInc, "##0.##")
      Me.adoPrueba.Recordset.Fields("Transp") = 0
      Me.adoPrueba.Recordset.Fields("Alim") = 0
      Me.adoPrueba.Recordset.Fields("INSSLab") = 0
      Me.adoPrueba.Recordset.Fields("Perdida") = 0
      Me.adoPrueba.Recordset.Fields("OtrasDeducc") = 0
      Me.adoPrueba.Recordset.Fields("NHextras") = sngHorasExtras
      Me.adoPrueba.Recordset.Fields("NHOrd") = sngHorasLaboradas
      Me.adoPrueba.Recordset.Fields("HJueves") = 0
      
    Me.adoPrueba.Recordset.Update
   
     If sngHorasExtras > 0 Then
        Me.adoIngresos.Recordset.AddNew
        Me.adoIngresos.Recordset.Fields("Cod_Empl") = CInt(sCodViejo)
        Me.adoIngresos.Recordset.Fields("Periodo") = iPeriodo
        Me.adoIngresos.Recordset.Fields("Año") = iAno
        Me.adoIngresos.Recordset.Fields("mes") = sMes
        Me.adoIngresos.Recordset.Fields("Cod_Ing") = "03"
        Me.adoIngresos.Recordset.Fields("Ingreso") = Format(sngHorasExtras * 2 * Me.adoEmpleadoSQL.Recordset.Fields("TarifaHoraria"), "###.##")
        Me.adoIngresos.Recordset.Update
      End If
        
        
      If sngInc > 0 Then
        Me.adoIngresos.Recordset.AddNew
        Me.adoIngresos.Recordset.Fields("Cod_Empl") = CInt(sCodViejo)
        Me.adoIngresos.Recordset.Fields("Periodo") = iPeriodo
        Me.adoIngresos.Recordset.Fields("Año") = iAno
        Me.adoIngresos.Recordset.Fields("mes") = sMes
        Me.adoIngresos.Recordset.Fields("Cod_Ing") = "05"
        Me.adoIngresos.Recordset.Fields("Ingreso") = Format(sngInc, "##.##")
        Me.adoIngresos.Recordset.Update
      End If
        
      If sngHorasLaboradas >= 48 Then
        Me.adoIngresos.Recordset.AddNew
        Me.adoIngresos.Recordset.Fields("Cod_Empl") = CInt(sCodViejo)
        Me.adoIngresos.Recordset.Fields("Periodo") = iPeriodo
        Me.adoIngresos.Recordset.Fields("Año") = iAno
        Me.adoIngresos.Recordset.Fields("mes") = sMes
        Me.adoIngresos.Recordset.Fields("Cod_Ing") = "09"
        Me.adoIngresos.Recordset.Fields("Ingreso") = 32
        Me.adoIngresos.Recordset.Update
      End If
      
      If sngSeptimo > 0 Then
        Me.adoIngresos.Recordset.AddNew
        Me.adoIngresos.Recordset.Fields("Cod_Empl") = CInt(sCodViejo)
        Me.adoIngresos.Recordset.Fields("Periodo") = iPeriodo
        Me.adoIngresos.Recordset.Fields("Año") = iAno
        Me.adoIngresos.Recordset.Fields("mes") = sMes
        Me.adoIngresos.Recordset.Fields("Cod_Ing") = "04"
        Me.adoIngresos.Recordset.Fields("Ingreso") = Format(sngSeptimo, "##.##")
        Me.adoIngresos.Recordset.Update
      End If
      
      
        
  End If
   
 End If
   
  Me.adoEmpleadoSQL.Recordset.MoveNext
   
Loop

End Sub

Private Sub cmdRevisionHorarios_Click()

Dim sFecha1 As Variant
Dim sFecha2 As Variant

sFecha1 = Mid$("08/08/2005", 7, 4) & "-" & Mid$("08/08/2005", 4, 2) & "-" & Mid$("08/08/2005", 1, 2)
sFecha2 = Mid$("14/08/2005", 7, 4) & "-" & Mid$("14/08/2005", 4, 2) & "-" & Mid$("14/08/2005", 1, 2)


Me.adoEmpleadoSQL.CommandType = adCmdText
Me.adoEmpleadoSQL.RecordSource = "SELECT CodEmpleado, CodTipoNomina FROM Empleado WHERE CodTipoNomina ='02'"
Me.adoEmpleadoSQL.Refresh




Me.adoHorarioEmpl.CommandType = adCmdTable
Me.adoHorarioEmpl.RecordSource = "HorarioEmpleado"
Me.adoHorarioEmpl.Refresh

Me.adoAsistencia.CommandType = adCmdText
Me.adoAsistencia.RecordSource = "SELECT * FROM AsistenciaEmpleado WHERE FechaEntrada = CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102)"  'AND FechaSalida <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND CodEmpleado ='" & Me.adoEmpleadoSQL.Recordset.Fields("CodEmpleado") & "'"
Me.adoAsistencia.Refresh



Do While Not Me.adoEmpleadoSQL.Recordset.EOF
 
 Me.adoAsistencia.CommandType = adCmdText
 Me.adoAsistencia.RecordSource = "SELECT * FROM AsistenciaEmpleado WHERE FechaEntrada = CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND CodEmpleado ='" & Me.adoEmpleadoSQL.Recordset.Fields("CodEmpleado") & "'" ' AND FechaSalida <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND CodEmpleado ='" & Me.adoEmpleadoSQL.Recordset.Fields("CodEmpleado") & "'"
 Me.adoAsistencia.Refresh
 
 If Not Me.adoAsistencia.Recordset.EOF Then
   'If Me.adoAsistencia.Recordset.RecordCount = 1 And Me.adoAsistencia.Recordset.Fields("FechaEntrada") = "10/08/2005" Then
      Me.adoAsistencia.Recordset.Fields("FechaSalida") = "08/08/2005"
      Me.adoAsistencia.Recordset.Update
      Me.adoAsistencia.Refresh
   'End If
 End If
 
 Me.adoEmpleadoSQL.Recordset.MoveNext

Loop




End Sub

Private Sub cmdTrasladoDevengado_Click()


Dim sFecha1 As Variant
Dim sCodViejo As String
Dim iConta As Integer
Dim iAciertos As Integer
Dim sFecha2 As String
Dim iHistorico As Integer

sFecha1 = Mid$("01/08/2005", 7, 4) & "-" & Mid$("01/08/2005", 4, 2) & "-" & Mid$("01/08/2005", 1, 2)
sFecha2 = Mid$("07/08/2005", 7, 4) & "-" & Mid$("07/08/2005", 4, 2) & "-" & Mid$("07/08/2005", 1, 2)


'Me.adoHorarioEmpl.CommandType = adCmdText
'Me.adoHorarioEmpl.RecordSource = "SELECT CodEmpleado, CodTipoNomina FROM Empleado WHERE CodTipoNomina ='02'"
'Me.adoHorarioEmpl.Refresh
'
'Do While Not Me.adoHorarioEmpl.Recordset.EOF
'
'  Me.adoEmpleadoSQL.CommandType = adCmdText
'  Me.adoEmpleadoSQL.RecordSource = "SELECT * FROM AsistenciaEmpleado WHERE CodEmpleado ='" & Me.adoHorarioEmpl.Recordset.Fields("CodEmpleado") & "' AND FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND FechaSalida < = CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102)"
'  Me.adoEmpleadoSQL.Refresh
'
' If Not Me.adoEmpleadoSQL.Recordset.EOF Then
'  If Me.adoEmpleadoSQL.Recordset.RecordCount = 1 And Me.adoEmpleadoSQL.Recordset.Fields("FechaEntrada") = "01/08/2005" Then
'     Me.adoEmpleadoSQL.Recordset.Delete
'     Me.adoEmpleadoSQL.Refresh
'     'MsgBox "Codigo: " & Me.adoHorarioEmpl.Recordset.Fields("CodEmpleado")
'  End If
' End If
'
'  Me.adoHorarioEmpl.Recordset.MoveNext
'
'Loop
'

Me.adoPrueba.ConnectionString = "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=" & "PAYROLL"
Me.adoPrueba.CommandType = adCmdTable
Me.adoPrueba.RecordSource = "Historico"
Me.adoPrueba.Refresh

Me.adoEmpleadoSQL.ConnectionString = "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=" & "PAYROLL"
Me.adoEmpleadoSQL.CommandType = adCmdTable
Me.adoEmpleadoSQL.RecordSource = "Empleado"
Me.adoEmpleadoSQL.Refresh

'
'
Me.adoHistorico.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
Me.adoHistorico.CommandType = adCmdTable
Me.adoHistorico.RecordSource = "Empleado"
Me.adoHistorico.Refresh

iHistorico = 41
 
 
 
Do While Not Me.adoEmpleadoSQL.Recordset.EOF

   sCodViejo = Mid$(Me.adoEmpleadoSQL.Recordset.Fields("CodEmpleado"), 2, 5)
   iConta = 1
   iAciertos = 0

   Do While iConta <= 5

      If InStr(1, Mid$(sCodViejo, iConta, 1), "0", vbTextCompare) = 1 Then
       iAciertos = iAciertos + 1

      Else
        iConta = 6
      End If

   iConta = iConta + 1

  Loop
  
  sCodViejo = Mid$(sCodViejo, iAciertos + 1, Len(sCodViejo) - 1)



  Me.adoHistorico.CommandType = adCmdText
  Me.adoHistorico.RecordSource = "SELECT * FROM Empleado WHERE Cod_Empl = " & CInt(sCodViejo)
  Me.adoHistorico.Refresh
'
'If Not Me.adoHistorico.Recordset.EOF Then
'
'  Me.adoPrueba.Recordset.AddNew
'  Me.adoPrueba.Recordset.Fields("Id") = iHistorico
'  Me.adoPrueba.Recordset.Fields("CodEmpleado") = Me.adoEmpleadoSQL.Recordset.Fields("CodEmpleado")
'  Me.adoPrueba.Recordset.Fields("FechaContrato") = Me.adoHistorico.Recordset.Fields("Fech_Ing")
'  Me.adoPrueba.Recordset.Fields("SueldoActual") = 0
'  Me.adoPrueba.Recordset.Fields("SueldoAnterior") = 0
'  Me.adoPrueba.Recordset.Fields("SueldoInicial") = 0
'  Me.adoPrueba.Recordset.Update
'
'  iHistorico = iHistorico + 1
'Else
'  MsgBox "El empleado " & sCodViejo & " no se encuentra en la BD anterior de Access"
'
'End If
  
  If Not Me.adoHistorico.Recordset.EOF Then
     Me.adoEmpleadoSQL.Recordset.Fields("TarifaHoraria") = Me.adoHistorico.Recordset.Fields("Sal_hora")
     Me.adoEmpleadoSQL.Recordset.Update

  Else
     MsgBox "El codigo " & sCodViejo & " se tiene que revisar"


  End If

  Me.adoEmpleadoSQL.Recordset.MoveNext

Loop








End Sub

Private Sub Command1_Click()

Dim dHEntradaHoy As Date

'Me.txtSalida.Text = Me.adoPrueba.Recordset.Fields("LSalida") - Me.adoPrueba.Recordset.Fields("LEntrada")

'Para deducir el valor del valor de horas trabajadas por dia

 Me.txtSalida.Text = (DateDiff("n", Me.adoPrueba.Recordset.Fields("LEntrada"), Me.adoPrueba.Recordset.Fields("LSalida")) / 60) - (Me.adoPrueba.Recordset.Fields("TComida") / 60)



If Me.adoPrueba.Recordset.Fields("LEntrada") > Time Then
         dHEntradaHoy = Me.adoPrueba.Recordset.Fields("LEntrada")
Else
         dHEntradaHoy = Time
End If


End Sub


Public Function BuscarIncentivo(sTotal As String, Fecha As Date) As Single


Dim fecAct As Date

fecAct = Format(Date, "Short Date")
adoAntiguedad.RecordSource = "Antiguedad"
adoAntiguedad.Refresh

  ' 1 Año
If fecAct - Fecha >= 365 And fecAct - Fecha <= 2 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 1 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
  ' 2 Años
    
ElseIf fecAct - Fecha >= 2 * 365 And fecAct - Fecha <= 3 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 2 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
   
   ' 3 Años
        
ElseIf fecAct - Fecha >= 3 * 365 And fecAct - Fecha <= 4 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 3 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
     
   ' 4 Años
ElseIf fecAct - Fecha >= 4 * 365 And fecAct - Fecha <= 5 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 4 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
  ' 5 Años
ElseIf fecAct - Fecha >= 5 * 365 And fecAct - Fecha <= 6 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 5 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
  ' 6 Años
ElseIf fecAct - Fecha >= 6 * 365 And fecAct - Fecha <= 7 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 6 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
   
   
   ' 7 Años
ElseIf fecAct - Fecha >= 7 * 365 And fecAct - Fecha <= 8 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 7 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
  ' 8 Años
ElseIf fecAct - Fecha >= 8 * 365 And fecAct - Fecha <= 9 * 365 Then
        
   adoAntiguedad.Recordset.Find "[años_acum] like " & 8 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
   
   ' 9 Años
   
ElseIf fecAct - Fecha >= 9 * 365 And fecAct - Fecha <= 10 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 9 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
  ' 10 Años
  
ElseIf fecAct - Fecha >= 10 * 365 And fecAct - Fecha <= 11 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 10 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
    
 ' 11 Años
ElseIf fecAct - Fecha >= 11 * 365 And fecAct - Fecha <= 12 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 11 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
  ' 12 Años
ElseIf fecAct - Fecha >= 12 * 365 And fecAct - Fecha <= 13 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 12 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
   
' 13 Años
  
ElseIf fecAct - Fecha >= 13 * 365 And fecAct - Fecha <= 14 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 13 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
   
' 14 Años
ElseIf fecAct - Fecha >= 14 * 365 And fecAct - Fecha <= 15 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 14 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
' 15 Años
ElseIf fecAct - Fecha >= 15 * 365 And fecAct - Fecha <= 16 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 15 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
 
 ' 16 Años
ElseIf fecAct - Fecha >= 16 * 365 And fecAct - Fecha <= 16 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 16 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
 ' 17 Años
ElseIf fecAct - Fecha >= 17 * 365 And fecAct - Fecha <= 18 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 17 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
' 18 Años
ElseIf fecAct - Fecha >= 18 * 365 And fecAct - Fecha <= 19 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 18 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
 
 ' 19 Años
 ElseIf fecAct - Fecha >= 19 * 365 And fecAct - Fecha <= 20 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 19 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
  ' 20 Años
 ElseIf fecAct - Fecha >= 20 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 20 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
   
 Else
   BuscarIncentivo = 0
  
End If

'adoAntiguedad.RecordSource = "Departamento"
'adoAntiguedad.Refresh




End Function

Private Sub Form_Load()

Dim ruta As String
Dim conexion As Variant
Dim Server As String

'RutaServer = App.Path + "\CntNominas.dll"
'
'  With Me.dtaServidor
'     .DatabaseName = RutaServer
'     .RecordSource = "Servidor"
'     .Refresh
'  End With
'
'  If Not IsNull(Me.dtaServidor.Recordset.Servidor) Then
'   Server = Me.dtaServidor.Recordset.Servidor
'  Else
'   MsgBox "No se ha definido el Servidor", vbCritical, "Sistmea de Nominas"
'   Exit Sub
'  End If


Server = "PAYROLL"

ruta = App.Path & "\PlanMetro.mdb"
conexion = "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=" & Server


Me.adoEmpleadoViejo.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & ruta
Me.adoIngresos.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & ruta
Me.adoPrueba.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & ruta
Me.adoFechaPlanilla.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & ruta
Me.adoAsistencia.ConnectionString = conexion

Me.adoFechaPlanilla.CommandType = adCmdText
'Me.adoFechaPlanilla.RecordSource


Me.adoPrueba.CommandType = adCmdTable
Me.adoPrueba.RecordSource = "Devengado_Hora"
Me.adoPrueba.Refresh

'sFecha1 = Mid$("01/08/2005", 7, 4) & "-" & Mid$("01/08/2005", 4, 2) & "-" & Mid$("01/08/2005", 1, 2)
'sFecha2 = Mid$("07/08/2005", 7, 4) & "-" & Mid$("07/08/2005", 4, 2) & "-" & Mid$("07/08/2005", 1, 2)

Me.adoEmpleadoSQL.CommandType = adCmdText
Me.adoEmpleadoSQL.RecordSource = "SELECT CodEmpleado, TarifaHoraria, CodTipoNomina FROM Empleado WHERE CodTipoNomina ='02'"
Me.adoEmpleadoSQL.Refresh

Me.adoEmpleadoViejo.CommandType = adCmdText
Me.adoEmpleadoViejo.RecordSource = "SELECT * FROM Empleado"
Me.adoEmpleadoViejo.Refresh

Me.adoIngresos.CommandType = adCmdText
Me.adoIngresos.RecordSource = "SELECT * FROM Ingreso_Empl"
Me.adoIngresos.Refresh

Me.adoFechaPlanilla.CommandType = adCmdText
Me.adoFechaPlanilla.RecordSource = "SELECT * FROM Fecha_Planilla WHERE Actual =True"
Me.adoFechaPlanilla.Refresh

Me.adoAsistencia.CommandType = adCmdText
Me.adoAsistencia.RecordSource = "SELECT * FROM AsistenciaEmpleado" ' WHERE FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND FechaSalida < = CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102)"
Me.adoAsistencia.Refresh


'iPeriodo = Me.adoFechaPlanilla.Recordset.Fields("Periodo")
'iAno = Me.adoFechaPlanilla.Recordset.Fields("año")
'sMes = Me.adoFechaPlanilla.Recordset.Fields("mes")





End Sub


