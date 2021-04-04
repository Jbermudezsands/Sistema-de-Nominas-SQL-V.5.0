VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form frmRepAsistencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes - Asistencias Empleados"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9990
   Begin MSAdodcLib.Adodc adoTipoNomina 
      Height          =   375
      Left            =   5040
      Top             =   6960
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
      Caption         =   "adoTipoNomina"
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
   Begin VB.TextBox txtSQL 
      Height          =   855
      Left            =   120
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   8400
      Width           =   8775
   End
   Begin VB.Data dtaServidor 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Width           =   4455
   End
   Begin MSAdodcLib.Adodc adoIncentivo 
      Height          =   330
      Left            =   720
      Top             =   8160
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   582
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
      Connect         =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemasNominas;Data Source=METRO"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemasNominas;Data Source=METRO"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Incentivos"
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
   Begin MSAdodcLib.Adodc adoHorasExtras 
      Height          =   330
      Left            =   4680
      Top             =   7560
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
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
      Caption         =   "Horas Extras"
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
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   840
      Top             =   7440
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   582
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
      Caption         =   "Consulta"
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
   Begin MSAdodcLib.Adodc adoTurno 
      Height          =   375
      Left            =   6720
      Top             =   6360
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "Turno"
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
   Begin MSAdodcLib.Adodc adoPermisos 
      Height          =   330
      Left            =   1800
      Top             =   7080
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
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
      Caption         =   "Permisos"
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
   Begin MSAdodcLib.Adodc adoAsistencia 
      Height          =   330
      Left            =   1320
      Top             =   6720
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
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
      Caption         =   "Asistencia"
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
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   5520
      TabIndex        =   28
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Reporte"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   27
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Calculo de Horas Labradas y Extras"
      Height          =   1815
      Left            =   240
      TabIndex        =   21
      Top             =   1200
      Width           =   9615
      Begin VB.TextBox txtEmpleado 
         Height          =   285
         Left            =   7680
         TabIndex        =   31
         Text            =   "%"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdHorasLaboradas 
         Caption         =   "&Calcular"
         Enabled         =   0   'False
         Height          =   495
         Left            =   7680
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpHHasta 
         Height          =   255
         Left            =   3360
         TabIndex        =   25
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   85262337
         CurrentDate     =   38612
      End
      Begin MSComCtl2.DTPicker dtpHDesde 
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   85262337
         CurrentDate     =   38612
      End
      Begin XtremeSuiteControls.ProgressBar ospHoras 
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   960
         Width           =   6135
         _Version        =   786432
         _ExtentX        =   10821
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   14737632
         Scrolling       =   1
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ProgressBar osProgress1 
         Height          =   255
         Left            =   2520
         TabIndex        =   34
         Top             =   1440
         Width           =   3855
         _Version        =   786432
         _ExtentX        =   6800
         _ExtentY        =   450
         _StockProps     =   93
         BackColor       =   14737632
         Scrolling       =   1
         Appearance      =   6
      End
      Begin VB.Label Label8 
         Caption         =   "Empleado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   32
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   " al"
         Height          =   195
         Left            =   2880
         TabIndex        =   24
         Top             =   480
         Width           =   165
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   480
         TabIndex        =   22
         Top             =   480
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Nomina"
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   9615
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   7680
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cboTipoNomina 
         Height          =   315
         Left            =   2040
         TabIndex        =   19
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblNoNomina 
         Caption         =   "172"
         Height          =   255
         Left            =   6240
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Numero:"
         Height          =   195
         Left            =   5400
         TabIndex        =   13
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Nomina"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones del Reporte"
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   9615
      Begin VB.OptionButton OptAsistenciaDpto 
         Caption         =   "Asistencia por Departamento"
         Height          =   255
         Left            =   6240
         TabIndex        =   37
         Top             =   1080
         Width           =   3015
      End
      Begin VB.OptionButton optEntradasNoRegistradas 
         Caption         =   "Entradas No Registradas"
         Height          =   255
         Left            =   3840
         TabIndex        =   36
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton OptAusencia 
         Caption         =   "Ausencia Empleados"
         Height          =   375
         Left            =   3840
         TabIndex        =   35
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optRealES 
         Caption         =   "Entrada y Salidas de Empleados"
         Height          =   255
         Left            =   6240
         TabIndex        =   29
         Top             =   360
         Width           =   3015
      End
      Begin VB.OptionButton optLaboradas 
         Caption         =   "General Laboradas y Extras"
         Height          =   255
         Left            =   6240
         TabIndex        =   14
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton optSalidasNoRegistradas 
         Caption         =   "Salidas No Registradas"
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox cboDepto 
         Height          =   315
         ItemData        =   "AsistReportes.frx":0000
         Left            =   7680
         List            =   "AsistReportes.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1800
         Width           =   1815
      End
      Begin VB.ComboBox cboCargo 
         Height          =   315
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton optCodigo 
         Caption         =   "Codigo de Empleado"
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   720
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optCargo 
         Caption         =   "Por Cargo"
         Height          =   255
         Left            =   6240
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optDepto 
         Caption         =   "Departamento"
         Height          =   255
         Left            =   6240
         TabIndex        =   8
         Top             =   1800
         Width           =   1455
      End
      Begin VB.OptionButton optFechaIngreso 
         Caption         =   "Por Fecha de Ingreso"
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton optSexo 
         Caption         =   "Por Sexo"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   255
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   85262337
         CurrentDate     =   38563
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   85262337
         CurrentDate     =   38566
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Asistencia por:"
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "al"
         Height          =   195
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   465
      End
   End
   Begin MSAdodcLib.Adodc adoInasistencia 
      Height          =   330
      Left            =   480
      Top             =   6360
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
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
      Caption         =   "Inasistencia"
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
   Begin MSAdodcLib.Adodc AdoHorario 
      Height          =   330
      Left            =   3480
      Top             =   6360
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
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
      Caption         =   "AdoHorario"
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
Attribute VB_Name = "frmRepAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ConexionRep As String, TipoNomina As String, HorasLaboradas As Double, HorasExtra As Double



Public Function ActAnt(Fecha As Date, snProd As Single) As Single

Dim fecAct As Date
Dim sTotal As String
'Dim wsWS As Workspace
'Dim dbPlanMetro As Database
'Dim rsAnt As Recordset
'Dim rsDevHora As Recordset
Dim sCad As String
Dim BuscarIncentivo As Single
Dim snTemp As Single
Dim snFactorInss As Single
Dim Cont  As Integer



Me.adoTipoNomina.CommandType = adCmdText
Me.adoTipoNomina.RecordSource = "SELECT años_acum, porcent FROM Antiguedad"
Me.adoTipoNomina.Refresh
    
'    Cont = 0
'
'    Do While Cont <= rsAnt.RecordCount - 1
'
'       snTemp = rsAnt.Fields("Ingreso")
'       rsAnt.MoveNext
'       Cont = Cont + 1
'
'    Loop
    
    
'    If Not IsNull(rsAnt.Fields("Ings")) Then
'     snTemp = rsAnt.Fields("Ings")
'    End If
'    sTotal = CStr(snProd + snTemp)
    fecAct = Format(Date, "Short Date")
'
'Set rsAnt = dbPlanMetro.OpenRecordset("Antiguedad", dbOpenDynaset)

  ' 1 Año
If fecAct - Fecha >= 365 And fecAct - Fecha <= 2 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 1 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
  ' 2 Años
    
ElseIf fecAct - Fecha >= 2 * 365 And fecAct - Fecha <= 3 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 2 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
   ' 3 Años
        
ElseIf fecAct - Fecha >= 3 * 365 And fecAct - Fecha <= 4 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 3 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
   ' 4 Años
ElseIf fecAct - Fecha >= 4 * 365 And fecAct - Fecha <= 5 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 4 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
  ' 5 Años
ElseIf fecAct - Fecha >= 5 * 365 And fecAct - Fecha <= 6 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 5 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
  ' 6 Años
ElseIf fecAct - Fecha >= 6 * 365 And fecAct - Fecha <= 7 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 6 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
   ' 7 Años
ElseIf fecAct - Fecha >= 7 * 365 And fecAct - Fecha <= 8 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 7 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
  
  ' 8 Años
ElseIf fecAct - Fecha >= 8 * 365 And fecAct - Fecha <= 9 * 365 Then
        
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 8 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
   ' 9 Años
   
ElseIf fecAct - Fecha >= 9 * 365 And fecAct - Fecha <= 10 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 9 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
  
  ' 10 Años
  
ElseIf fecAct - Fecha >= 10 * 365 And fecAct - Fecha <= 11 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 10 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
    
 ' 11 Años
ElseIf fecAct - Fecha >= 11 * 365 And fecAct - Fecha <= 12 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 11 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
  ' 12 Años
ElseIf fecAct - Fecha >= 12 * 365 And fecAct - Fecha <= 13 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 12 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
    
' 13 Años
  
ElseIf fecAct - Fecha >= 13 * 365 And fecAct - Fecha <= 14 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 13 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
    
' 14 Años
ElseIf fecAct - Fecha >= 14 * 365 And fecAct - Fecha <= 15 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 14 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
' 15 Años
ElseIf fecAct - Fecha >= 15 * 365 And fecAct - Fecha <= 16 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 15 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
 ' 16 Años
ElseIf fecAct - Fecha >= 16 * 365 And fecAct - Fecha <= 16 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 16 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
    
 ' 17 Años
ElseIf fecAct - Fecha >= 17 * 365 And fecAct - Fecha <= 18 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 17 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
    
' 18 Años
ElseIf fecAct - Fecha >= 18 * 365 And fecAct - Fecha <= 19 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 18 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
    
 ' 19 Años
 ElseIf fecAct - Fecha >= 19 * 365 And fecAct - Fecha <= 20 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 19 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
  ' 20 Años
 ElseIf fecAct - Fecha >= 20 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 20 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
 Else
   ActAnt = 0
  
End If

'sCad = "SELECT Antig, HExtras, NHextras, [Sal_ordin] from Devengado_hora WHERE Cod_Empl = " & sCodEmpl & " AND Periodo = " & iPer & " AND Año = " & iAnno
'
''Actualizo en la tabla devengado hora
'Set rsAnt = dbPlanMetro.OpenRecordset(sCad, dbOpenDynaset)
'
'If rsAnt.RecordCount > 0 Then
'  rsAnt.Edit
'  BuscarIncentivo = ((rsAnt.Fields("Sal_ordin") / 48) * 56) * snFactorInss
'  rsAnt.Fields("Antig") = Format(BuscarIncentivo, "##00.##")
'  rsAnt.Fields("HExtras") = 0
'  rsAnt.Fields("NHextras") = 0
'  rsAnt.Update
'End If
'
'sCad = "SELECT * FROM Ingreso_Empl WHERE [Cod_Empl] = " & sCodEmpl & " AND [Periodo] = " & iPer & " AND [Año] = " & iAnno & " AND [Cod_Ing] = '05'"
'Set rsAnt = dbPlanMetro.OpenRecordset(sCad, dbOpenDynaset)
'
'If rsAnt.RecordCount > 0 Then
'  rsAnt.Edit
'  rsAnt.Fields("Ingreso") = Format(BuscarIncentivo, "##00.##")
'  rsAnt.Update
'End If

End Function

Private Sub cboTipoNomina_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 And Me.cboTipoNomina.Text <> "" Then
   
Me.adoTipoNomina.CommandType = adCmdText
Me.adoTipoNomina.RecordSource = "SELECT TipoNomina.CodTipoNomina, TipoNomina.Nomina, Nomina.NumNomina, Nomina.FechaNominaINI, Nomina.FechaNomina, " & _
                                "Nomina.Activa, TipoNomina.TipoPago FROM Nomina INNER JOIN TipoNomina ON dbo.Nomina.CodTipoNomina = dbo.TipoNomina.CodTipoNomina " & _
                                "WHERE (Nomina.Activa = 1) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "'"
Me.adoTipoNomina.Refresh

Me.dtpHDesde.Value = Me.adoTipoNomina.Recordset.Fields("FechaNominaINI")
Me.dtpHHasta.Value = Me.adoTipoNomina.Recordset.Fields("FechaNomina")
Me.dtpDesde.Value = Me.adoTipoNomina.Recordset.Fields("FechaNominaINI")
Me.dtpHasta.Value = Me.adoTipoNomina.Recordset.Fields("FechaNomina")
Me.lblNoNomina.Caption = Me.adoTipoNomina.Recordset.Fields("NumNomina")
Me.cmdHorasLaboradas.Enabled = True
Me.cmdReporte.Enabled = True
TipoNomina = Me.adoTipoNomina.Recordset.Fields("TipoPago")

End If


End Sub

Private Sub cmdBuscar_Click()
cboTipoNomina_KeyDown 13, 0

Me.adoHorasExtras.CommandType = adCmdText
Me.adoHorasExtras.RecordSource = "SELECT Id, CodEmpleado, NumNomina, CantHoras, Pagada FROM HorasExtras WHERE NumNomina =" & Me.lblNoNomina.Caption
Me.adoHorasExtras.Refresh


End Sub

Private Sub cmdHorasLaboradas_Click()


Dim sngHorasLaboradas As Single
Dim sngHorasExtras As Single
Dim sngTotalHoras As Single
Dim sDia As String
Dim bUbicacion As Byte
Dim lFecha1 As Long
Dim lFecha2 As Long
Dim dFecha1 As Date
Dim dFecha2 As Date
Dim sFecha1 As String
Dim sFecha2 As String
Dim sngNumLinea As Single
Dim sngID As Single
Dim sCodEmpleado As String, sCodEmpleado1 As String
Dim sngTotalHExtras As Single
Dim sngTotalHLaboradas As Single
Dim sngPagoTotal As Single
Dim sngTarifaHoraria As Single
Dim saDias(7) As String
Dim bConta As Byte
Dim sTurno As String
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim HorasDias As Double, DiasNomina As Double


saDias(1) = "Lun"
saDias(2) = "Mart"
saDias(3) = "Mierc"
saDias(4) = "Juev"
saDias(5) = "Viern"
saDias(6) = "Sab"
saDias(7) = "Dom"
DoEvents
If Me.cboTipoNomina.Text = "" And Trim(Me.lblNoNomina.Caption) = "" Then
   Exit Sub
End If

dFecha1 = Format(Me.dtpHDesde.Value, "yyyy-mm-dd")
dFecha2 = Me.dtpHHasta.Value
lFecha1 = dFecha1
lFecha2 = dFecha2
sFecha1 = Mid$(Me.dtpHDesde.Value, 7, 4) & "-" & Mid$(Me.dtpHDesde.Value, 4, 2) & "-" & Mid$(Me.dtpHDesde.Value, 1, 2)
sFecha2 = Mid$(Me.dtpHHasta.Value, 7, 4) & "-" & Mid$(Me.dtpHHasta.Value, 4, 2) & "-" & Mid$(Me.dtpHHasta.Value, 1, 2)

'''///////////////////////////////////////////////////////////////////////////////////////////////////
'''///////////////////CALCULO SI ES NOMINA SEMANAL O CATORCENAL /////////////////////////////////////
'''/////////////////////////////////////////////////////////////////////////////////////////////////
If DateDiff("d", sFecha1, sFecha2) <= 7 Then
  HorasDias = 9.75
  DiasNomina = 7
ElseIf DateDiff("d", sFecha1, sFecha2) <= 14 Then
  HorasDias = 19.5
  DiasNomina = 14
End If


Me.adoAsistencia.CommandType = adCmdText
Me.adoAsistencia.RecordSource = "SELECT AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.CodEmpleado1, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, " & _
                                "AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.CodTurno, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.Dia, " & _
                                "AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso, TipoNomina.Nomina FROM AsistenciaEmpleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                                "WHERE     (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND (TipoNomina.Nomina = '" & Me.cboTipoNomina.Text & " ') AND AsistenciaEmpleado.CodEmpleado1 LIKE '" & Me.txtEmpleado.Text & "' ORDER BY AsistenciaEmpleado.CodEmpleado1, AsistenciaEmpleado.FechaEntrada"
    '                            "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaSalida <> Null AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' ORDER BY AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC "
Me.adoAsistencia.Refresh

Me.adoTipoNomina.CommandType = adCmdText
Me.adoTipoNomina.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion " & _
                                 "WHERE DetalleHorasProduccion.NumNomina =" & Me.lblNoNomina.Caption
Me.adoTipoNomina.Refresh


Me.adoConsulta.CommandType = adCmdText
Me.adoConsulta.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion"
Me.adoConsulta.Refresh



'Me.adoConsulta.CommandType = adCmdText
'Me.adoConsulta.RecordSource = "SELECT  MAX(DetalleHorasProduccion.NumLinea) AS [MaximaLinea] FROM DetalleHorasProduccion " & _
'                               "WHERE DetalleHorasProduccion.NumNomina =" & Me.lblNoNomina.Caption
'Me.adoConsulta.Refresh

Me.adoTurno.CommandType = adCmdText
Me.adoTurno.RecordSource = "SELECT * FROM Turno"
Me.adoTurno.Refresh

Me.ospHoras.Visible = True
Me.ospHoras.Min = 0
Me.ospHoras.Max = Me.adoAsistencia.Recordset.RecordCount
Me.ospHoras.Value = 0
Me.ospHoras.Visible = True

If Not Me.adoAsistencia.Recordset.EOF Then
   sCodEmpleado = Me.adoAsistencia.Recordset.Fields("CodEmpleado")
End If


If Len(Me.txtEmpleado.Text) <> 6 Then
     sCodEmpleado = "%"
End If


sTurno = Me.adoAsistencia.Recordset.Fields("CodTurno")

Do While Not Me.adoAsistencia.Recordset.EOF
  DoEvents
  Me.Caption = "Procesando " & Me.adoAsistencia.Recordset.Fields("CodEmpleado1")
  Me.ospHoras.Value = Me.ospHoras.Value + 1
  
'  If Me.adoAsistencia.Recordset.Fields("CodEmpleado1") = "S116020057" Then
'     sCodEmpleado = Me.adoAsistencia.Recordset.Fields("CodEmpleado")
'  End If
     
  If Trim(Me.txtEmpleado.Text) = "%" Then
     sCodEmpleado = "%"
     sCodEmpleado1 = "%"
  Else
     sCodEmpleado = Me.adoAsistencia.Recordset.Fields("CodEmpleado")
     sCodEmpleado1 = Me.adoAsistencia.Recordset.Fields("CodEmpleado1")
  End If
     
  sFecha1 = Mid$(Me.adoAsistencia.Recordset.Fields("FechaEntrada"), 7, 4) & "-" & Mid$(Me.adoAsistencia.Recordset.Fields("FechaEntrada"), 4, 2) & "-" & Mid$(Me.adoAsistencia.Recordset.Fields("FechaEntrada"), 1, 2)
  
  Me.adoPermisos.CommandType = adCmdText
  Me.adoPermisos.RecordSource = "SELECT * FROM Permisos WHERE CodEmpleado ='" & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & "' AND Fecha =CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND RegresoPendiente =0 AND Justificado =0"
  Me.adoPermisos.Refresh
  
  Me.AdoHorario.RecordSource = "SELECT CodEmpleado, DATEDIFF(hour, LEntrada, LSalida) - TComida / 60 AS Lunes, DATEDIFF(hour, MEntrada, MSalida) - TComida / 60 AS Martes, DATEDIFF(hour, MCEntrada, MCSalida) - TComida / 60 AS Miercoles, DATEDIFF(hour, JEntrada, JSalida) - TComida / 60 AS Jueves, DATEDIFF(hour, VEntrada, VSalida) - TComida / 60 AS Viernes, TComida, TurnoLunes, TurnoMartes, TurnoMiercoles, TurnoJueves, TurnoViernes, TurnoSabado, TurnoDomingo, DATEDIFF(hour,SEntrada, SSalida) AS Sabado, DATEDIFF(hour, DEntrada, DSalida) AS Domingo From HorarioEmpleado " & _
                               "WHERE (CodEmpleado = '" & sCodEmpleado1 & "')"
  Me.AdoHorario.Refresh
  
    
  
  If Me.adoAsistencia.Recordset.Fields("FechaEntrada") = Me.adoAsistencia.Recordset.Fields("FechaSalida") Then
      Me.adoTurno.CommandType = adCmdText
      Me.adoTurno.RecordSource = "SELECT * FROM Turno WHERE CodTurno ='" & sTurno & "'"
      Me.adoTurno.Refresh
      
      bUbicacion = InStr(1, Format(Me.adoAsistencia.Recordset.Fields("FechaEntrada"), "Long Date"), " ", vbTextCompare)
      sDia = UCase(Mid$(Format(Me.adoAsistencia.Recordset.Fields("FechaEntrada"), "Long Date"), 1, bUbicacion - 2))
      
            
            
      Select Case sDia
     
      Case "LUNES":
          
          sDia = "Lun"
          CalcularLaboradas (sDia)
            sngHorasLaboradas = HorasLaboradas
            sngHorasExtras = HorasExtra
          
'          sDia = "Lun"
'
'          If Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss") <= Format(Me.adoTurno.Recordset.Fields("LEntrada"), "hh:mm:ss") Then
'            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:00:00" Then
'              sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
'              sngHorasExtras = 0
'            Else
'              sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoTurno.Recordset.Fields("LSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
'              sngHorasExtras = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) - sngHorasLaboradas - (Me.adoTurno.Recordset.Fields("TComida") / 60)
'            End If
'
'
'
'         ElseIf (Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:00:00" And Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss") >= "06:30:00" And Format(Me.adoTurno.Recordset.Fields("LSalida"), "hh:mm:ss") >= Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) Then  ' Or (Me.adoAsistencia.Recordset.Fields("HoraSalida") < "23:59:59" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "17:00:00") Then
'             sngHorasLaboradas = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("LEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
'             sngHorasExtras = 0
'
'          Else
'             sngHorasLaboradas = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("LEntrada"), "hh:mm:ss"), Format(Me.adoTurno.Recordset.Fields("LSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
'             sngHorasExtras = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("LSalida"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
'
'          End If
'
'          If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= Format(Me.adoTurno.Recordset.Fields("LSalida"), "hh:mm:ss") And Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") > "12:00:00" Then
'             sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
'             sngHorasExtras = 0
'
'          Else
'
'            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:00:00" Then
'              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
'              sngHorasExtras = 0
'            Else
'              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Format(Me.adoTurno.Recordset.Fields("LSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
'              ' sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - sngHorasLaboradas
'            End If
'
'          End If
'
'          If sngHorasLaboradas <= 0 Then
'            sngHorasLaboradas = 0
'            sngHorasExtras = 0
'          End If
'

           

            
       Case "MARTES":
                  
          sDia = "Mart"
          
             
             
             
          If Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss") <= Format(Me.adoTurno.Recordset.Fields("MEntrada"), "hh:mm:ss") Then
            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:15:00" Then
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
              sngHorasExtras = 0
            Else
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoTurno.Recordset.Fields("MSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) - sngHorasLaboradas - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            End If
          ElseIf (Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:15:00" And Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss") >= "06:30:00" And Format(Me.adoTurno.Recordset.Fields("MSalida"), "hh:mm:ss") >= Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) Then ' Or (Me.adoAsistencia.Recordset.Fields("HoraSalida") < "23:59:59" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "17:00:00") Then
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("MEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
             sngHorasExtras = 0
             
          Else
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("MEntrada"), "hh:mm:ss"), Format(Me.adoTurno.Recordset.Fields("MSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Me.adoTurno.Recordset.Fields("MSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
          End If
           
          If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= Format(Me.adoTurno.Recordset.Fields("MSalida"), "hh:mm:ss") And Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") > "12:00:00" Then
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = 0
          Else
            
            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:15:00" Then
              sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
              sngHorasExtras = 0
            Else
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoTurno.Recordset.Fields("MSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             ' sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - sngHorasLaboradas
            End If
             
          End If
            
          If sngHorasLaboradas <= 0 Then
            sngHorasLaboradas = 0
            sngHorasExtras = 0
          End If
           
        Case "MIÉRCOLES":
                  
          sDia = "Mierc"
          
             
          
          If Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss") <= Format(Me.adoTurno.Recordset.Fields("MCEntrada"), "hh:mm:ss") Then
            
            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= Format(Me.adoTurno.Recordset.Fields("MCSalida"), "hh:mm:ss") Then
             
             If (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) < 5 Then
                sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
                sngHorasExtras = 0
             Else
                sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
                sngHorasExtras = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) - sngHorasLaboradas - (Me.adoTurno.Recordset.Fields("TComida") / 60)

             End If
                
           End If
             
          ElseIf (Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:15:00" And Format(Me.adoTurno.Recordset.Fields("MCSalida"), "hh:mm:ss") >= Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) Then  ' ) Or (Me.adoAsistencia.Recordset.Fields("HoraSalida") < "23:59:59" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "17:00:00") Then
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("MCEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
             sngHorasExtras = 0
             
          ElseIf (Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= Format(Me.adoTurno.Recordset.Fields("MCSalida"), "hh:mm:ss")) Then
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("MCEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60 - (Me.adoTurno.Recordset.Fields("TComida") / 60))
             sngHorasExtras = 0
          
          Else
          
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("MCEntrada"), "hh:mm:ss"), Format(Me.adoTurno.Recordset.Fields("MCSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("MCSalida"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
             
           End If
           
          If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= Format(Me.adoTurno.Recordset.Fields("MCSalida"), "hh:mm:ss") And Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") > "12:00:00" Then
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = 0
             
          Else
          
            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:15:00" Then
              sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
              'Me.adoTipoNomina.Recordset.Fields("NumLinea") = sngNumLinea
              sngHorasExtras = 0
            ElseIf Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") > Format(Me.adoTurno.Recordset.Fields("MCSalida"), "hh:mm:ss") Then
            
              sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoTurno.Recordset.Fields("MCSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
              sngHorasExtras = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("MCSalida"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)  ' - sngHorasLaboradas
              '////////////////////////REDONDEO A HORAS ENTEROS /////////////////////////
              
            End If
             
          End If
          
          If sngHorasLaboradas <= 0 Then
            sngHorasLaboradas = 0
            sngHorasExtras = 0
          End If
           
       Case "JUEVES":
                  
          sDia = "Juev"
          
            
          
          
          If Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss") <= Format(Me.adoTurno.Recordset.Fields("JEntrada"), "hh:mm:ss") Then
            
            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:15:00" Then
              sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
              sngHorasExtras = 0
            ElseIf Format(Me.adoTurno.Recordset.Fields("JSalida"), "hh:mm:ss") <= Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") And (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoTurno.Recordset.Fields("JSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60) > 0 Then
              sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoTurno.Recordset.Fields("JSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
              sngHorasExtras = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) - sngHorasLaboradas - (Me.adoTurno.Recordset.Fields("TComida") / 60)

            End If
          
          ElseIf (Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") < "12:15:00" And Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss") >= "06:30:00" And Format(Me.adoTurno.Recordset.Fields("JSalida"), "hh:mm:ss") >= Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) Then   ') Or (Me.adoAsistencia.Recordset.Fields("HoraSalida") < "23:59:59" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "17:00:00") Then
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("JEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
             sngHorasExtras = 0
             
          ElseIf Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss") >= Format(Me.adoTurno.Recordset.Fields("JEntrada"), "hh:mm:ss") And Format(Me.adoTurno.Recordset.Fields("JSalida"), "hh:mm:ss") <= Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") Then
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoTurno.Recordset.Fields("JSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("JSalida"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
          
          Else
          
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("JEntrada"), "hh:mm:ss"), Format(Me.adoTurno.Recordset.Fields("JSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("JSalida"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
             
          End If
           
          If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= Format(Me.adoTurno.Recordset.Fields("JSalida"), "hh:mm:ss") And Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") > "12:15:00" Then
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = 0
             
          Else
            
            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:15:00" Then
              sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
              sngHorasExtras = 0
            Else
              sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoTurno.Recordset.Fields("JSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
              
           '   sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) ' - sngHorasLaboradas
            End If
             
          End If
          
          If sngHorasLaboradas <= 0 Then
            sngHorasLaboradas = 0
            sngHorasExtras = 0
          End If
           
        Case "VIERNES":
          
           sDia = "Viern"
          
             
          
          If Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss") <= Format(Me.adoTurno.Recordset.Fields("VEntrada"), "hh:mm:ss") Then
            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:15:00" Then
              sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
              sngHorasExtras = 0
            Else
             sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoTurno.Recordset.Fields("VSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) - sngHorasLaboradas - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            End If
             
           ElseIf (Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") < "12:15:00" And Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss") >= "06:30:00" And Format(Me.adoTurno.Recordset.Fields("VSalida"), "hh:mm:ss") >= Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) Then ' ) Or (Me.adoAsistencia.Recordset.Fields("HoraSalida") < "23:59:59" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "17:00:00") Then
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("VEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
             sngHorasExtras = 0
             
           Else
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("VEntrada"), "hh:mm:ss"), Format(Me.adoTurno.Recordset.Fields("VSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("VSalida"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
           End If
           
          If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= Format(Me.adoTurno.Recordset.Fields("VSalida"), "hh:mm:ss") And Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") > "12:00:00" Then
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = 0
             
          Else
          
            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:15:00" Then
              sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
              sngHorasExtras = 0
            Else
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoTurno.Recordset.Fields("VSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
           '  sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - sngHorasLaboradas
            End If
             
          End If
          
          If sngHorasLaboradas <= 0 Then
            sngHorasLaboradas = 0
            sngHorasExtras = 0
          End If
           
        Case Else:
                  
             If sDia = "SÁBADO" Then
                sDia = "Sab"
                
                 
                 
             Else
               sDia = "Dom"
             End If
             
               
          
           If (Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:00:00" And Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss") >= "05:00:00") Or (Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") < "23:59:59" And Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss") >= "17:00:00") Then
              sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
           ElseIf (Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") >= "12:00:00" And Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:45:00") Then
              sngHorasExtras = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), "12:00:00") / 60)
           Else
              sngHorasExtras = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
           End If
             'Me.adoAsistencia.Recordset.Fields("HLaboradas") = 0
             sngHorasLaboradas = 0
            
            
      End Select
      
        If sDia <> "Sab" Then
         If sDia <> "Dom" Then
          If sngHorasLaboradas <= 0 Then
            sngHorasLaboradas = 0
            sngHorasExtras = 0
           End If
          End If
        End If
  Else
  
  
      Me.adoTurno.CommandType = adCmdText
      Me.adoTurno.RecordSource = "SELECT * FROM Turno WHERE CodTurno ='" & sTurno & "'"
      Me.adoTurno.Refresh
     
      bUbicacion = InStr(1, Format(Me.adoAsistencia.Recordset.Fields("FechaEntrada"), "Long Date"), " ", vbTextCompare)
      sDia = Mid$(Format(Me.adoAsistencia.Recordset.Fields("FechaEntrada"), "Long Date"), 1, bUbicacion - 2)
  
      Select Case sDia
      
      Case "SÁBADO":
         
         sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida") / 60)) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
         sngHorasLaboradas = 0
         
         
         
  
      Case "DOMINGO":
         
         sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida") / 60)) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
         sngHorasLaboradas = 0
        
               
      Case "LUNES":
         
            
         
         If Format(Me.adoTurno.Recordset.Fields("LSalida"), "hh:mm:ss") <= Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") Then
            sngHorasLaboradas = Format((DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Me.adoTurno.Recordset.Fields("LSalida")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            sngHorasExtras = DateDiff("n", Me.adoTurno.Recordset.Fields("LSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60
         Else
             
            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "00:00:00" Then
               sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
               sngHorasExtras = 0
            Else
               sngHorasLaboradas = Format((DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
               sngHorasExtras = 0
            End If
         
         End If
         
            
         
         
       Case "MARTES":
         
          
         
         If Format(Me.adoTurno.Recordset.Fields("MSalida"), "hh:mm:ss") <= Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") Then
            sngHorasLaboradas = Format((DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Format(Me.adoTurno.Recordset.Fields("MSalida"), "hh:mm:ss")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            sngHorasExtras = DateDiff("n", Me.adoTurno.Recordset.Fields("MSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60
         Else
             
            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "00:00:00" Then
               sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
               sngHorasExtras = 0
            Else
               sngHorasLaboradas = Format((DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
               sngHorasExtras = 0
            End If
         
         End If
         
       
         
         
       Case "MIÉRCOLES":
         
            
         
         If Format(Me.adoTurno.Recordset.Fields("MCSalida"), "hh:mm:ss") <= Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") Then
            sngHorasLaboradas = Format((DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Me.adoTurno.Recordset.Fields("MCSalida")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            sngHorasExtras = DateDiff("n", Me.adoTurno.Recordset.Fields("MCSalida"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60
         Else
             
            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "00:00:00" Then
               sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
               sngHorasExtras = 0
            Else
               sngHorasLaboradas = Format((DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
               sngHorasExtras = 0
            End If
         
         End If
         
      Case "JUEVES":
         
          
         
         If Format(Me.adoTurno.Recordset.Fields("JSalida"), "hh:mm:ss") <= Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") Then
            sngHorasLaboradas = Format((DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada")), "23:59:59") / 60) + DateDiff("n", "00:00:00", Format(Me.adoTurno.Recordset.Fields("JSalida"), "hh:mm:ss")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            sngHorasExtras = DateDiff("n", Format(Me.adoTurno.Recordset.Fields("JSalida"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60
         Else
             
            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "00:00:00" Then
               sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
               sngHorasExtras = 0
            Else
               sngHorasLaboradas = Format((DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
               sngHorasExtras = 0
            End If
         
         End If
       
       Case "VIERNES":
         
         
         
         If Format(Me.adoTurno.Recordset.Fields("VSalida"), "hh:mm:ss") <= Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") Then
            sngHorasLaboradas = Format((DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Format(Me.adoTurno.Recordset.Fields("VSalida"), "hh:mm:ss")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            sngHorasExtras = DateDiff("n", Format(Me.adoTurno.Recordset.Fields("VSalida"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
         Else
             
            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "00:00:00" Then
               sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
               sngHorasExtras = 0
            Else
               sngHorasLaboradas = Format((DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
               sngHorasExtras = 0
            End If
         
         End If
             
                   
      End Select
       
  
     
  End If
  
  
  
  If Not Me.adoPermisos.Recordset.EOF Then
     sngHorasLaboradas = sngHorasLaboradas - (DateDiff("n", Me.adoPermisos.Recordset.Fields("HoraInicio"), Me.adoPermisos.Recordset.Fields("HoraFin")) / 60)
     Me.adoAsistencia.Recordset.Fields("bPermiso") = 1
  End If
         
  Me.adoPermisos.CommandType = adCmdText
  Me.adoPermisos.RecordSource = "SELECT * FROM ExtraTurno WHERE CodEmpleado ='" & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & "' AND FechaEntrada =CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND (bActivo = 0) AND (NOT (HoraSalida IS NULL)) AND (NOT (FechaEntrada IS NULL))"
  Me.adoPermisos.Refresh
          
  If Not Me.adoPermisos.Recordset.EOF And Not IsNull(Me.adoPermisos.Recordset.Fields("HorasLaboradas")) Then
     sngHorasExtras = sngHorasExtras + Me.adoPermisos.Recordset.Fields("HorasLaboradas")
  End If
       
  If sngHorasLaboradas <> 0 Then
     sngHorasLaboradas = Format(sngHorasLaboradas, "##,##0.00")
  Else
     sngHorasLaboradas = 0
  End If
  
  If sngHorasExtras > 1 Then
     sngHorasExtras = Format(sngHorasExtras, "##,##0.00")
  Else
     sngHorasExtras = 0
  End If
  
  
  Me.adoAsistencia.Recordset.Fields("Dia") = sDia
  Me.adoAsistencia.Recordset.Fields("HLaboradas") = sngHorasLaboradas
  Me.adoAsistencia.Recordset.Fields("HExtras") = sngHorasExtras
  Me.adoAsistencia.Recordset.Update
  
  sngHorasExtras = 0
  sngHorasLaboradas = 0
  
  Me.adoAsistencia.Recordset.MoveNext
  DoEvents
   
Loop
     
sFecha1 = Mid$(Me.dtpHDesde.Value, 7, 4) & "-" & Mid$(Me.dtpHDesde.Value, 4, 2) & "-" & Mid$(Me.dtpHDesde.Value, 1, 2)
sFecha2 = Mid$(Me.dtpHHasta.Value, 7, 4) & "-" & Mid$(Me.dtpHHasta.Value, 4, 2) & "-" & Mid$(Me.dtpHHasta.Value, 1, 2)

Me.adoAsistencia.CommandType = adCmdText
Me.adoAsistencia.RecordSource = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.TarifaHoraria, Empleado.SalarioMinimo, TipoNomina.Nomina, Empleado.Activo FROM Empleado INNER JOIN TipoNomina ON dbo.Empleado.CodTipoNomina = dbo.TipoNomina.CodTipoNomina " & _
                                "WHERE (TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "') AND (dbo.Empleado.Activo = 1) AND Empleado.CodEmpleado LIKE '" & sCodEmpleado & "' ORDER BY CodEmpleado1 ASC"
                                
Me.adoAsistencia.Refresh
   
Me.adoTipoNomina.CommandType = adCmdText
Me.adoTipoNomina.RecordSource = "SELECT AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, " & _
                                "AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.CodTurno, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.Dia, " & _
                                "AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso, TipoNomina.Nomina FROM AsistenciaEmpleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                                "WHERE (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND (TipoNomina.Nomina = '" & Me.cboTipoNomina.Text & " ') ORDER BY AsistenciaEmpleado.CodEmpleado ASC"
Me.adoTipoNomina.Refresh
   
   
Me.ospHoras.Value = 0
Me.ospHoras.Min = 0
Me.ospHoras.Max = Me.adoAsistencia.Recordset.RecordCount
Me.ospHoras.Value = 0


bConta = 1
     
     
     
Do While Not Me.adoAsistencia.Recordset.EOF
  
  DoEvents
  Me.Caption = "Procesando " & Me.adoAsistencia.Recordset.Fields("CodEmpleado1")
  Me.ospHoras.Value = Me.ospHoras.Value + 1
  
  Me.adoConsulta.CommandType = adCmdText
  Me.adoConsulta.RecordSource = "SELECT Max(Id) AS [MaximaId] FROM HorasExtras"
  Me.adoConsulta.Refresh
    
  If Not IsNull(Me.adoConsulta.Recordset.Fields("MaximaId")) Then
     sngID = Me.adoConsulta.Recordset.Fields("MaximaId") + 1
  Else
     sngID = 1
  End If
  
 
  
  Me.adoTipoNomina.CommandType = adCmdText
  Me.adoTipoNomina.RecordSource = "SELECT Sum(AsistenciaEmpleado.HExtras) AS [SumaExtras] FROM AsistenciaEmpleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                                 "WHERE (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND (TipoNomina.Nomina = '" & Me.cboTipoNomina.Text & "') " & _
                                 "AND CodEmpleado ='" & Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado")) & "'"
  Me.adoTipoNomina.Refresh
  
  
  If Not IsNull(Me.adoTipoNomina.Recordset.Fields("SumaExtras")) Then
    If Me.adoTipoNomina.Recordset.Fields("SumaExtras") <> 0 Then
     sngTotalHExtras = Format(Me.adoTipoNomina.Recordset.Fields("SumaExtras"), "##,##0.00")
    Else
     sngTotalHExtras = 0
    End If
    
  Else
     sngTotalHExtras = 0
  End If
  
 
  Me.adoConsulta.CommandType = adCmdText
  Me.adoConsulta.RecordSource = "SELECT Id, CodEmpleado, NumNomina, CantHoras, Pagada FROM HorasExtras " & _
                               "WHERE NumNomina =" & Me.lblNoNomina.Caption & " AND CodEmpleado ='" & Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado")) & "'"
  Me.adoConsulta.Refresh
            
  If Not Me.adoConsulta.Recordset.EOF Then
    If sngTotalHExtras <> 0 Then
     Me.adoConsulta.Recordset.Fields("CantHoras") = Format(sngTotalHExtras, "##,##0.00")
     Me.adoConsulta.Recordset.Fields("Pagada") = 0
     Me.adoConsulta.Recordset.Update
    Else
     Me.adoConsulta.Recordset.Fields("CantHoras") = 0
     Me.adoConsulta.Recordset.Fields("Pagada") = 0
     Me.adoConsulta.Recordset.Update
    
    End If
    
     Me.adoConsulta.Refresh
     
  ElseIf sngTotalHExtras <> 0 Then
     Me.adoConsulta.Recordset.AddNew
     Me.adoConsulta.Recordset.Fields("Id") = sngID
     Me.adoConsulta.Recordset.Fields("CodEmpleado") = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
     Me.adoConsulta.Recordset.Fields("NumNomina") = Me.lblNoNomina.Caption
     If sngTotalHExtras <> 0 Then
        Me.adoConsulta.Recordset.Fields("CantHoras") = Format(sngTotalHExtras, "##,##0.00")
     Else
        Me.adoConsulta.Recordset.Fields("CantHoras") = 0
     End If
     
     Me.adoConsulta.Recordset.Fields("Pagada") = 0
     Me.adoConsulta.Recordset.Update
     Me.adoConsulta.Refresh
            
  Else
  
     Me.adoConsulta.Recordset.AddNew
     Me.adoConsulta.Recordset.Fields("Id") = sngID
     Me.adoConsulta.Recordset.Fields("CodEmpleado") = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
     Me.adoConsulta.Recordset.Fields("NumNomina") = Me.lblNoNomina.Caption
     
     Me.adoConsulta.Recordset.Fields("CantHoras") = 0
    
     
     Me.adoConsulta.Recordset.Fields("Pagada") = 0
     Me.adoConsulta.Recordset.Update
     Me.adoConsulta.Refresh
            
  End If
  
Me.adoAsistencia.Recordset.MoveNext
DoEvents
Loop




Me.adoAsistencia.CommandType = adCmdText
'Me.adoAsistencia.RecordSource = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.TarifaHoraria, TipoNomina.Nomina, AsistenciaEmpleado.FechaEntrada, " & _
'                                "AsistenciaEmpleado.FechaSalida FROM AsistenciaEmpleado INNER JOIN Empleado ON dbo.AsistenciaEmpleado.CodEmpleado = dbo.Empleado.CodEmpleado INNER JOIN " & _
'                                "TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
'                                 "WHERE (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND (TipoNomina.Nomina = '" & Me.cboTipoNomina.Text & " ') AND Empleado.CodEmpleado LIKE '" & sCodEmpleado & "' ORDER BY AsistenciaEmpleado.CodEmpleado1 ASC"
Me.adoAsistencia.RecordSource = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.TarifaHoraria, TipoNomina.Nomina, Empleado.Activo FROM Empleado INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina  " & _
                                "WHERE  (TipoNomina.Nomina = '" & Me.cboTipoNomina.Text & " ') AND (Empleado.CodEmpleado LIKE '" & sCodEmpleado & "') AND (Empleado.Activo = 1)"
Me.adoAsistencia.Refresh
   
Me.adoTipoNomina.CommandType = adCmdText
Me.adoTipoNomina.RecordSource = "SELECT AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, " & _
                                "AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.CodTurno, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.Dia, " & _
                                "AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso FROM AsistenciaEmpleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                                "WHERE (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) ORDER BY AsistenciaEmpleado.CodEmpleado ASC"
Me.adoTipoNomina.Refresh
   
Me.adoHorasExtras.CommandType = adCmdText
Me.adoHorasExtras.RecordSource = "SELECT AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, " & _
                                "AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.CodTurno, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.Dia, " & _
                                "AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso FROM AsistenciaEmpleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                                "WHERE (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) ORDER BY AsistenciaEmpleado.CodEmpleado ASC"
Me.adoHorasExtras.Refresh
   
   
   
   
Me.ospHoras.Value = 0
Me.ospHoras.Min = 0
Me.ospHoras.Max = Me.adoAsistencia.Recordset.RecordCount + 1
Me.ospHoras.Value = 0

'If Me.cboTipoNomina.Text <> "Produccion" Then

'/////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////ELIMINO LOS REGISTROS DE HORAS PARA ESTE TRABAJADOR ///////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////

If TipoNomina = "Salario Fijo" Or TipoNomina = "Salario Destajo" Or TipoNomina = "Salario Fijo,Destajo y Comision" Or TipoNomina = "Salario Destajo y Comision" Then

Do While Not Me.adoAsistencia.Recordset.EOF
DoEvents
Me.Caption = "Procesando " & Me.adoAsistencia.Recordset.Fields("CodEmpleado1")

sFecha1 = Format(Me.dtpHDesde.Value, "yyyy-mm-dd")
sFecha2 = Format(Me.dtpHHasta.Value, "yyyy-mm-dd")

sCodEmpleado = Me.adoAsistencia.Recordset.Fields("CodEmpleado")

 If TipoNomina = "Salario Fijo" Then
   sngTarifaHoraria = BuscaTarifaHoraria(sCodEmpleado)
 End If

bConta = 1

  

'  Me.AdoHorario.RecordSource = "SELECT CodEmpleado, DATEDIFF(hour, LEntrada, LSalida) - TComida / 60 AS Lunes, DATEDIFF(hour, MEntrada, MSalida) - TComida / 60 AS Martes, DATEDIFF(hour, MCEntrada, MCSalida) - TComida / 60 AS Miercoles, DATEDIFF(hour, JEntrada, JSalida) - TComida / 60 AS Jueves, DATEDIFF(hour, VEntrada, VSalida) - TComida / 60 AS Viernes, TComida, TurnoLunes, TurnoMartes, TurnoMiercoles, TurnoJueves, TurnoViernes, TurnoSabado, TurnoDomingo, DATEDIFF(hour,SEntrada, SSalida) AS Sabado, DATEDIFF(hour, DEntrada, DSalida) AS Domingo From HorarioEmpleado " & _
'                               "WHERE (CodEmpleado = '" & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & "')"
Me.AdoHorario.RecordSource = "SELECT CodEmpleado, (DATEDIFF(minute, LEntrada, LSalida) - TComida) / 60.0 AS Lunes, (DATEDIFF(minute, MEntrada, MSalida) - TComida) / 60.0 AS Martes, (DATEDIFF(minute, MCEntrada, MCSalida) - TComida) / 60.0 AS Miercoles, (DATEDIFF(minute, JEntrada, JSalida) - TComida) / 60.0 AS Jueves, (DATEDIFF(minute, VEntrada, VSalida) - TComida) / 60.0 AS Viernes, TComida, TurnoLunes, TurnoMartes, TurnoMiercoles, TurnoJueves, TurnoViernes, TurnoSabado, TurnoDomingo, DATEDIFF(hour,SEntrada, SSalida) AS Sabado, DATEDIFF(hour, DEntrada, DSalida) AS Domingo From HorarioEmpleado " & _
                               "WHERE (CodEmpleado = '" & Me.adoAsistencia.Recordset.Fields("CodEmpleado1") & "')"
  
  Me.AdoHorario.Refresh

Me.adoHorasExtras.CommandType = adCmdText
'Me.adoHorasExtras.RecordSource = "SELECT AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, " & _
'                                "AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.CodTurno, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.Dia, " & _
'                                "AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso FROM AsistenciaEmpleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
'                                "WHERE (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) " & _
'                                "AND AsistenciaEmpleado.CodEmpleado ='" & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & "' ORDER BY AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada"
Me.adoHorasExtras.RecordSource = "SELECT DISTINCT AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.CodTurno, AsistenciaEmpleado.HLaboradas , AsistenciaEmpleado.Dia, AsistenciaEmpleado.HExtras, AsistenciaEmpleado.bPermiso FROM AsistenciaEmpleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina  " & _
                                 "WHERE (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND (AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102)) AND (AsistenciaEmpleado.CodEmpleado = '" & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & "') AND (AsistenciaEmpleado.CodEmpleado1 = '" & Me.adoAsistencia.Recordset.Fields("CodEmpleado1") & "') ORDER BY AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada"
Me.adoHorasExtras.Refresh

sngTotalHLaboradas = 0
sngPagoTotal = 0
Me.ospHoras.Value = Me.ospHoras.Value + 1

Dim dLunes, dMartes, dMiercoles, dJueves, dViernes As Double

dLunes = 0
dMartes = 0
dMiercoles = 0
dJueves = 0
dViernes = 0

 Me.osProgress1.Min = 0
 Me.osProgress1.Max = 7
 Me.osProgress1.Value = 0
 Me.osProgress1.Visible = True
 
 '------------------------------------------------------------------------------------
 '-------------LLENO DE CEROS LA SEMANA ANTES DE GRABARLA -----------------------------
 '-----------------------------------------------------------------------------------------
 rs.Open "UPDATE [DetalleHorasProduccion] Set [Lunes] = 0,[Martes] = 0,[Miercoles] = 0,[Jueves] = 0,[Viernes] = 0,[Sabado] = 0,[Domingo] = 0,[TotalHoras] = 0,[SalarioHora] = 0,[TotalSalarioHora] = 0 Where (CodEmpleado = " & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & ") And (NumNomina = " & Me.lblNoNomina.Caption & ")", Conexion
 
'Do While bConta <= 7 And Not Me.adoHorasExtras.Recordset.EOF


'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////ELIMINO LOS REGISTROS DE ESTE EMPLEADO PARA LA NOMINA ////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'rs.Open "DELETE FROM DetalleHorasProduccion WHERE (DetalleHorasProduccion.CodEmpleado = '" & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & "') AND (DetalleHorasProduccion.Pagado = 0) AND (DetalleHorasProduccion.NumNomina = " & Trim(Me.lblNoNomina.Caption) & ")", Conexion

'//////////////////////////////////




Do While bConta <= DiasNomina And Not Me.adoHorasExtras.Recordset.EOF
  DoEvents
  Me.adoConsulta.CommandType = adCmdText
  Me.adoConsulta.RecordSource = "SELECT Max(NumLinea) AS [MaximaLinea] FROM DetalleHorasProduccion"
  Me.adoConsulta.Refresh
    
'  If Not IsNull(Me.AdoConsulta.Recordset.Fields("MaximaLinea")) Then
'     sngNumLinea = Me.AdoConsulta.Recordset.Fields("MaximaLinea") + 1
'  Else
'     sngNumLinea = 1
'  End If
  
  If bConta <= 5 Then
    sngNumLinea = 1
    sFecha2 = Format(DateAdd("d", 6, sFecha1), "yyyy-mm-dd")
    
      If Format(Me.adoHorasExtras.Recordset("FechaEntrada"), "yyyy-mm-dd") > sFecha2 Then
        sngNumLinea = 2
        sFecha1 = Format(DateAdd("d", 7, sFecha1), "yyyy-mm-dd")
        sFecha2 = Format(DateAdd("d", 6, sFecha1), "yyyy-mm-dd")
        bConta = 6
        dLunes = 0
        dMartes = 0
        dMiercoles = 0
        dJueves = 0
        dViernes = 0
     End If
  ElseIf bConta = 6 Then
    sngNumLinea = 2
    sFecha1 = Format(DateAdd("d", 7, sFecha1), "yyyy-mm-dd")
    sFecha2 = Format(DateAdd("d", 6, sFecha1), "yyyy-mm-dd")
     dLunes = 0
     dMartes = 0
     dMiercoles = 0
     dJueves = 0
     dViernes = 0
  End If
  

  
  
  
  If Not Me.adoAsistencia.Recordset.EOF Then

  

'     Me.adoTipoNomina.CommandType = adCmdText
     Me.adoTipoNomina.RecordSource = "SELECT Sum(AsistenciaEmpleado.Hlaboradas) AS [SumaLaboradas] FROM AsistenciaEmpleado " & _
                                  "WHERE (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND CodEmpleado ='" & Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado")) & "' " & _
                                  "AND Dia ='" & Me.adoHorasExtras.Recordset.Fields("Dia") & "'"
     Me.adoTipoNomina.Refresh
     

  
     
    
    Select Case Me.adoHorasExtras.Recordset.Fields("Dia")
  
    Case "Lun":
    
       Me.AdoHorario.Refresh
       If Not Me.AdoHorario.Recordset.EOF Then
         If DateDiff("d", sFecha1, sFecha2) <= 7 Then
           HorasDias = Me.AdoHorario.Recordset("Lunes")
         ElseIf DateDiff("d", sFecha1, sFecha2) <= 14 Then
            HorasDias = CDbl(Me.AdoHorario.Recordset("Lunes"))
         End If

       End If
       
      
      
       
       If Not IsNull(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas")) Then
         If Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > 0 And Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") <= HorasDias Then
'            sngHorasLaboradas = Format(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas"), "##,##0.00")
'            dLunes = Format(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas"), "##,##0.00")
            sngHorasLaboradas = Format(Me.adoHorasExtras.Recordset.Fields("HLaboradas"), "##,##0.00")
            dLunes = Format(Me.adoHorasExtras.Recordset.Fields("HLaboradas"), "##,##0.00")
            
         ElseIf Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > HorasDias Then   '9.75
            sngHorasLaboradas = HorasDias
            dLunes = HorasDias
         Else
            sngHorasLaboradas = 0
            dLunes = 0
         End If
    
      Else
         sngHorasLaboradas = 0
         dLunes = 0
      End If
       
      Me.adoConsulta.CommandType = adCmdText
      Me.adoConsulta.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion " & _
                                 "WHERE  DetalleHorasProduccion.CodEmpleado ='" & sCodEmpleado & "' AND (DetalleHorasProduccion.NumNomina = '" & Trim(Me.lblNoNomina.Caption) & "') AND (NumLinea = " & sngNumLinea & ")"
      Me.adoConsulta.Refresh
            
     If Not Me.adoConsulta.Recordset.EOF Then
       If sngHorasLaboradas <> 0 Then
         Me.adoConsulta.Recordset.Fields("Lunes") = Format(sngHorasLaboradas, "##,##0.00")
       Else
         Me.adoConsulta.Recordset.Fields("Lunes") = 0
       End If
    
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
     Else
       Me.adoConsulta.Recordset.AddNew
       Me.adoConsulta.Recordset.Fields("NumLinea") = sngNumLinea
       Me.adoConsulta.Recordset.Fields("CodEmpleado") = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
       Me.adoConsulta.Recordset.Fields("NumNomina") = Trim(Me.lblNoNomina.Caption)
       If sngHorasLaboradas <> 0 Then
          Me.adoConsulta.Recordset.Fields("Lunes") = Format(sngHorasLaboradas, "##,##0.00")
       Else
          Me.adoConsulta.Recordset.Fields("Lunes") = 0
       End If
     
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
            
            
     End If
       
       
    Case "Mart":
    
       Me.AdoHorario.Refresh
       If Not Me.AdoHorario.Recordset.EOF Then
         If DateDiff("d", sFecha1, sFecha2) <= 7 Then
           HorasDias = Me.AdoHorario.Recordset("Martes")
         ElseIf DateDiff("d", sFecha1, sFecha2) <= 14 Then
            HorasDias = CDbl(Me.AdoHorario.Recordset("Martes")) * 2
         End If
        
       End If
              
       If Not IsNull(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas")) Then
         If Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > 0 And Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") <= HorasDias Then
'            sngHorasLaboradas = Format(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas"), "##,##0.00")
'            dMartes = Format(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas"), "##,##0.00")
            sngHorasLaboradas = Format(Me.adoHorasExtras.Recordset.Fields("HLaboradas"), "##,##0.00")
            dMartes = Format(Me.adoHorasExtras.Recordset.Fields("HLaboradas"), "##,##0.00")
            
         ElseIf Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > HorasDias Then
            sngHorasLaboradas = HorasDias
            dMartes = HorasDias
         Else
            sngHorasLaboradas = 0
            dMartes = 0
         End If
    
      Else
         sngHorasLaboradas = 0
         dMartes = 0
      End If
       
      Me.adoConsulta.CommandType = adCmdText
      Me.adoConsulta.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion " & _
                                 "WHERE  DetalleHorasProduccion.CodEmpleado ='" & sCodEmpleado & "' AND DetalleHorasProduccion.NumNomina = '" & Me.lblNoNomina.Caption & "' AND (NumLinea = " & sngNumLinea & ")"
      Me.adoConsulta.Refresh
            
     If Not Me.adoConsulta.Recordset.EOF Then
       If sngHorasLaboradas <> 0 Then
         Me.adoConsulta.Recordset.Fields("Martes") = Format(sngHorasLaboradas, "##,##0.00")
       Else
         Me.adoConsulta.Recordset.Fields("Martes") = 0
       End If
    
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
     Else
       Me.adoConsulta.Recordset.AddNew
       Me.adoConsulta.Recordset.Fields("NumLinea") = sngNumLinea
       Me.adoConsulta.Recordset.Fields("CodEmpleado") = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
       Me.adoConsulta.Recordset.Fields("NumNomina") = Me.lblNoNomina.Caption
       If sngHorasLaboradas <> 0 Then
          Me.adoConsulta.Recordset.Fields("Martes") = Format(sngHorasLaboradas, "##,##0.00")
          Me.adoConsulta.Recordset.Fields("Lunes") = Format(dLunes, "##,##0.00")
       Else
          Me.adoConsulta.Recordset.Fields("Martes") = 0
          Me.adoConsulta.Recordset.Fields("Lunes") = Format(dLunes, "##,##0.00")
       End If
     
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
            
            
     End If
       
       
       
     Case "Mierc":
        
     
     
       Me.AdoHorario.Refresh
       If Not Me.AdoHorario.Recordset.EOF Then
         If DateDiff("d", sFecha1, sFecha2) <= 7 Then
           HorasDias = Me.AdoHorario.Recordset("Miercoles")
         ElseIf DateDiff("d", sFecha1, sFecha2) <= 14 Then
            HorasDias = CDbl(Me.AdoHorario.Recordset("Miercoles")) * 2
         End If
        
       End If
       
       If Not IsNull(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas")) Then
         If Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > 0 And Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") <= HorasDias Then
            sngHorasLaboradas = Format(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas"), "##,##0.00")
            dMiercoles = Format(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas"), "##,##0.00")
'''            sngHorasLaboradas = Format(Me.adoHorasExtras.Recordset.Fields("HLaboradas"), "##,##0.00")
'''            dMiercoles = Format(Me.adoHorasExtras.Recordset.Fields("HLaboradas"), "##,##0.00")
            
         ElseIf Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > HorasDias Then
            sngHorasLaboradas = HorasDias
            dMiercoles = HorasDias
         Else
            sngHorasLaboradas = 0
            dMiercoles = 0
         End If
    
      Else
         sngHorasLaboradas = 0
         dMiercoles = 0
      End If
       
      Me.adoConsulta.CommandType = adCmdText
      Me.adoConsulta.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion " & _
                                 "WHERE  DetalleHorasProduccion.CodEmpleado ='" & sCodEmpleado & "' AND DetalleHorasProduccion.NumNomina = '" & Me.lblNoNomina.Caption & "' AND (NumLinea = " & sngNumLinea & ")"
      Me.adoConsulta.Refresh
            
     If Not Me.adoConsulta.Recordset.EOF Then
       If sngHorasLaboradas <> 0 Then
         Me.adoConsulta.Recordset.Fields("Miercoles") = Format(sngHorasLaboradas, "##,##0.00")
         Me.adoConsulta.Recordset.Fields("Martes") = Format(dMartes, "##,##0.00")
         Me.adoConsulta.Recordset.Fields("Lunes") = Format(dLunes, "##,##0.00")
       Else
         Me.adoConsulta.Recordset.Fields("Miercoles") = 0
         Me.adoConsulta.Recordset.Fields("Martes") = Format(dMartes, "##,##0.00")
         Me.adoConsulta.Recordset.Fields("Lunes") = Format(dLunes, "##,##0.00")
         
       End If
    
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
     Else
       Me.adoConsulta.Recordset.AddNew
       Me.adoConsulta.Recordset.Fields("NumLinea") = sngNumLinea
       Me.adoConsulta.Recordset.Fields("CodEmpleado") = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
       Me.adoConsulta.Recordset.Fields("NumNomina") = Me.lblNoNomina.Caption
       If sngHorasLaboradas <> 0 Then
          Me.adoConsulta.Recordset.Fields("Miercoles") = Format(sngHorasLaboradas, "##,##0.00")
       Else
          Me.adoConsulta.Recordset.Fields("Miercoles") = 0
       End If
     
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
            
            
     End If
       
     Case "Juev":
     
       Me.AdoHorario.Refresh
       If Not Me.AdoHorario.Recordset.EOF Then
         If DateDiff("d", sFecha1, sFecha2) <= 7 Then
           HorasDias = Me.AdoHorario.Recordset("Jueves")
         ElseIf DateDiff("d", sFecha1, sFecha2) <= 14 Then
            HorasDias = CDbl(Me.AdoHorario.Recordset("Jueves")) * 2
         End If
        
       End If
       
       
       If Not IsNull(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas")) Then
         If Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > 0 And Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") <= HorasDias Then
            sngHorasLaboradas = Format(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas"), "##,##0.00")
            dJueves = Format(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas"), "##,##0.00")
            
         ElseIf Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > HorasDias Then
            sngHorasLaboradas = HorasDias
            dJueves = HorasDias
            
         Else
            sngHorasLaboradas = 0
            dJueves = 0
         End If
    
      Else
         sngHorasLaboradas = 0
         dJueves = 0
      End If
       
       
      Me.adoConsulta.CommandType = adCmdText
      Me.adoConsulta.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion " & _
                                 "WHERE  DetalleHorasProduccion.CodEmpleado ='" & sCodEmpleado & "' AND DetalleHorasProduccion.NumNomina = '" & Me.lblNoNomina.Caption & "' AND (NumLinea = " & sngNumLinea & ")"
      Me.adoConsulta.Refresh
            
     If Not Me.adoConsulta.Recordset.EOF Then
       If sngHorasLaboradas <> 0 Then
         Me.adoConsulta.Recordset.Fields("Jueves") = Format(sngHorasLaboradas, "##,##0.00")
         Me.adoConsulta.Recordset.Fields("Martes") = Format(dMartes, "##,##0.00")
         Me.adoConsulta.Recordset.Fields("Lunes") = Format(dLunes, "##,##0.00")
         Me.adoConsulta.Recordset.Fields("Miercoles") = Format(dMiercoles, "##,##0.00")
       Else
         Me.adoConsulta.Recordset.Fields("Jueves") = 0
         Me.adoConsulta.Recordset.Fields("Martes") = Format(dMartes, "##,##0.00")
         Me.adoConsulta.Recordset.Fields("Lunes") = Format(dLunes, "##,##0.00")
         Me.adoConsulta.Recordset.Fields("Miercoles") = Format(dMiercoles, "##,##0.00")
         
       End If
    
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
     Else
       Me.adoConsulta.Recordset.AddNew
       Me.adoConsulta.Recordset.Fields("NumLinea") = sngNumLinea
       Me.adoConsulta.Recordset.Fields("CodEmpleado") = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
       Me.adoConsulta.Recordset.Fields("NumNomina") = Me.lblNoNomina.Caption
       If sngTotalHExtras <> 0 Then
         If sngHorasLaboradas = 0 Then
          Me.adoConsulta.Recordset.Fields("Jueves") = 0
         Else
          Me.adoConsulta.Recordset.Fields("Jueves") = Format(sngHorasLaboradas, "##,##0.00")
         End If
       Else
          Me.adoConsulta.Recordset.Fields("Jueves") = 0
       End If
     
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
            
            
     End If
       
       
     Case "Viern":
     
       Me.AdoHorario.Refresh
       If Not Me.AdoHorario.Recordset.EOF Then
         If DateDiff("d", sFecha1, sFecha2) <= 7 Then
           HorasDias = Me.AdoHorario.Recordset("Viernes")
         ElseIf DateDiff("d", sFecha1, sFecha2) <= 14 Then
            HorasDias = CDbl(Me.AdoHorario.Recordset("Viernes")) * 2
         End If
       End If
       
       If Not IsNull(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas")) Then
         If Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > 0 And Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") <= HorasDias Then
            sngHorasLaboradas = Format(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas"), "##,##0.00")
            dViernes = Format(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas"), "##,##0.00")
            
         ElseIf Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > HorasDias Then    '9
            sngHorasLaboradas = HorasDias
            dViernes = HorasDias
            
         Else
            sngHorasLaboradas = 0
            dViernes = 0
         End If
    
      Else
         sngHorasLaboradas = 0
         dViernes = 0
      End If
       
      Me.adoConsulta.CommandType = adCmdText
      Me.adoConsulta.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion " & _
                                 "WHERE  DetalleHorasProduccion.CodEmpleado ='" & sCodEmpleado & "' AND DetalleHorasProduccion.NumNomina = '" & Me.lblNoNomina.Caption & "' AND (NumLinea = " & sngNumLinea & ")"
      Me.adoConsulta.Refresh
            
     If Not Me.adoConsulta.Recordset.EOF Then
       If sngHorasLaboradas <> 0 Then
         Me.adoConsulta.Recordset.Fields("Viernes") = Format(sngHorasLaboradas, "##,##0.00")
         Me.adoConsulta.Recordset.Fields("Martes") = Format(dMartes, "##,##0.00")
         Me.adoConsulta.Recordset.Fields("Lunes") = Format(dLunes, "##,##0.00")
         Me.adoConsulta.Recordset.Fields("Miercoles") = Format(dMiercoles, "##,##0.00")
         Me.adoConsulta.Recordset.Fields("Jueves") = Format(dJueves, "##,##0.00")
         
       Else
         Me.adoConsulta.Recordset.Fields("Viernes") = 0
         Me.adoConsulta.Recordset.Fields("Martes") = Format(dMartes, "##,##0.00")
         Me.adoConsulta.Recordset.Fields("Lunes") = Format(dLunes, "##,##0.00")
         Me.adoConsulta.Recordset.Fields("Miercoles") = Format(dMiercoles, "##,##0.00")
         Me.adoConsulta.Recordset.Fields("Jueves") = Format(dJueves, "##,##0.00")
         
       End If
    
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
     Else
       Me.adoConsulta.Recordset.AddNew
       Me.adoConsulta.Recordset.Fields("NumLinea") = sngNumLinea
       Me.adoConsulta.Recordset.Fields("CodEmpleado") = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
       Me.adoConsulta.Recordset.Fields("NumNomina") = Me.lblNoNomina.Caption
       If sngTotalHExtras <> 0 Then
         If sngHorasLaboradas = 0 Then
          Me.adoConsulta.Recordset.Fields("Viernes") = 0
         Else
          Me.adoConsulta.Recordset.Fields("Viernes") = Format(sngHorasLaboradas, "##,##0.00")
         End If
         
       Else
          Me.adoConsulta.Recordset.Fields("Viernes") = 0
       End If
     
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
            
            
     End If
       
     End Select
         
   If Me.adoHorasExtras.Recordset.Fields("Dia") <> "Sab" And Me.adoHorasExtras.Recordset.Fields("Dia") <> "Dom" Then
         
     sCodEmpleado = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
     
      sngTarifaHoraria = BuscaTarifaHoraria(sCodEmpleado)
      
'     sngTarifaHoraria = Me.adoAsistencia.Recordset.Fields("TarifaHoraria")
     sngTotalHLaboradas = sngTotalHLaboradas + Format(sngHorasLaboradas, "##,##0.00")
              
     Me.adoConsulta.CommandType = adCmdText
     Me.adoConsulta.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion " & _
                                 "WHERE  DetalleHorasProduccion.CodEmpleado ='" & sCodEmpleado & "' AND DetalleHorasProduccion.NumNomina = '" & Trim(Me.lblNoNomina.Caption) & "' AND (NumLinea = " & sngNumLinea & ")"
      Me.adoConsulta.Refresh
            
     If Not Me.adoConsulta.Recordset.EOF Then
       If sngTotalHLaboradas <> 0 And Me.adoConsulta.Recordset.Fields("Lunes") + Me.adoConsulta.Recordset.Fields("Martes") + Me.adoConsulta.Recordset.Fields("Miercoles") + Me.adoConsulta.Recordset.Fields("Jueves") + Me.adoConsulta.Recordset.Fields("Viernes") <> 0 Then
         Me.adoConsulta.Recordset.Fields("TotalHoras") = Format(Me.adoConsulta.Recordset.Fields("Lunes") + Me.adoConsulta.Recordset.Fields("Martes") + Me.adoConsulta.Recordset.Fields("Miercoles") + Me.adoConsulta.Recordset.Fields("Jueves") + Me.adoConsulta.Recordset.Fields("Viernes"), "##,##0.00")
       Else
         Me.adoConsulta.Recordset.Fields("TotalHoras") = 0
       End If
       
       Me.adoConsulta.Recordset.Fields("SalarioHora") = sngTarifaHoraria
       
         'If Me.adoAsistencia.Recordset.Fields("TarifaHoraria") = 0 Then
         If sngTarifaHoraria = 0 Then
            Me.adoConsulta.Recordset.Fields("TotalSalarioHora") = 0
         Else
           'If Me.adoAsistencia.Recordset.Fields("TarifaHoraria") * sngTotalHLaboradas = 0 Then
           If sngTarifaHoraria * sngTotalHLaboradas = 0 Then
            Me.adoConsulta.Recordset.Fields("TotalSalarioHora") = 0
           ElseIf Me.adoConsulta.Recordset.Fields("Lunes") + Me.adoConsulta.Recordset.Fields("Martes") + Me.adoConsulta.Recordset.Fields("Miercoles") + Me.adoConsulta.Recordset.Fields("Jueves") + Me.adoConsulta.Recordset.Fields("Viernes") <> 0 Then
            sngTotalHLaboradas = CDbl(dLunes) + CDbl(dMartes) + CDbl(dMiercoles) + CDbl(dJueves) + CDbl(dViernes) 'Me.adoConsulta.Recordset.Fields("Lunes") + Me.adoConsulta.Recordset.Fields("Martes") + Me.adoConsulta.Recordset.Fields("Miercoles") + Me.adoConsulta.Recordset.Fields("Jueves") + Me.adoConsulta.Recordset.Fields("Viernes")
            Me.adoConsulta.Recordset.Fields("TotalSalarioHora") = Format(sngTarifaHoraria * sngTotalHLaboradas, "##,###.##")
           End If
           
         End If
         
       sngPagoTotal = Me.adoConsulta.Recordset.Fields("TotalSalarioHora")
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
       
      End If
              
    End If
              
'    Me.adoAsistencia.Recordset.MoveNext
    DoEvents
    
    End If
     
     
     sngHorasLaboradas = 0
     sngTarifaHoraria = 0
      If Not Me.adoHorasExtras.Recordset.EOF Then
        Me.adoHorasExtras.Recordset.MoveNext
      End If
     Me.osProgress1.Value = Me.osProgress1.Value + 1
     bConta = bConta + 1
     DoEvents
     
    Loop
   Me.adoAsistencia.Recordset.MoveNext
   DoEvents
Loop

End If

Me.ospHoras.Visible = False
ospHoras.Value = 0
Me.osProgress1.Visible = False

End Sub

Private Sub cmdReporte_Click()

Dim sSQl As String
Dim rptAsistenciaGen As New arepAsistencia
Dim rptAsistenciaSexo As New arepAsistenciaSexo
Dim rptAsistenciaDepto As New arepAsistenciaDepto
Dim rptAsistenciaCargo As New arepAsistenciaCargo
Dim rptLaboradasExtras As New arepHLaboradas
Dim rptAsistenciaES As New arepAsistenciaReal
Dim rptAusencia As New ArepAsistenciaAusencia
Dim rptAsistenciaDepto2 As New arepAsistenciaDepto2


Dim rpt As Object, TotalCostoFijo As Double
Dim fPreview As New FrmPreview
Dim rs As New ADODB.Recordset

Dim lFecha1 As Long
Dim lFecha2 As Long
Dim dFecha1 As Date
Dim dFecha2 As Date
Dim sFecha1 As String
Dim sFecha2 As String


dFecha1 = Me.dtpDesde.Value
dFecha2 = Me.dtpHasta.Value

sFecha1 = Mid$(Me.dtpDesde.Value, 7, 4) & "-" & Mid$(Me.dtpDesde.Value, 4, 2) & "-" & Mid$(Me.dtpDesde.Value, 1, 2)
sFecha2 = Mid$(Me.dtpHasta.Value, 7, 4) & "-" & Mid$(Me.dtpHasta.Value, 4, 2) & "-" & Mid$(Me.dtpHasta.Value, 1, 2)

lFecha1 = dFecha1
lFecha2 = dFecha2





If Me.cboTipoNomina.Text <> "" And Me.lblNoNomina.Caption <> "" Then

If Me.optCodigo.Value Then

'sSQl = "SELECT AsistenciaEmpleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, " & _
'       "Empleado.Direccion, Empleado.Nacionalidad, Empleado.Sexo, Empleado.NumCedula, Departamento.Departamento, " & _
'       "Turno.CodTurno, Cargo.Cargo, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, " & _
'       "AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.Dia, " & _
'       "AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso, TipoNomina.Nomina, TipoNomina.Periodo, Turno.TComida " & _
'        "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN " & _
'        "Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
'        "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN TipoNomina ON dbo.AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
'        "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' ORDER BY Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC"


'sSQL = "SELECT dbo.AsistenciaEmpleado.CodEmpleado, dbo.Empleado.Nombre1, dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2, " & _
'       "dbo.Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
'       "dbo.Turno.CodTurno, dbo.Cargo.Cargo, dbo.AsistenciaEmpleado.FechaEntrada, dbo.AsistenciaEmpleado.HoraEntrada, " & _
'       "dbo.AsistenciaEmpleado.FechaSalida, dbo.AsistenciaEmpleado.HoraSalida, dbo.AsistenciaEmpleado.bActivo, dbo.AsistenciaEmpleado.HLaboradas, " & _
'       "dbo.AsistenciaEmpleado.Dia, dbo.AsistenciaEmpleado.HExtras, dbo.AsistenciaEmpleado.bPermiso, dbo.TipoNomina.Nomina, " & _
'       "dbo.TipoNomina.Periodo , dbo.Turno.TComida, dbo.HorasExtras.CantHoras, dbo.HorasExtras.NumNomina " & _
'       "FROM dbo.AsistenciaEmpleado INNER JOIN " & _
'       "dbo.Empleado ON dbo.AsistenciaEmpleado.CodEmpleado = dbo.Empleado.CodEmpleado INNER JOIN " & _
'       "dbo.Turno ON dbo.AsistenciaEmpleado.CodTurno = dbo.Turno.CodTurno INNER JOIN " & _
'       "dbo.Departamento ON dbo.Empleado.CodDepartamento = dbo.Departamento.CodDepartamento INNER JOIN " & _
'       "dbo.Cargo ON dbo.Empleado.CodCargo = dbo.Cargo.CodCargo INNER JOIN " & _
'       "dbo.TipoNomina ON dbo.AsistenciaEmpleado.CodTipoNomina = dbo.TipoNomina.CodTipoNomina INNER JOIN " & _
'       "dbo.HorasExtras ON dbo.Empleado.CodEmpleado = dbo.HorasExtras.CodEmpleado " & _
'       "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' AND dbo.HorasExtras.NumNomina =" & Me.lblNoNomina.Caption & " ORDER BY Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC"
        
If Me.txtEmpleado.Text = "%" Then
    sSQl = "SELECT dbo.AsistenciaEmpleado.CodEmpleado, dbo.AsistenciaEmpleado.CodEmpleado1, dbo.Empleado.Nombre1, dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2, " & _
                          "dbo.Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
                          "dbo.Turno.CodTurno, dbo.Cargo.Cargo, dbo.AsistenciaEmpleado.FechaEntrada, dbo.AsistenciaEmpleado.HREntrada, dbo.AsistenciaEmpleado.HoraEntrada, " & _
                          "dbo.AsistenciaEmpleado.FechaSalida, dbo.AsistenciaEmpleado.HoraSalida, dbo.AsistenciaEmpleado.bActivo, dbo.AsistenciaEmpleado.HLaboradas, " & _
                          "dbo.AsistenciaEmpleado.Dia, dbo.AsistenciaEmpleado.HExtras, dbo.AsistenciaEmpleado.bPermiso, dbo.TipoNomina.Nomina, " & _
                          "dbo.TipoNomina.Periodo , dbo.Turno.TComida, dbo.HorasExtras.CantHoras, dbo.HorasExtras.NumNomina " & _
           "FROM dbo.AsistenciaEmpleado INNER JOIN dbo.Empleado ON dbo.AsistenciaEmpleado.CodEmpleado = dbo.Empleado.CodEmpleado INNER JOIN " & _
                          "dbo.Turno ON dbo.AsistenciaEmpleado.CodTurno = dbo.Turno.CodTurno INNER JOIN " & _
                          "dbo.Departamento ON dbo.Empleado.CodDepartamento = dbo.Departamento.CodDepartamento INNER JOIN " & _
                          "dbo.Cargo ON dbo.Empleado.CodCargo = dbo.Cargo.CodCargo INNER JOIN " & _
                          "dbo.TipoNomina ON dbo.AsistenciaEmpleado.CodTipoNomina = dbo.TipoNomina.CodTipoNomina INNER JOIN " & _
                          "dbo.HorasExtras ON dbo.Empleado.CodEmpleado = dbo.HorasExtras.CodEmpleado " & _
                          "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND dbo.HorasExtras.NumNomina =" & Me.lblNoNomina.Caption & _
                          " GROUP BY dbo.AsistenciaEmpleado.HREntrada, dbo.AsistenciaEmpleado.CodEmpleado1, dbo.AsistenciaEmpleado.CodEmpleado, dbo.Empleado.Nombre1, dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2, " & _
                          "dbo.Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
                          "dbo.Turno.CodTurno, dbo.Cargo.Cargo, dbo.AsistenciaEmpleado.FechaEntrada, dbo.AsistenciaEmpleado.HoraEntrada, " & _
                          "dbo.AsistenciaEmpleado.FechaSalida, dbo.AsistenciaEmpleado.HoraSalida, dbo.AsistenciaEmpleado.bActivo, dbo.AsistenciaEmpleado.HLaboradas, " & _
                          "dbo.AsistenciaEmpleado.Dia, dbo.AsistenciaEmpleado.HExtras, dbo.AsistenciaEmpleado.bPermiso, dbo.TipoNomina.Nomina, " & _
                          "dbo.TipoNomina.Periodo , dbo.Turno.TComida, dbo.HorasExtras.CantHoras, dbo.HorasExtras.NumNomina"

Else
      sSQl = "SELECT dbo.AsistenciaEmpleado.CodEmpleado, dbo.AsistenciaEmpleado.CodEmpleado1, dbo.Empleado.Nombre1, dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2, " & _
                          "dbo.Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
                          "dbo.Turno.CodTurno, dbo.Cargo.Cargo, dbo.AsistenciaEmpleado.FechaEntrada, dbo.AsistenciaEmpleado.HREntrada, dbo.AsistenciaEmpleado.HoraEntrada, " & _
                          "dbo.AsistenciaEmpleado.FechaSalida, dbo.AsistenciaEmpleado.HoraSalida, dbo.AsistenciaEmpleado.bActivo, dbo.AsistenciaEmpleado.HLaboradas, " & _
                          "dbo.AsistenciaEmpleado.Dia, dbo.AsistenciaEmpleado.HExtras, dbo.AsistenciaEmpleado.bPermiso, dbo.TipoNomina.Nomina, " & _
                          "dbo.TipoNomina.Periodo , dbo.Turno.TComida, dbo.HorasExtras.CantHoras, dbo.HorasExtras.NumNomina " & _
           "FROM dbo.AsistenciaEmpleado INNER JOIN dbo.Empleado ON dbo.AsistenciaEmpleado.CodEmpleado = dbo.Empleado.CodEmpleado INNER JOIN " & _
                          "dbo.Turno ON dbo.AsistenciaEmpleado.CodTurno = dbo.Turno.CodTurno INNER JOIN " & _
                          "dbo.Departamento ON dbo.Empleado.CodDepartamento = dbo.Departamento.CodDepartamento INNER JOIN " & _
                          "dbo.Cargo ON dbo.Empleado.CodCargo = dbo.Cargo.CodCargo INNER JOIN " & _
                          "dbo.TipoNomina ON dbo.AsistenciaEmpleado.CodTipoNomina = dbo.TipoNomina.CodTipoNomina INNER JOIN " & _
                          "dbo.HorasExtras ON dbo.Empleado.CodEmpleado = dbo.HorasExtras.CodEmpleado " & _
                          "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND (HorasExtras.NumNomina =" & Me.lblNoNomina.Caption & " ) AND (AsistenciaEmpleado.CodEmpleado1 = '" & Me.txtEmpleado.Text & "') " & _
                          " GROUP BY dbo.AsistenciaEmpleado.HREntrada, dbo.AsistenciaEmpleado.CodEmpleado1, dbo.AsistenciaEmpleado.CodEmpleado, dbo.Empleado.Nombre1, dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2, " & _
                          "dbo.Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
                          "dbo.Turno.CodTurno, dbo.Cargo.Cargo, dbo.AsistenciaEmpleado.FechaEntrada, dbo.AsistenciaEmpleado.HoraEntrada, " & _
                          "dbo.AsistenciaEmpleado.FechaSalida, dbo.AsistenciaEmpleado.HoraSalida, dbo.AsistenciaEmpleado.bActivo, dbo.AsistenciaEmpleado.HLaboradas, " & _
                          "dbo.AsistenciaEmpleado.Dia, dbo.AsistenciaEmpleado.HExtras, dbo.AsistenciaEmpleado.bPermiso, dbo.TipoNomina.Nomina, " & _
                          "dbo.TipoNomina.Periodo , dbo.Turno.TComida, dbo.HorasExtras.CantHoras, dbo.HorasExtras.NumNomina"

End If

txtSQL.Text = sSQl

'rptAsistenciaGen.DataControl1.ConnectionString = ConexionRep
'rptAsistenciaGen.DataControl1.Source = sSQl
'rptAsistenciaGen.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value
''rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
'rptAsistenciaGen.Show 1

     Set rpt = New arepAsistencia
     rpt.DataControl1.ConnectionString = ConexionRep
     rpt.DataControl1.Source = sSQl
     fPreview.RunReport rpt
     fPreview.Show 1


'AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "'

ElseIf Me.optLaboradas.Value Then

   
   
   sSQl = "SELECT TOP 100 PERCENT dbo.Empleado.CodEmpleado, dbo.Empleado.Nombre1, dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2, " & _
          "dbo.Empleado.TarifaHoraria, dbo.TipoNomina.Nomina, dbo.DetalleHorasProduccion.NumNomina, dbo.DetalleHorasProduccion.Lunes, " & _
          "dbo.DetalleHorasProduccion.Martes, dbo.DetalleHorasProduccion.Miercoles, dbo.DetalleHorasProduccion.Jueves, " & _
          "dbo.DetalleHorasProduccion.Viernes, dbo.DetalleHorasProduccion.Sabado, dbo.DetalleHorasProduccion.Domingo, " & _
          "dbo.DetalleHorasProduccion.TotalHoras, dbo.DetalleHorasProduccion.SalarioHora, dbo.DetalleHorasProduccion.TotalSalarioHora, " & _
          "dbo.HorasExtras.CantHoras , dbo.Departamento.Departamento, dbo.HorasExtras.NumNomina " & _
          "FROM dbo.Departamento INNER JOIN " & _
          "dbo.Empleado ON dbo.Departamento.CodDepartamento = dbo.Empleado.CodDepartamento INNER JOIN dbo.DetalleHorasProduccion ON dbo.Empleado.CodEmpleado = dbo.DetalleHorasProduccion.CodEmpleado INNER JOIN " & _
          "dbo.HorasExtras ON dbo.Empleado.CodEmpleado = dbo.HorasExtras.CodEmpleado INNER JOIN " & _
          "dbo.TipoNomina ON dbo.Empleado.CodTipoNomina = dbo.TipoNomina.CodTipoNomina " & _
          "WHERE TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' AND DetalleHorasProduccion.NumNomina =" & Me.lblNoNomina.Caption & " AND HorasExtras.NumNomina =" & Me.lblNoNomina.Caption & " ORDER BY Empleado.CodEmpleado ASC"
          
      rptLaboradasExtras.DataControl1.ConnectionString = ConexionRep
      rptLaboradasExtras.DataControl1.Source = sSQl
      rptLaboradasExtras.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value
      'rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
      rptLaboradasExtras.Show 1

      
      

ElseIf Me.optSexo.Value Then

sSQl = "SELECT AsistenciaEmpleado.CodEmpleado,Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
       "Turno.CodTurno, Cargo.Cargo, Historico.FechaNacimiento, Historico.FechaContrato, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.HExtras, AsistenciaEmpleado.bPermiso, TipoNomina.Nomina , TipoNomina.Periodo, Turno.TComida, dbo.AsistenciaEmpleado.HREntrada " & _
       "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
       "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
       "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' ORDER BY Empleado.Sexo, Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC "

rptAsistenciaSexo.DataControl1.ConnectionString = ConexionRep
rptAsistenciaSexo.DataControl1.Source = sSQl
rptAsistenciaSexo.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value & ", Por Sexo"
'rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
rptAsistenciaSexo.Show 1

ElseIf Me.optDepto.Value And Me.cboDepto.Text <> "" Then

   If Me.cboDepto.Text = "Todos" Then
   
    sSQl = "SELECT AsistenciaEmpleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
           "Turno.CodTurno, Cargo.Cargo, Historico.FechaNacimiento, Historico.FechaContrato, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, dbo.AsistenciaEmpleado.HREntrada, AsistenciaEmpleado.HExtras, AsistenciaEmpleado.bPermiso, TipoNomina.Nomina , TipoNomina.Periodo, Turno.TComida " & _
           "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
           "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
           "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' ORDER BY Departamento.Departamento, Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC "
   Else
   
    sSQl = "SELECT AsistenciaEmpleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
           "Turno.CodTurno, Cargo.Cargo, Historico.FechaNacimiento, Historico.FechaContrato, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, dbo.AsistenciaEmpleado.HREntrada, AsistenciaEmpleado.HExtras, AsistenciaEmpleado.bPermiso, TipoNomina.Nomina , TipoNomina.Periodo, Turno.TComida " & _
           "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
           "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
           "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' AND Departamento.Departamento ='" & Me.cboDepto.Text & "' ORDER BY Departamento.Departamento, Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC "
   End If
    
    
    
    
   rptAsistenciaDepto.DataControl1.ConnectionString = ConexionRep
   rptAsistenciaDepto.DataControl1.Source = sSQl
   rptAsistenciaDepto.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value & ", Por Depto"
   'rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
   rptAsistenciaDepto.Show 1
   
ElseIf Me.OptAsistenciaDpto.Value And Me.cboDepto.Text <> "" Then

    If Me.cboDepto.Text = "Todos" Then
    
     sSQl = "SELECT AsistenciaEmpleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.Apellido1 + ' ' + Empleado.Apellido2 + ' ' + Empleado.Nombre1 + ' ' + Empleado.Nombre2  As Nombre, Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
            "Turno.CodTurno, Cargo.Cargo, Historico.FechaNacimiento, Historico.FechaContrato, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, dbo.AsistenciaEmpleado.HREntrada, AsistenciaEmpleado.HExtras, AsistenciaEmpleado.bPermiso, TipoNomina.Nomina , TipoNomina.Periodo, Turno.TComida " & _
            "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
            "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
            "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' ORDER BY Departamento.Departamento, Empleado.Apellido1, Empleado.Apellido2, AsistenciaEmpleado.FechaEntrada ASC "
    Else
    
     sSQl = "SELECT AsistenciaEmpleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.Apellido1 + ' ' + Empleado.Apellido2 + ' ' + Empleado.Nombre1 + ' ' + Empleado.Nombre2  As Nombre, Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
            "Turno.CodTurno, Cargo.Cargo, Historico.FechaNacimiento, Historico.FechaContrato, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, dbo.AsistenciaEmpleado.HREntrada, AsistenciaEmpleado.HExtras, AsistenciaEmpleado.bPermiso, TipoNomina.Nomina , TipoNomina.Periodo, Turno.TComida " & _
            "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
            "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
            "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' AND Departamento.Departamento ='" & Me.cboDepto.Text & "' ORDER BY Departamento.Departamento, Empleado.Apellido1, Empleado.Apellido2, AsistenciaEmpleado.FechaEntrada ASC "
    End If
     
     
     
     
    rptAsistenciaDepto2.DataControl1.ConnectionString = ConexionRep
    rptAsistenciaDepto2.DataControl1.Source = sSQl
    rptAsistenciaDepto2.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value & ", Por Depto"
    'rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
    rptAsistenciaDepto2.Show 1

ElseIf Me.optCargo.Value And Me.cboCargo.Text <> "" Then

  If Me.cboCargo.Text = "Todos" Then
    sSQl = "SELECT dbo.AsistenciaEmpleado.HREntrada,AsistenciaEmpleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
           "Turno.CodTurno, Cargo.Cargo, Historico.FechaNacimiento, Historico.FechaContrato, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.HExtras, AsistenciaEmpleado.bPermiso, TipoNomina.Nomina , TipoNomina.Periodo, Turno.TComida " & _
           "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
           "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
           "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' ORDER BY Cargo.Cargo, Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC "
  Else
   
    sSQl = "SELECT dbo.AsistenciaEmpleado.HREntrada, AsistenciaEmpleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
           "Turno.CodTurno, Cargo.Cargo, Historico.FechaNacimiento, Historico.FechaContrato, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.HExtras, AsistenciaEmpleado.bPermiso, TipoNomina.Nomina , TipoNomina.Periodo, Turno.TComida " & _
           "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
           "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
           "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' AND Cargo.Cargo ='" & Me.cboCargo.Text & "' ORDER BY Cargo.Cargo, Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC "
   End If
    
    
    
    
   rptAsistenciaCargo.DataControl1.ConnectionString = ConexionRep
   rptAsistenciaCargo.DataControl1.Source = sSQl
'   rptAsistenciaCargo.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value & ", Por Cargo"
   'rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
   rptAsistenciaCargo.Show 1


ElseIf Me.optSalidasNoRegistradas.Value Then



sSQl = "SELECT dbo.AsistenciaEmpleado.HREntrada, AsistenciaEmpleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, " & _
       "Empleado.Direccion, Empleado.Nacionalidad, Empleado.Sexo, Empleado.NumCedula, Departamento.Departamento, " & _
       "Turno.CodTurno, Cargo.Cargo, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, " & _
       "AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.Dia, " & _
       "AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso, TipoNomina.Nomina, TipoNomina.Periodo, Turno.TComida " & _
        "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN " & _
        "Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
        "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN TipoNomina ON dbo.AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
        "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND (AsistenciaEmpleado.FechaSalida IS NULL) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' ORDER BY Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC"

rptAsistenciaGen.LblTitulo.Caption = Me.optSalidasNoRegistradas.Caption
rptAsistenciaGen.DataControl1.ConnectionString = ConexionRep
rptAsistenciaGen.DataControl1.Source = sSQl
rptAsistenciaGen.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value
'rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
rptAsistenciaGen.Show 1

ElseIf Me.optEntradasNoRegistradas.Value Then



sSQl = "SELECT dbo.AsistenciaEmpleado.HREntrada, AsistenciaEmpleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, " & _
       "Empleado.Direccion, Empleado.Nacionalidad, Empleado.Sexo, Empleado.NumCedula, Departamento.Departamento, " & _
       "Turno.CodTurno, Cargo.Cargo, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, " & _
       "AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.Dia, " & _
       "AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso, TipoNomina.Nomina, TipoNomina.Periodo, Turno.TComida " & _
        "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN " & _
        "Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
        "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN TipoNomina ON dbo.AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
        "WHERE AsistenciaEmpleado.FechaSalida >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND (AsistenciaEmpleado.FechaEntrada IS NULL) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' ORDER BY Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC"

rptAsistenciaGen.LblTitulo.Caption = Me.optSalidasNoRegistradas.Caption
rptAsistenciaGen.DataControl1.ConnectionString = ConexionRep
rptAsistenciaGen.DataControl1.Source = sSQl
rptAsistenciaGen.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value
'rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
rptAsistenciaGen.Show 1

ElseIf Me.OptAusencia Then
 sSQl = "SELECT DISTINCT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres FROM  Empleado INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina WHERE (Empleado.CodEmpleado1 NOT IN (SELECT DISTINCT AsistenciaEmpleado.CodEmpleado1  FROM AsistenciaEmpleado INNER JOIN Empleado AS Empleado_1 ON AsistenciaEmpleado.CodEmpleado = Empleado_1.CodEmpleado " & _
        "WHERE  (AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102)) AND (AsistenciaEmpleado.FechaSalida >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102)))) AND (Empleado.Activo = 1) AND (TipoNomina.Nomina = '" & Me.cboTipoNomina.Text & "') ORDER BY Empleado.CodEmpleado1"


' sSQl = "SELECT DISTINCT CodEmpleado1, Nombre1 + ' ' + Nombre2 + ' ' + Apellido1 + ' ' + Apellido2 AS Nombres From Empleado WHERE (CodEmpleado1 NOT IN (SELECT DISTINCT AsistenciaEmpleado.CodEmpleado1  FROM AsistenciaEmpleado INNER JOIN Empleado AS Empleado_1 ON AsistenciaEmpleado.CodEmpleado = Empleado_1.CodEmpleado " & _
'        "WHERE (AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102)) AND (AsistenciaEmpleado.FechaSalida >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102)))) AND (Activo = 1) AND (CodTipoNomina = '" & Me.cboTipoNomina.Text & "') ORDER BY CodEmpleado1"
'

 
 rptAusencia.DataControl1.ConnectionString = ConexionRep
 rptAusencia.DataControl1.Source = sSQl
 rptAusencia.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value
 rptAusencia.Show 1




Else

sSQl = "SELECT dbo.AsistenciaEmpleado.CodEmpleado, dbo.AsistenciaEmpleado.CodEmpleado1, dbo.Empleado.Nombre1, dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2, " & _
                      "dbo.Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
                      "dbo.Turno.CodTurno, dbo.Cargo.Cargo, dbo.AsistenciaEmpleado.FechaEntrada, dbo.AsistenciaEmpleado.HoraEntrada, " & _
                      "dbo.AsistenciaEmpleado.FechaSalida, dbo.AsistenciaEmpleado.HoraSalida, dbo.AsistenciaEmpleado.bActivo, dbo.AsistenciaEmpleado.HLaboradas, dbo.AsistenciaEmpleado.HREntrada, dbo.AsistenciaEmpleado.HRSalida," & _
                      "dbo.AsistenciaEmpleado.Dia, dbo.AsistenciaEmpleado.HExtras, dbo.AsistenciaEmpleado.bPermiso, dbo.TipoNomina.Nomina, " & _
                      "dbo.TipoNomina.Periodo , dbo.Turno.TComida, dbo.HorasExtras.CantHoras, dbo.HorasExtras.NumNomina " & _
       "FROM dbo.AsistenciaEmpleado INNER JOIN dbo.Empleado ON dbo.AsistenciaEmpleado.CodEmpleado = dbo.Empleado.CodEmpleado INNER JOIN " & _
                      "dbo.Turno ON dbo.AsistenciaEmpleado.CodTurno = dbo.Turno.CodTurno INNER JOIN " & _
                      "dbo.Departamento ON dbo.Empleado.CodDepartamento = dbo.Departamento.CodDepartamento INNER JOIN " & _
                      "dbo.Cargo ON dbo.Empleado.CodCargo = dbo.Cargo.CodCargo INNER JOIN " & _
                      "dbo.TipoNomina ON dbo.AsistenciaEmpleado.CodTipoNomina = dbo.TipoNomina.CodTipoNomina INNER JOIN " & _
                      "dbo.HorasExtras ON dbo.Empleado.CodEmpleado = dbo.HorasExtras.CodEmpleado " & _
                      "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' AND dbo.HorasExtras.NumNomina =" & Me.lblNoNomina.Caption & _
                      " GROUP BY dbo.AsistenciaEmpleado.CodEmpleado1, dbo.AsistenciaEmpleado.CodEmpleado, dbo.Empleado.Nombre1, dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2, " & _
                      "dbo.Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
                      "dbo.Turno.CodTurno, dbo.Cargo.Cargo, dbo.AsistenciaEmpleado.FechaEntrada, dbo.AsistenciaEmpleado.HoraEntrada, " & _
                      "dbo.AsistenciaEmpleado.FechaSalida, dbo.AsistenciaEmpleado.HoraSalida, dbo.AsistenciaEmpleado.bActivo, dbo.AsistenciaEmpleado.HLaboradas, dbo.AsistenciaEmpleado.HREntrada, dbo.AsistenciaEmpleado.HRSalida, " & _
                      "dbo.AsistenciaEmpleado.Dia, dbo.AsistenciaEmpleado.HExtras, dbo.AsistenciaEmpleado.bPermiso, dbo.TipoNomina.Nomina, " & _
                      "dbo.TipoNomina.Periodo , dbo.Turno.TComida, dbo.HorasExtras.CantHoras, dbo.HorasExtras.NumNomina ORDER BY AsistenciaEmpleado.CodEmpleado1"


sSQl = "SELECT     AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.Direccion, "
                   sSQl = sSQl + "  Empleado.Nacionalidad, Empleado.Sexo, Empleado.NumCedula, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida,"
                   sSQl = sSQl + "   AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.HREntrada, AsistenciaEmpleado.HRSalida, AsistenciaEmpleado.Dia,"
                   sSQl = sSQl + "   AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso, HorasExtras.CantHoras, HorasExtras.NumNomina"
sSQl = sSQl + " FROM         AsistenciaEmpleado INNER JOIN"
sSQl = sSQl + "                      Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN"
sSQl = sSQl + "                      HorasExtras ON Empleado.CodEmpleado = HorasExtras.CodEmpleado"
sSQl = sSQl + " WHERE   AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND dbo.HorasExtras.NumNomina =" & Me.lblNoNomina.Caption & ""
sSQl = sSQl + " GROUP BY AsistenciaEmpleado.CodEmpleado1, AsistenciaEmpleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.Direccion,"
sSQl = sSQl + "                      Empleado.Nacionalidad, Empleado.Sexo, Empleado.NumCedula, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida,"
sSQl = sSQl + "                      AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.HREntrada, AsistenciaEmpleado.HRSalida, AsistenciaEmpleado.Dia,"
sSQl = sSQl + "                      AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso, HorasExtras.CantHoras, HorasExtras.NumNomina"
sSQl = sSQl + " ORDER BY AsistenciaEmpleado.CodEmpleado1"

rptAsistenciaES.DataControl1.ConnectionString = ConexionRep
rptAsistenciaES.DataControl1.Source = sSQl
rptAsistenciaES.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value
'rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
rptAsistenciaES.Show 1

End If



End If




End Sub

Private Sub CmdSalir_Click()

Unload Me

End Sub




Private Sub dtpDesde_Change()
    If Me.OptAusencia.Value = True Then
      Me.dtpHasta.Value = Me.dtpDesde.Value
    End If
End Sub

Private Sub dtpDesde_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   dtpHasta.SetFocus
End If


End Sub

Private Sub dtpHasta_Change()
     If Me.OptAusencia.Value = True Then
      Me.dtpHasta.Value = Me.dtpDesde.Value
    End If
End Sub

Private Sub dtpHasta_KeyDown(KeyCode As Integer, Shift As Integer)
Dim dFecha1 As Date
Dim dFecha2 As Date
Dim sFecha1 As String
Dim sFecha2 As String

If Me.OptAusencia.Value = True Then
  Me.dtpHasta.Value = Me.dtpDesde.Value
End If


If KeyCode = 13 Then

  dFecha1 = Me.dtpDesde.Value
  dFecha2 = Me.dtpHasta.Value

  sFecha1 = Mid$(Me.dtpDesde.Value, 7, 4) & "-" & Mid$(Me.dtpDesde.Value, 4, 2) & "-" & Mid$(Me.dtpDesde.Value, 1, 2)
  sFecha2 = Mid$(Me.dtpHasta.Value, 7, 4) & "-" & Mid$(Me.dtpHasta.Value, 4, 2) & "-" & Mid$(Me.dtpHasta.Value, 1, 2)





  Me.adoTipoNomina.CommandType = adCmdText
  Me.adoTipoNomina.RecordSource = "SELECT TipoNomina.CodTipoNomina, TipoNomina.Nomina, Nomina.NumNomina, Nomina.FechaNominaINI, Nomina.FechaNomina, " & _
                                "Nomina.Activa FROM Nomina INNER JOIN TipoNomina ON dbo.Nomina.CodTipoNomina = dbo.TipoNomina.CodTipoNomina " & _
                                "WHERE Nomina.FechaNominaINI= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND " & _
                                "Nomina.FechaNomina= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102)"
                                
                                
  Me.adoTipoNomina.Refresh


  If Not Me.adoTipoNomina.Recordset.EOF Then
    Me.dtpDesde.Value = Me.adoTipoNomina.Recordset.Fields("FechaNominaINI")
    Me.dtpHasta.Value = Me.adoTipoNomina.Recordset.Fields("FechaNomina")
    Me.lblNoNomina.Caption = Me.adoTipoNomina.Recordset.Fields("NumNomina")
    cmdReporte.Enabled = True
    cboTipoNomina.Text = Me.adoTipoNomina.Recordset.Fields("Nomina")
    cmdReporte.SetFocus
    
  End If

End If
End Sub

Private Sub Form_Activate()

Me.cboTipoNomina.SetFocus

End Sub
Public Sub CalcularLaboradas(sDia As String)
  Dim sngHorasLaboradas As Double, sngHorasExtras As Double


          sDia = "Lun"

          If Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss") <= Format(Me.adoTurno.Recordset.Fields("LEntrada"), "hh:mm:ss") Then
            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:00:00" Then
              sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
              sngHorasExtras = 0
            Else
              sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoTurno.Recordset.Fields("LSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
              sngHorasExtras = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) - sngHorasLaboradas - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            End If



         ElseIf (Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:00:00" And Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss") >= "06:30:00" And Format(Me.adoTurno.Recordset.Fields("LSalida"), "hh:mm:ss") >= Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) Then  ' Or (Me.adoAsistencia.Recordset.Fields("HoraSalida") < "23:59:59" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "17:00:00") Then
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("LEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
             sngHorasExtras = 0

          Else
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("LEntrada"), "hh:mm:ss"), Format(Me.adoTurno.Recordset.Fields("LSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Format(Me.adoTurno.Recordset.Fields("LSalida"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)

          End If

          If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= Format(Me.adoTurno.Recordset.Fields("LSalida"), "hh:mm:ss") And Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") > "12:00:00" Then
             sngHorasLaboradas = (DateDiff("n", Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = 0

          Else

            If Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss") <= "12:00:00" Then
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")) / 60)
              sngHorasExtras = 0
            Else
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Format(Me.adoTurno.Recordset.Fields("LSalida"), "hh:mm:ss")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
              ' sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - sngHorasLaboradas
            End If

          End If

          If sngHorasLaboradas <= 0 Then
            sngHorasLaboradas = 0
            sngHorasExtras = 0
          End If
          
          
          HorasLaboradas = sngHorasLaboradas
          HorasExtra = sngHorasExtras


End Sub





Private Sub Form_Load()

 Dim RutaServer As String
 Dim Server As String, Fecha As String
' Dim Conexion As String
 Dim Clave As String
 Dim User As String
 
 

 Dim ConexionSTR1 As String
 Dim TxtClaveEntrada As String
'abro el archivo para solo lectura de la cadena de conexion
 Dim NextLine As String
 Dim Autorizado As Boolean
   Autorizado = False

' Open App.Path + "\SysInfo.dll" For Input As #1
'  Do Until EOF(1)
'   Line Input #1, NextLine
'        ConexionSTR1 = Trim(NextLine)
'   Loop
' Close #1
  
Me.dtpDesde.Value = Format(Now, "dd/mm/yyyy")
Me.dtpHasta.Value = Format(Now, "dd/mm/yyyy")
Me.dtpHDesde.Value = Format(Now, "dd/mm/yyyy")
Me.dtpHHasta.Value = Format(Now, "dd/mm/yyyy")
  



ConexionRep = Conexion

Me.AdoHorario.ConnectionString = Conexion
Me.adoAsistencia.ConnectionString = Conexion
Me.adoInasistencia.ConnectionString = Conexion
Me.adoPermisos.ConnectionString = Conexion
Me.adoTipoNomina.ConnectionString = Conexion
Me.adoTurno.ConnectionString = Conexion

Me.adoConsulta.ConnectionString = Conexion
Me.adoIncentivo.ConnectionString = Conexion
Me.adoHorasExtras.ConnectionString = Conexion



Me.adoAsistencia.CommandType = adCmdText
Me.adoAsistencia.RecordSource = "SELECT * FROM AsistenciaEmpleado"
Me.adoAsistencia.Refresh

Me.adoInasistencia.CommandType = adCmdText
Me.adoInasistencia.RecordSource = "SELECT * FROM Inasistencias"
Me.adoInasistencia.Refresh

Me.adoIncentivo.CommandType = adCmdText
Me.adoIncentivo.RecordSource = "SELECT NumIncentivo, CodEmpleado, CodTipoIncentivo, NumVeces, Pagado FROM Incentivo"
Me.adoIncentivo.Refresh

Me.adoPermisos.CommandType = adCmdText
Me.adoPermisos.RecordSource = "SELECT * FROM Permisos"
Me.adoPermisos.Refresh

Me.adoTurno.CommandType = adCmdText
Me.adoTurno.RecordSource = "SELECT * FROM Turno"
Me.adoTurno.Refresh

Me.adoTipoNomina.CommandType = adCmdTable
Me.adoTipoNomina.RecordSource = "Departamento"
Me.adoTipoNomina.Refresh

Me.cboDepto.AddItem "Todos"

Do While Not Me.adoTipoNomina.Recordset.EOF
   DoEvents
   Me.cboDepto.AddItem Me.adoTipoNomina.Recordset.Fields("Departamento")
   Me.adoTipoNomina.Recordset.MoveNext
DoEvents
Loop



Me.adoTipoNomina.CommandType = adCmdTable
Me.adoTipoNomina.RecordSource = "Cargo"
Me.adoTipoNomina.Refresh

Me.cboCargo.AddItem "Todos"

Do While Not Me.adoTipoNomina.Recordset.EOF
  DoEvents
   Me.cboCargo.AddItem Me.adoTipoNomina.Recordset.Fields("Cargo")
   Me.adoTipoNomina.Recordset.MoveNext
DoEvents
Loop

Me.adoTipoNomina.CommandType = adCmdText
Me.adoTipoNomina.RecordSource = "SELECT TipoNomina.CodTipoNomina, TipoNomina.Nomina, Nomina.NumNomina, Nomina.FechaNominaINI, Nomina.FechaNomina, " & _
                                "Nomina.Activa FROM Nomina INNER JOIN TipoNomina ON dbo.Nomina.CodTipoNomina = dbo.TipoNomina.CodTipoNomina " & _
                                "WHERE (Nomina.Activa = 1)"
Me.adoTipoNomina.Refresh

Do While Not Me.adoTipoNomina.Recordset.EOF
DoEvents
   Me.cboTipoNomina.AddItem Me.adoTipoNomina.Recordset.Fields("Nomina")
   Me.adoTipoNomina.Recordset.MoveNext
DoEvents
Loop



'Fecha = "23/06/2014"
'If DateDiff("d", Now, Fecha) < 0 Then
'  Unload Me
'End If



End Sub

Private Sub OptAsistenciaDpto_Click()
  If Me.OptAsistenciaDpto.Value = True Then
    Me.cboDepto.Text = "Todos"
  Else
    Me.cboDepto.Text = ""
  End If
End Sub

Private Sub OptAusencia_Click()
If Me.OptAusencia.Value = True Then
  Me.dtpHasta.Value = Me.dtpDesde.Value
End If
End Sub

Private Sub Option1_Click()

End Sub
