VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmReloj 
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13830
   Icon            =   "frmReloj.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   13830
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc adoExtraTurno 
      Height          =   330
      Left            =   5640
      Top             =   10320
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "ExtraTurno"
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
      Left            =   1680
      Top             =   10800
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
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
      Connect         =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=PAYROLL"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=PAYROLL"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Turno"
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
   Begin MSAdodcLib.Adodc adoSabados 
      Height          =   330
      Left            =   4560
      Top             =   11040
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
      Connect         =   $"frmReloj.frx":058A
      OLEDBString     =   $"frmReloj.frx":0612
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "AsistenciaEmpleado"
      Caption         =   "Sabados y Domingos"
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
   Begin MSCommLib.MSComm mscReloj2 
      Left            =   1680
      Top             =   10320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm mscReloj1 
      Left            =   600
      Top             =   10320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Data DtaServidor 
      Caption         =   "DtaServidor"
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
      Top             =   10800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc adoHorarios 
      Height          =   375
      Left            =   12120
      Top             =   10200
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
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
      Connect         =   $"frmReloj.frx":069A
      OLEDBString     =   $"frmReloj.frx":0722
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "HorarioEmpleado"
      Caption         =   "Horarios Empleados"
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
   Begin VB.Timer tmrReloj 
      Interval        =   1000
      Left            =   10320
      Top             =   10800
   End
   Begin MSAdodcLib.Adodc adoEmpleado 
      Height          =   375
      Left            =   13200
      Top             =   10560
      Visible         =   0   'False
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
      Connect         =   $"frmReloj.frx":07AA
      OLEDBString     =   $"frmReloj.frx":0832
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "vstEmpleadoNomina"
      Caption         =   "Empleado"
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
   Begin MSAdodcLib.Adodc adoEntrada 
      Height          =   375
      Left            =   3960
      Top             =   10800
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   $"frmReloj.frx":08BA
      OLEDBString     =   $"frmReloj.frx":0942
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frmReloj.frx":09CA
      Caption         =   "Entrada del Sisterma"
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
   Begin VB.Frame fraReloj2 
      Height          =   9975
      Left            =   7800
      TabIndex        =   1
      Top             =   240
      Width           =   7455
      Begin VB.Label lblMensaje2 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1095
         Left            =   480
         TabIndex        =   21
         Top             =   8760
         Width           =   6735
      End
      Begin VB.Label lblHora2 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   20
         Top             =   8160
         Width           =   4695
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   19
         Top             =   8280
         Width           =   495
      End
      Begin VB.Label lblFecha2 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   18
         Top             =   7680
         Width           =   4935
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   17
         Top             =   7800
         Width           =   615
      End
      Begin VB.Label lblNombre2 
         AutoSize        =   -1  'True
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1440
         TabIndex        =   16
         Top             =   6960
         Width           =   165
      End
      Begin VB.Label Label6 
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   7080
         Width           =   855
      End
      Begin VB.Label lblCodigo2 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         TabIndex        =   14
         Top             =   6000
         Width           =   4935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   13
         Top             =   6120
         Width           =   705
      End
      Begin VB.Label lblReloj2 
         AutoSize        =   -1  'True
         Caption         =   "RELOJ # 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   1875
      End
      Begin VB.Image imgReloj2 
         Height          =   4860
         Left            =   480
         Picture         =   "frmReloj.frx":0A55
         Stretch         =   -1  'True
         Top             =   840
         Width           =   6360
      End
   End
   Begin VB.Frame fraReloj1 
      Height          =   9975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      Begin VB.Label lblMensaje1 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   1095
         Left            =   360
         TabIndex        =   12
         Top             =   8760
         Width           =   6975
      End
      Begin VB.Label lblHora1 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   8160
         Width           =   4695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   10
         Top             =   8280
         Width           =   495
      End
      Begin VB.Label lblFecha1 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   7680
         Width           =   6015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   8
         Top             =   7800
         Width           =   615
      End
      Begin VB.Label lblNombre1 
         AutoSize        =   -1  'True
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1200
         TabIndex        =   7
         Top             =   6840
         Width           =   165
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   7080
         Width           =   855
      End
      Begin VB.Label lblCodigo1 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1200
         TabIndex        =   5
         Top             =   5880
         Width           =   6135
      End
      Begin VB.Label lblEtiqCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   4
         Top             =   6000
         Width           =   705
      End
      Begin VB.Label lblReloj1 
         AutoSize        =   -1  'True
         Caption         =   "RELOJ # 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1875
      End
      Begin VB.Image imgReloj1 
         Height          =   4860
         Left            =   360
         Picture         =   "frmReloj.frx":4865
         Stretch         =   -1  'True
         Top             =   840
         Width           =   6480
      End
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   675
      Left            =   7680
      TabIndex        =   25
      Top             =   10320
      Width           =   210
   End
   Begin VB.Label lblHora 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4440
      TabIndex        =   24
      Top             =   10440
      Width           =   165
   End
   Begin VB.Label Label5 
      Caption         =   "HORA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   23
      Top             =   10560
      Width           =   1095
   End
   Begin VB.Label lblOtroMensaje 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   10800
      Visible         =   0   'False
      Width           =   4215
   End
End
Attribute VB_Name = "frmReloj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
 Dim RutaServer As String
 Dim Server As String
' Dim Conexion As String
 
 RutaServer = App.Path + "\CntNominas.dll"

  With Me.dtaServidor
     .DatabaseName = RutaServer
     .RecordSource = "Servidor"
     .Refresh
  End With

  If Not IsNull(Me.dtaServidor.Recordset.Servidor) Then
   Server = Me.dtaServidor.Recordset.Servidor
  Else
   MsgBox "No se ha definido el Servidor", vbCritical, "Sistmea de Nominas"
   Exit Sub
  End If

'Borrar despues la siguiente linea
'Server = "MOISES\MOISES"

'Conexion = "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=" & Server

Me.adoEmpleado.ConnectionString = Conexion
Me.adoEntrada.ConnectionString = Conexion
Me.adoHorarios.ConnectionString = Conexion
Me.adoSabados.ConnectionString = Conexion
Me.adoTurno.ConnectionString = Conexion
Me.adoExtraTurno.ConnectionString = Conexion

Me.adoTurno.CommandType = adCmdTable
Me.adoTurno.RecordSource = "Turno"
Me.adoTurno.Refresh


Me.adoEmpleado.CommandType = adCmdTable
Me.adoEmpleado.RecordSource = "vstEmpleadoNomina"
Me.adoEmpleado.Refresh


Me.adoHorarios.CommandType = adCmdText
Me.adoHorarios.RecordSource = "SELECT * FROM HorarioEmpleado"
Me.adoHorarios.Refresh


Me.adoSabados.CommandType = adCmdText
Me.adoSabados.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, FechaSalida, HoraSalida, bActivo, CodTurno " & _
                             "FROM AsistenciaEmpleado WHERE FechaEntrada =" & CDate(Mid$(Now, 1, 10))
Me.adoSabados.Refresh


Me.adoEntrada.CommandType = adCmdText
Me.adoEntrada.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, HoraSalida, FechaSalida, bActivo, CodTurno " & _
                              "FROM AsistenciaEmpleado WHERE bActivo=1"
Me.adoEntrada.Refresh

Me.adoExtraTurno.CommandType = adCmdText
Me.adoExtraTurno.RecordSource = "SELECT CodEmpleado, CodDepartamento, CodTipoNomina, FechaEntrada, HoraEntrada, FechaSalida, HoraSalida, bActivo, TiempoReceso, HorasLaboradas FROM ExtraTurno " & _
                               "WHERE (bActivo = 1)"
Me.adoExtraTurno.Refresh



Me.mscReloj1.CommPort = 1
Me.mscReloj1.Settings = "9600,N,8,1"
Me.mscReloj1.InputMode = comInputModeText
Me.mscReloj1.Handshaking = comRTSXOnXOff
Me.mscReloj1.SThreshold = 1
Me.mscReloj1.RThreshold = 1

Me.mscReloj1.PortOpen = True

Me.mscReloj2.CommPort = 2
Me.mscReloj2.Settings = "9600,N,8,1"
Me.mscReloj2.InputMode = comInputModeText
Me.mscReloj2.Handshaking = comRTSXOnXOff
Me.mscReloj2.SThreshold = 1
Me.mscReloj2.RThreshold = 1


'Me.mscReloj2.PortOpen = True




End Sub






Private Sub mscReloj1_OnComm()

Dim Fecha As Long, Fecha1 As Date
Dim NumEmpleado As Variant
Dim Longitud As Byte
Dim dHoraEntrada As Variant
Dim dHoraSalida As Variant
Dim dHEntradaHoy As Date
Dim dHSalidaHoy As Date
Dim bFinSemana As Boolean
Dim sDia As String
Dim sFecha1 As String
Dim dHMaximaEntrada As Date
Dim sHMaximaSalida As Date

 On Error GoTo TratarError

 Me.lblMensaje1.ForeColor = &HC00000
 Fecha1 = Format(Now, "dd/mm/yyyy")
 
Select Case Me.mscReloj1.CommEvent

   Case CommBreak


   Case 2

 NumEmpleado = Me.mscReloj1.Input
 Longitud = Len(NumEmpleado)
       'Me.List1.AddItem (Mid(NumEmpleado, 1, Longitud - 2)) & " Reloj1"
'       Me.mscReloj.Output = "ATDT" & vbCr
 NumEmpleado = Mid(NumEmpleado, 1, Longitud - 2)
       
 bUbicacion = InStr(1, Format(Now, "Long Date"), ",")
 sDia = Mid$(Format(Now, "Long Date"), 1, bUbicacion - 1)
 
 Me.adoHorarios.Refresh
 Me.adoHorarios.Recordset.Find "CodEmpleado =" & NumEmpleado
 
 If Me.adoHorarios.Recordset.EOF Then
    
   lblCodigo1.Caption = NumEmpleado
   Me.lblNombre1.Caption = "No Encontrado"
   Me.lblFecha1.Caption = Format(Now, "Long Date")
   Me.lblHora1.Caption = Time()
   Me.lblMensaje1.ForeColor = &H40&
   Me.lblMensaje1.Caption = "EMPLEADO NO ENCONTRADO"
   Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\NoDisponible.jpg")
 
   Exit Sub
    
 End If
 
 
  
 
 Select Case sDia
 
 Case "Lunes":
 
      If Me.adoHorarios.Recordset.Fields("MEntrada") > Time Then
         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("MEntrada")
      Else
         dHEntradaHoy = Time
      End If
      
      sTurno = Me.adoHorarios.Recordset.Fields("TurnoLunes")
      
      Me.adoTurno.CommandType = adCmdText
      Me.adoTurno.RecordSource = "SELECT * FROM Turno WHERE CodTurno ='" & sTurno & "'"
      Me.adoTurno.Refresh
      
      dHoraEntrada = TimeSerial(CInt(Mid$(Me.adoHorarios.Recordset.Fields("MEntrada"), 1, 2)) - 1, Mid$(Me.adoHorarios.Recordset.Fields("MEntrada"), 4, 2), 0)
      dHoraSalida = Me.adoHorarios.Recordset.Fields("LSalida")
 
 Case "Martes":
 
      If Me.adoHorarios.Recordset.Fields("MEntrada") > Time Then
         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("MEntrada")
      Else
         dHEntradaHoy = Time
      End If
      
      sTurno = Me.adoHorarios.Recordset.Fields("TurnoMartes")
      Me.adoTurno.CommandType = adCmdText
      Me.adoTurno.RecordSource = "SELECT * FROM Turno WHERE CodTurno ='" & sTurno & "'"
      Me.adoTurno.Refresh
      
      dHoraEntrada = TimeSerial(CInt(Mid$(Me.adoHorarios.Recordset.Fields("MEntrada"), 1, 2)) - 1, Mid$(Me.adoHorarios.Recordset.Fields("MEntrada"), 4, 2), 0)
      dHoraSalida = Me.adoHorarios.Recordset.Fields("MSalida")
      
 Case "Miércoles":
 
      If Me.adoHorarios.Recordset.Fields("MCEntrada") > Time Then
         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("MCEntrada")
      Else
         dHEntradaHoy = Time
      End If
      
      sTurno = Me.adoHorarios.Recordset.Fields("TurnoMiercoles")
      
      Me.adoTurno.CommandType = adCmdText
      Me.adoTurno.RecordSource = "SELECT * FROM Turno WHERE CodTurno ='" & sTurno & "'"
      Me.adoTurno.Refresh
      
      dHoraEntrada = TimeSerial(CInt(Mid$(Me.adoHorarios.Recordset.Fields("MCEntrada"), 1, 2)) - 1, Mid$(Me.adoHorarios.Recordset.Fields("MCEntrada"), 4, 2), 0)
      dHoraSalida = Me.adoHorarios.Recordset.Fields("MCSalida")
      
 Case "Jueves":
 
      If Me.adoHorarios.Recordset.Fields("MEntrada") > Time Then
         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("JEntrada")
      Else
         dHEntradaHoy = Time
      End If
      
      
      sTurno = Me.adoHorarios.Recordset.Fields("TurnoJueves")
      
      Me.adoTurno.CommandType = adCmdText
      Me.adoTurno.RecordSource = "SELECT * FROM Turno WHERE CodTurno ='" & sTurno & "'"
      Me.adoTurno.Refresh
      
      dHoraEntrada = TimeSerial(CInt(Mid$(Me.adoHorarios.Recordset.Fields("JEntrada"), 1, 2)) - 1, Mid$(Me.adoHorarios.Recordset.Fields("JEntrada"), 4, 2), 0)
      dHoraSalida = Me.adoHorarios.Recordset.Fields("JSalida")
      
 Case "Viernes":
 
      If Me.adoHorarios.Recordset.Fields("VEntrada") > Time Then
         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("MEntrada")
      Else
         dHEntradaHoy = Time
      End If
      
      
      sTurno = Me.adoHorarios.Recordset.Fields("TurnoViernes")
      Me.adoTurno.CommandType = adCmdText
      Me.adoTurno.RecordSource = "SELECT * FROM Turno WHERE CodTurno ='" & sTurno & "'"
      Me.adoTurno.Refresh
      
      dHoraEntrada = TimeSerial(CInt(Mid$(Me.adoHorarios.Recordset.Fields("VEntrada"), 1, 2)) - 1, Mid$(Me.adoHorarios.Recordset.Fields("VEntrada"), 4, 2), 0)
      dHoraSalida = Me.adoHorarios.Recordset.Fields("VSalida")
      
 Case "Sábado":
        
       If Me.adoHorarios.Recordset.Fields("DEntrada") > Time Then
         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("DEntrada")
      Else
         dHEntradaHoy = Time
      End If
      
       sTurno = Me.adoHorarios.Recordset.Fields("TurnoSabado")
       
       Me.adoTurno.CommandType = adCmdText
       Me.adoTurno.RecordSource = "SELECT * FROM Turno WHERE CodTurno ='" & sTurno & "'"
       Me.adoTurno.Refresh
       
       bFinSemana = True
       dHoraEntrada = TimeSerial(CInt(Mid$(Time, 1, 2)) - 1, Mid$(Time, 4, 2), 0)
       dHoraSalida = Time
       
 Case "Domingo":
       
      If Me.adoHorarios.Recordset.Fields("DEntrada") > Time Then
         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("DEntrada")
      Else
         dHEntradaHoy = Time
      End If
       
       sTurno = Me.adoHorarios.Recordset.Fields("TurnoDomingo")
       
       Me.adoTurno.CommandType = adCmdText
       Me.adoTurno.RecordSource = "SELECT * FROM Turno WHERE CodTurno ='" & sTurno & "'"
       Me.adoTurno.Refresh
       
       bFinSemana = True
       dHoraEntrada = TimeSerial(CInt(Mid$(Time, 1, 2)) - 1, Mid$(Time, 4, 2), 0)
       dHoraSalida = Time
 
 
 
       
 End Select
 
 
 Me.adoHorarios.Refresh
 dHMaximaEntrada = Me.adoTurno.Recordset.Fields("HMaximaEntrada")
 dHMaximaSalida = Me.adoTurno.Recordset.Fields("HMaximaSalida")
 
 
 'NumReloj = Int((2 * Rnd) + 2)
 dHora = Time()
 
 'Me.lblOtroMensaje.Caption = "Empleado: " & NumEmpleado & ", Reloj: " & NumReloj
   
   Me.adoExtraTurno.Recordset.Find "CodEmpleado ='" & NumEmpleado & "'"
   
   If Not Me.adoExtraTurno.Recordset.EOF Then
      If Me.adoExtraTurno.Recordset.Fields("FechaSalida") = Mid$(Now, 1, 10) Then
         Me.adoExtraTurno.Recordset.Fields("bActivo") = 0
         If Me.adoExtraTurno.Recordset.Fields("HoraEntrada") > Time Then
            Me.adoExtraTurno.Recordset.Fields("HorasLaboradas") = Format((DateDiff("n", Me.adoExtraTurno.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) - (Me.adoExtraTurno.Recordset.Fields("TiempoReceso") / 60), "##.##")
            Me.adoExtraTurno.Recordset.Fields("HorasLaboradas") = Me.adoExtraTurno.Recordset.Fields("HorasLaboradas") + Format((DateDiff("n", "00:00:00", Time) / 60), "##.##")
         Else
            Me.adoExtraTurno.Recordset.Fields("HorasLaboradas") = Format((DateDiff("n", Time, Me.adoExtraTurno.Recordset.Fields("HoraEntrada")) / 60) - (Me.adoExtraTurno.Recordset.Fields("TiempoReceso") / 60), "##.##")
             
         End If
                 
         Me.adoExtraTurno.Recordset.Fields("HoraSalida") = Time
         Me.adoExtraTurno.Recordset.Update
         lblCodigo1.Caption = NumEmpleado
         lblNombre1.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
         lblFecha1.Caption = Format(Now, "Long Date")
         lblHora1.Caption = Time()
         Me.lblMensaje1.ForeColor = &HFF&
         Me.lblMensaje1.Caption = "SALIDA REGISTRADA!!"
         Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
         Me.adoEntrada.Refresh
         Me.adoExtraTurno.Refresh
         Exit Sub
      End If
   End If
   
  
   
   
   Me.adoEmpleado.Recordset.Find "CodEmpleado ='" & NumEmpleado & "'"
     
    If Not Me.adoEmpleado.Recordset.EOF Then
       Me.adoEntrada.Recordset.Find "CodEmpleado ='" & NumEmpleado & "'"
       
       If Not Me.adoEntrada.Recordset.EOF Then
         lblCodigo1.Caption = NumEmpleado
             
         If (Abs(DateDiff("n", Me.adoEntrada.Recordset.Fields("HoraEntrada"), dHora)) > 60 And (DateDiff("n", dHora, dHMaximaSalida) > 0 And Me.adoEntrada.Recordset.Fields("FechaEntrada") = Fecha1)) Or (bFinSemana And Abs(DateDiff("n", Me.adoEntrada.Recordset.Fields("HoraEntrada"), dHora)) > 60 And Me.adoEntrada.Recordset.Fields("FechaEntrada") = CDate(Mid$(Now, 1, 10))) Then
            Me.adoEntrada.Recordset.Fields("FechaSalida") = Mid$(Now, 1, 10)
            Me.adoEntrada.Recordset.Fields("HoraSalida") = Time()
            Me.adoEntrada.Recordset.Fields("bActivo") = False
            'Me.adoEntrada.Recordset.Fields("bSalida") = True
            Me.adoEntrada.Recordset.Update
            Me.adoEntrada.Refresh
            lblCodigo1.Caption = NumEmpleado
            lblNombre1.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
            lblFecha1.Caption = Format(Now, "Long Date")
            lblHora1.Caption = Time()
            Me.lblMensaje1.Caption = "SALIDA REGISTRADA"
            Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
            
         ElseIf DateDiff("n", Me.adoEntrada.Recordset.Fields("HoraEntrada"), Time) > 60 And Me.adoEntrada.Recordset.Fields("FechaEntrada") = Fecha1 Then
            Me.adoEntrada.Recordset.Fields("FechaSalida") = Mid$(Now, 1, 10)
            Me.adoEntrada.Recordset.Fields("HoraSalida") = Time()
            Me.adoEntrada.Recordset.Fields("bActivo") = False
            'Me.adoEntrada.Recordset.Fields("bSalida") = True
            Me.adoEntrada.Recordset.Update
            Me.adoEntrada.Refresh
            lblCodigo1.Caption = NumEmpleado
            lblNombre1.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
            lblFecha1.Caption = Format(Now, "Long Date")
            lblHora1.Caption = Time()
            Me.lblMensaje1.ForeColor = &HFF&
            Me.lblMensaje1.Caption = "SALIDA REGISTRADA"
            Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
            
            
         ElseIf (DateDiff("n", dHora, dHMaximaSalida) < 0) And CDate(Mid$(Now, 1, 10)) - Me.adoEntrada.Recordset.Fields("FechaEntrada") >= 1 Then
                 
 
            
            Me.adoEntrada.Recordset.Fields("bActivo") = False
            'Me.adoEntrada.Recordset.Fields("bSalida") = True
            Me.adoEntrada.Recordset.Update
            'Me.adoEntrada.Refresh

            Me.adoEntrada.Refresh
            
            Me.adoEntrada.Recordset.AddNew
            Me.adoEntrada.Recordset.Fields("CodEmpleado") = NumEmpleado
            Me.adoEntrada.Recordset.Fields("FechaEntrada") = Mid$(Now, 1, 10)
            Me.adoEntrada.Recordset.Fields("CodTurno") = sTurno
            Me.adoEntrada.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
            Me.adoEntrada.Recordset.Fields("HoraEntrada") = dHEntradaHoy
        
            Me.adoEntrada.Recordset.Fields("bActivo") = True
            Me.adoEntrada.Recordset.Update
            Me.adoEntrada.Refresh
            
            lblCodigo1.Caption = NumEmpleado
            lblNombre1.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
            lblFecha1.Caption = Format(Now, "Long Date")
            lblHora1.Caption = Time()
            Me.lblMensaje1.ForeColor = &HFF&
            Me.lblMensaje1.Caption = "PERDIO EL DIA POR NO MARCAR SALIDA"
            Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
                 
                 
         Else
         
            lblCodigo1.Caption = NumEmpleado
            lblNombre1.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
            lblFecha1.Caption = Format(Now, "Long Date")
            lblHora1.Caption = Time()
            Me.lblMensaje1.ForeColor = &HFF&
            Me.lblMensaje1.Caption = "YA FUE REGISTRADO!!"
            Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
            Me.adoEntrada.Refresh
            
         End If
         
             
      ElseIf Time >= dHoraEntrada Then
         
        Me.adoSabados.Refresh
        
        Me.adoSabados.CommandType = adCmdText
        'Fecha = Format(Now, "dd/mm/yyyy")
        Fecha1 = Format(Now, "dd/mm/yyyy")
        sFecha1 = Mid$(Fecha1, 7, 4) & "-" & Mid$(Fecha1, 4, 2) & "-" & Mid$(Fecha1, 1, 2)

        
        Me.adoSabados.CommandType = adCmdText
        Me.adoSabados.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, FechaSalida, HoraSalida, bActivo, CodTurno " & _
                                     "FROM AsistenciaEmpleado WHERE (CodEmpleado = N'" & NumEmpleado & "') AND (FechaSalida = CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) ) ORDER BY HoraSalida DESC"
        Me.adoSabados.Refresh
         
        If Not Me.adoSabados.Recordset.EOF Then
           dHoraSalida = Me.adoSabados.Recordset.Fields("HoraSalida")
        End If
        
        If DateDiff("n", Time, dHMaximaEntrada) > 0 And Not bFinSemana Then
        
        Me.adoEntrada.Recordset.AddNew
        Me.adoEntrada.Recordset.Fields("CodEmpleado") = NumEmpleado
        Me.adoEntrada.Recordset.Fields("FechaEntrada") = Mid$(Now, 1, 10)
        Me.adoEntrada.Recordset.Fields("CodTurno") = sTurno
        Me.adoEntrada.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
        Me.adoEntrada.Recordset.Fields("HoraEntrada") = dHEntradaHoy
        
        Me.adoEntrada.Recordset.Fields("bActivo") = True
        Me.adoEntrada.Recordset.Update
        Me.adoEntrada.Refresh
        
        lblCodigo1.Caption = NumEmpleado
        lblNombre1.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
        lblFecha1.Caption = Format(Now, "Long Date")
        lblHora1.Caption = Time()
        Me.lblMensaje1.Caption = "ENTRADA REGISTRADA"
        Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
                              
        ElseIf bFinSemana Then
        
        Me.adoSabados.Refresh
        
        Me.adoSabados.CommandType = adCmdText
        Me.adoSabados.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, FechaSalida, HoraSalida, bActivo, CodTurno " & _
                                     "FROM AsistenciaEmpleado WHERE (CodEmpleado = N'" & NumEmpleado & "') AND (FechaEntrada = CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102))"
        Me.adoSabados.Refresh
         
         
         
        If Not Me.adoSabados.Recordset.EOF Then
           
'           Me.adoSabados.CommandType = adCmdText
'           Me.adoSabados.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, HoraSalida, FechaSalida, bActivo " & _
'                              "FROM AsistenciaEmpleado WHERE bActivo=1"
'           Me.adoSabados.Refresh
           
           
           lblCodigo1.Caption = NumEmpleado
           lblNombre1.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
           lblFecha1.Caption = Format(Now, "Long Date")
           lblHora1.Caption = Time()
           Me.lblMensaje1.ForeColor = &HFF&
           
           If Me.adoSabados.Recordset.Fields("bActivo") Then
              Me.lblMensaje1.Caption = "YA FUE REGISTRADO!!"
           Else
              Me.lblMensaje1.Caption = "SALIDA YA FUE REGISTRADA!!!!!"
           End If
           
           Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
           
           Exit Sub
           
        Else
        
          Me.adoSabados.CommandType = adCmdText
          Me.adoSabados.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, HoraSalida, FechaSalida, bActivo " & _
                              "FROM AsistenciaEmpleado WHERE bActivo=0 AND FechaSalida = CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND CodEmpleado =" & NumEmpleado
          Me.adoSabados.Refresh
          
          
          If Not Me.adoSabados.Recordset.EOF Then
          
             If DateDiff("n", dHMaximaEntrada, Time) > 0 Then
                Me.adoEntrada.Recordset.AddNew
                Me.adoEntrada.Recordset.Fields("CodEmpleado") = NumEmpleado
                Me.adoEntrada.Recordset.Fields("FechaEntrada") = Mid$(Now, 1, 10)
                Me.adoEntrada.Recordset.Fields("CodTurno") = sTurno
                Me.adoEntrada.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
                Me.adoEntrada.Recordset.Fields("HoraEntrada") = dHEntradaHoy
        
                Me.adoEntrada.Recordset.Fields("bActivo") = True
                Me.adoEntrada.Recordset.Update
                Me.adoEntrada.Refresh
        
                lblCodigo1.Caption = NumEmpleado
                lblNombre1.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
                lblFecha1.Caption = Format(Now, "Long Date")
                lblHora1.Caption = Time()
                Me.lblMensaje1.Caption = "ENTRADA REGISTRADA"
                Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
             
             Else
             
                lblCodigo1.Caption = NumEmpleado
                lblNombre1.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
                lblFecha1.Caption = Format(Now, "Long Date")
                lblHora1.Caption = Time()
                Me.lblMensaje1.ForeColor = &HFF&
                Me.lblMensaje1.Caption = "NO SE PUEDE ENTRAR ANTES DE SU HORA!!!!!"
                Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
             End If
             
          Else
             Me.adoEntrada.Recordset.AddNew
             Me.adoEntrada.Recordset.Fields("CodEmpleado") = NumEmpleado
             Me.adoEntrada.Recordset.Fields("FechaEntrada") = Mid$(Now, 1, 10)
             Me.adoEntrada.Recordset.Fields("CodTurno") = sTurno
             Me.adoEntrada.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
             Me.adoEntrada.Recordset.Fields("HoraEntrada") = dHEntradaHoy
        
             Me.adoEntrada.Recordset.Fields("bActivo") = True
             Me.adoEntrada.Recordset.Update
             Me.adoEntrada.Refresh
        
             lblCodigo1.Caption = NumEmpleado
             lblNombre1.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
             lblFecha1.Caption = Format(Now, "Long Date")
             lblHora1.Caption = Time()
             Me.lblMensaje1.Caption = "ENTRADA REGISTRADA"
             Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
          
          
             
          End If
        
        End If
        
         
'        Me.adoEntrada.Recordset.AddNew
'        Me.adoEntrada.Recordset.Fields("CodEmpleado") = NumEmpleado
'        Me.adoEntrada.Recordset.Fields("FechaEntrada") = Mid$(Now, 1, 10)
'        Me.adoEntrada.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
'        Me.adoEntrada.Recordset.Fields("HoraEntrada") = Time
'        Me.adoEntrada.Recordset.Fields("bActivo") = True
'        Me.adoEntrada.Recordset.Fields("CodTurno") = sTurno
'        Me.adoEntrada.Recordset.Update
'        Me.adoEntrada.Refresh
'
'        lblCodigo2.Caption = NumCodigo
'        lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
'        lblFecha2.Caption = Format(Now, "Long Date")
'        lblHora2.Caption = Time()
'        Me.lblMensaje2.Caption = "ENTRADA REGISTRADA"
'        Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
'
        
      Else
      
        sFecha1 = Mid$(Fecha1, 7, 4) & "-" & Mid$(Fecha1, 4, 2) & "-" & Mid$(Fecha1, 1, 2)
        
        Me.adoSabados.CommandType = adCmdText
        Me.adoSabados.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, HoraSalida, FechaSalida, bActivo " & _
                              "FROM AsistenciaEmpleado WHERE bActivo=0 AND FechaSalida = CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND CodEmpleado =" & NumEmpleado
        Me.adoSabados.Refresh
        
        If Not Me.adoSabados.Recordset.EOF Then
          lblCodigo1.Caption = NumEmpleado
          lblNombre1.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
          lblFecha1.Caption = Format(Now, "Long Date")
          lblHora1.Caption = Time()
          Me.lblMensaje1.ForeColor = &HFF&
          Me.lblMensaje1.Caption = "YA FUE REGISTRADO!!"
          Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
          
        Else
          lblCodigo1.Caption = NumEmpleado
          lblNombre1.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
          lblFecha1.Caption = Format(Now, "Long Date")
          lblHora1.Caption = Time()
          Me.lblMensaje1.ForeColor = &HFF&
          Me.lblMensaje1.Caption = "PERDIO EL DIA POR NO MARCAR ENTRADA"
          Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
          
        End If
        
        
        End If
                              
                              
      Else
        lblCodigo1.Caption = NumEmpleado
        lblNombre1.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
        lblFecha1.Caption = Format(Now, "Long Date")
        lblHora1.Caption = Time()
        Me.lblMensaje1.ForeColor = &HFF&
        Me.lblMensaje1.Caption = "NO SE PUEDE ENTRAR ANTES DE SU HORA!!!!!"
        Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
        
                              
      End If
       
   Else
   
     lblCodigo1.Caption = NumEmpleado
     lblNombre1.Caption = "No Encontrado"
     lblFecha1.Caption = Format(Now, "Long Date")
     lblHora1.Caption = Time()
     Me.lblMensaje1.ForeColor = &H40&
     Me.lblMensaje1.Caption = "EMPLEADO NO ENCONTRADO"
     Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\NoDisponible.jpg")
  End If

  Me.adoEntrada.Refresh
  Me.adoEmpleado.Refresh
       
End Select

Exit Sub
  
TratarError:
  If Err.Number = 53 Then
     Me.imgReloj1.Picture = LoadPicture(App.Path & "\Fotos\NoDisponible.jpg")
  End If
  Me.adoEmpleado.Refresh
  Me.adoEntrada.Refresh
  Me.adoHorarios.Refresh



End Sub

Private Sub mscReloj2_OnComm()


Dim NumEmpleado As Variant
Dim Fecha As Long, Fecha1 As Date
Dim Longitud As Byte
Dim Tiempo As Long
Dim dHoraEntrada As Variant
Dim dHoraSalida As Variant
Dim dHEntradaHoy As Date
Dim dHSalidaHoy As Date
Dim bFinSemana As Boolean
Dim sDia As String
Dim sFecha1 As Variant
Dim dHMaximaEntrada As Date
Dim sHMaximaSalida As Date

 
On Error GoTo TratarError

 Me.lblMensaje2.ForeColor = &HC00000
 Fecha1 = Format(Now, "dd/mm/yyyy")
 
Select Case Me.mscReloj2.CommEvent

   Case CommBreak


   Case 2

NumEmpleado = Mid$(Me.mscReloj2.Input, 1, 6)
Longitud = Len(NumEmpleado)
       'Me.List1.AddItem (Mid(NumEmpleado, 1, Longitud - 2)) & " Reloj1"
'       Me.mscReloj.Output = "ATDT" & vbCr
'NumEmpleado = Mid$(NumEmpleado, 1, Longitud - 2)
       
bUbicacion = InStr(1, Format(Now, "Long Date"), ",")
sDia = Mid$(Format(Now, "Long Date"), 1, bUbicacion - 1)
 
Me.adoHorarios.Refresh
Me.adoHorarios.Recordset.Find "CodEmpleado =" & NumEmpleado
 
 
 If Me.adoHorarios.Recordset.EOF Then
    
   lblCodigo2.Caption = NumEmpleado
   Me.lblNombre2.Caption = "No Encontrado"
   Me.lblFecha2.Caption = Format(Now, "Long Date")
   Me.lblHora2.Caption = Time()
   Me.lblMensaje2.ForeColor = &H40&
   Me.lblMensaje2.Caption = "EMPLEADO NO ENCONTRADO"
   Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\NoDisponible.jpg")
 
   Exit Sub
    
 End If
 
 
 Select Case sDia
 
 Case "Lunes":
 
      If Me.adoHorarios.Recordset.Fields("MEntrada") > Time Then
         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("MEntrada")
      Else
         dHEntradaHoy = Time
      End If
      
      sTurno = Me.adoHorarios.Recordset.Fields("TurnoLunes")
      dHoraEntrada = TimeSerial(CInt(Mid$(Me.adoHorarios.Recordset.Fields("MEntrada"), 1, 2)) - 1, Mid$(Me.adoHorarios.Recordset.Fields("MEntrada"), 4, 2), 0)
      dHoraSalida = Me.adoHorarios.Recordset.Fields("LSalida")
 
 Case "Martes":
 
      If Me.adoHorarios.Recordset.Fields("MEntrada") > Time Then
         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("MEntrada")
      Else
         dHEntradaHoy = Time
      End If
      
      sTurno = Me.adoHorarios.Recordset.Fields("TurnoMartes")
      dHoraEntrada = TimeSerial(CInt(Mid$(Me.adoHorarios.Recordset.Fields("MEntrada"), 1, 2)) - 1, Mid$(Me.adoHorarios.Recordset.Fields("MEntrada"), 4, 2), 0)
      dHoraSalida = Me.adoHorarios.Recordset.Fields("MSalida")
      
 Case "Miércoles":
 
      If Me.adoHorarios.Recordset.Fields("MCEntrada") > Time Then
         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("MCEntrada")
      Else
         dHEntradaHoy = Time
      End If
      
      sTurno = Me.adoHorarios.Recordset.Fields("TurnoMiercoles")
      dHoraEntrada = TimeSerial(CInt(Mid$(Me.adoHorarios.Recordset.Fields("MCEntrada"), 1, 2)) - 1, Mid$(Me.adoHorarios.Recordset.Fields("MCEntrada"), 4, 2), 0)
      dHoraSalida = Me.adoHorarios.Recordset.Fields("MCSalida")
      
 Case "Jueves":
 
      If Me.adoHorarios.Recordset.Fields("MEntrada") > Time Then
         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("JEntrada")
      Else
         dHEntradaHoy = Time
      End If
      
      
      sTurno = Me.adoHorarios.Recordset.Fields("TurnoJueves")
      dHoraEntrada = TimeSerial(CInt(Mid$(Me.adoHorarios.Recordset.Fields("JEntrada"), 1, 2)) - 1, Mid$(Me.adoHorarios.Recordset.Fields("JEntrada"), 4, 2), 0)
      dHoraSalida = Me.adoHorarios.Recordset.Fields("JSalida")
      
 Case "Viernes":
 
      If Me.adoHorarios.Recordset.Fields("VEntrada") > Time Then
         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("MEntrada")
      Else
         dHEntradaHoy = Time
      End If
      
      
      sTurno = Me.adoHorarios.Recordset.Fields("TurnoViernes")
      dHoraEntrada = TimeSerial(CInt(Mid$(Me.adoHorarios.Recordset.Fields("VEntrada"), 1, 2)) - 1, Mid$(Me.adoHorarios.Recordset.Fields("VEntrada"), 4, 2), 0)
      dHoraSalida = Me.adoHorarios.Recordset.Fields("VSalida")
      
 Case "Sábado":
        
       If Me.adoHorarios.Recordset.Fields("DEntrada") > Time Then
         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("DEntrada")
      Else
         dHEntradaHoy = Time
      End If
      
       sTurno = Me.adoHorarios.Recordset.Fields("TurnoSabado")
       bFinSemana = True
       dHoraEntrada = TimeSerial(CInt(Mid$(Time, 1, 2)) - 1, Mid$(Time, 4, 2), 0)
       dHoraSalida = Time
       
 Case "Domingo":
       
      If Me.adoHorarios.Recordset.Fields("DEntrada") > Time Then
         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("DEntrada")
      Else
         dHEntradaHoy = Time
      End If
       
       sTurno = Me.adoHorarios.Recordset.Fields("TurnoDomingo")
       bFinSemana = True
       dHoraEntrada = TimeSerial(CInt(Mid$(Time, 1, 2)) - 1, Mid$(Time, 4, 2), 0)
       dHoraSalida = Time
 
 
 
       
 End Select
 
 
   Me.adoHorarios.Refresh
 
   Me.adoTurno.CommandType = adCmdText
   Me.adoTurno.RecordSource = "SELECT * FROM Turno WHERE CodTurno ='" & sTurno & "'"
   Me.adoTurno.Refresh
   
   dHMaximaEntrada = Me.adoTurno.Recordset.Fields("HMaximaEntrada")
   dHMaximaSalida = Me.adoTurno.Recordset.Fields("HMaximaSalida")
 
 'NumReloj = Int((2 * Rnd) + 2)
 dHora = Time()
 
 'Me.lblOtroMensaje.Caption = "Empleado: " & NumEmpleado & ", Reloj: " & NumReloj
 
    Me.adoExtraTurno.Recordset.Find "CodEmpleado ='" & NumEmpleado & "'"
   
   If Not Me.adoExtraTurno.Recordset.EOF Then
      If Me.adoExtraTurno.Recordset.Fields("FechaSalida") = Mid$(Now, 1, 10) Then
         Me.adoExtraTurno.Recordset.Fields("bActivo") = 0
         If Me.adoExtraTurno.Recordset.Fields("HoraEntrada") > Time Then
            Me.adoExtraTurno.Recordset.Fields("HorasLaboradas") = Format((DateDiff("n", Me.adoExtraTurno.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) - (Me.adoExtraTurno.Recordset.Fields("TiempoReceso") / 60), "##.##")
            Me.adoExtraTurno.Recordset.Fields("HorasLaboradas") = Me.adoExtraTurno.Recordset.Fields("HorasLaboradas") + Format((DateDiff("n", "00:00:00", Time) / 60), "##.##")
         Else
            Me.adoExtraTurno.Recordset.Fields("HorasLaboradas") = Format((DateDiff("n", Time, Me.adoExtraTurno.Recordset.Fields("HoraEntrada")) / 60) - (Me.adoExtraTurno.Recordset.Fields("TiempoReceso") / 60), "##.##")
             
         End If
                 
         Me.adoExtraTurno.Recordset.Fields("HoraSalida") = Time
         Me.adoExtraTurno.Recordset.Update
         lblCodigo2.Caption = NumEmpleado
         lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
         lblFecha2.Caption = Format(Now, "Long Date")
         lblHora2.Caption = Time()
         Me.lblMensaje2.ForeColor = &HFF&
         Me.lblMensaje2.Caption = "SALIDA REGISTRADA!!"
         Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
         Me.adoEntrada.Refresh
         Me.adoExtraTurno.Refresh
         Exit Sub
      End If
   End If
    
     
    Me.adoEmpleado.Recordset.Find "CodEmpleado ='" & NumEmpleado & "'"
     
    If Not Me.adoEmpleado.Recordset.EOF Then
       Me.adoEntrada.Recordset.Find "CodEmpleado ='" & NumEmpleado & "'"
       
       If Not Me.adoEntrada.Recordset.EOF Then
         lblCodigo2.Caption = NumEmpleado
             
         If (Abs(DateDiff("n", Me.adoEntrada.Recordset.Fields("HoraEntrada"), dHora)) > 60 And (DateDiff("n", dHora, dHMaximaSalida) > 0 And Me.adoEntrada.Recordset.Fields("FechaEntrada") = Fecha1)) Or (bFinSemana And Abs(DateDiff("n", Me.adoEntrada.Recordset.Fields("HoraEntrada"), dHora)) > 60 And Me.adoEntrada.Recordset.Fields("FechaEntrada") = CDate(Mid$(Now, 1, 10))) Then
            Me.adoEntrada.Recordset.Fields("FechaSalida") = Mid$(Now, 1, 10)
            Me.adoEntrada.Recordset.Fields("HoraSalida") = Time()
            Me.adoEntrada.Recordset.Fields("bActivo") = False
            'Me.adoEntrada.Recordset.Fields("bSalida") = True
            Me.adoEntrada.Recordset.Update
            Me.adoEntrada.Refresh
            lblCodigo2.Caption = NumEmpleado
            lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
            lblFecha2.Caption = Format(Now, "Long Date")
            lblHora2.Caption = Time()
            Me.lblMensaje2.Caption = "SALIDA REGISTRADA"
            Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
            
            
         ElseIf DateDiff("n", Me.adoEntrada.Recordset.Fields("HoraEntrada"), Time) > 60 And Me.adoEntrada.Recordset.Fields("FechaEntrada") = Fecha1 Then
            
            Me.adoEntrada.Recordset.Fields("FechaSalida") = Mid$(Now, 1, 10)
            Me.adoEntrada.Recordset.Fields("HoraSalida") = Time()
            Me.adoEntrada.Recordset.Fields("bActivo") = False
            'Me.adoEntrada.Recordset.Fields("bSalida") = True
            Me.adoEntrada.Recordset.Update
            
            lblCodigo2.Caption = NumEmpleado
            lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
            lblFecha2.Caption = Format(Now, "Long Date")
            lblHora2.Caption = Time()
            Me.lblMensaje2.ForeColor = &HFF&
            Me.lblMensaje2.Caption = "SALIDA REGISTRADA!!"
            Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
            Me.adoEntrada.Refresh
            
            
         ElseIf (DateDiff("n", dHora, dHMaximaSalida) < 0) And CDate(Mid$(Now, 1, 10)) - Me.adoEntrada.Recordset.Fields("FechaEntrada") >= 1 Then
                 
'
            
            Me.adoEntrada.Recordset.Fields("bActivo") = False
            'Me.adoEntrada.Recordset.Fields("bSalida") = True
            Me.adoEntrada.Recordset.Update
            'Me.adoEntrada.Refresh

            Me.adoEntrada.Refresh
            
            Me.adoEntrada.Recordset.AddNew
            Me.adoEntrada.Recordset.Fields("CodEmpleado") = NumEmpleado
            Me.adoEntrada.Recordset.Fields("FechaEntrada") = Mid$(Now, 1, 10)
            Me.adoEntrada.Recordset.Fields("CodTurno") = sTurno
            Me.adoEntrada.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
            Me.adoEntrada.Recordset.Fields("HoraEntrada") = dHEntradaHoy
        
            Me.adoEntrada.Recordset.Fields("bActivo") = True
            Me.adoEntrada.Recordset.Update
            Me.adoEntrada.Refresh
            
            
            lblCodigo2.Caption = NumEmpleado
            lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
            lblFecha2.Caption = Format(Now, "Long Date")
            lblHora2.Caption = Time()
            Me.lblMensaje2.ForeColor = &HFF&
            Me.lblMensaje2.Caption = "PERDIO EL DIA POR NO MARCAR SALIDA"
            Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
            Me.adoEntrada.Refresh
                 
         Else
         
            lblCodigo2.Caption = NumEmpleado
            lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
            lblFecha2.Caption = Format(Now, "Long Date")
            lblHora2.Caption = Time()
            Me.lblMensaje2.ForeColor = &HFF&
            Me.lblMensaje2.Caption = "YA FUE REGISTRADO!!"
            Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
            Me.adoEntrada.Refresh
            
         End If
         
             
      ElseIf Time >= dHoraEntrada Then
         
        Me.adoSabados.Refresh
        
        Fecha1 = Format(Now, "dd/mm/yyyy")
        Fecha = Fecha1
        
        sFecha1 = Mid$(Fecha1, 7, 4) & "-" & Mid$(Fecha1, 4, 2) & "-" & Mid$(Fecha1, 1, 2)

        
        Me.adoSabados.CommandType = adCmdText
        Me.adoSabados.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, FechaSalida, HoraSalida, bActivo, CodTurno " & _
                                     "FROM AsistenciaEmpleado WHERE (CodEmpleado = N'" & NumEmpleado & "') AND (FechaSalida = CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) ) ORDER BY HoraSalida DESC"
        Me.adoSabados.Refresh
         
        If Not Me.adoSabados.Recordset.EOF Then
           dHoraSalida = Me.adoSabados.Recordset.Fields("HoraSalida")
        End If
        
        If DateDiff("n", Time, dHMaximaEntrada) > 0 And Not bFinSemana Then
        
        Me.adoEntrada.Recordset.AddNew
        Me.adoEntrada.Recordset.Fields("CodEmpleado") = NumEmpleado
        Me.adoEntrada.Recordset.Fields("FechaEntrada") = Mid$(Now, 1, 10)
        Me.adoEntrada.Recordset.Fields("CodTurno") = sTurno
        Me.adoEntrada.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
        Me.adoEntrada.Recordset.Fields("HoraEntrada") = dHEntradaHoy
        
        Me.adoEntrada.Recordset.Fields("bActivo") = True
        Me.adoEntrada.Recordset.Update
        Me.adoEntrada.Refresh
        
        lblCodigo2.Caption = NumEmpleado
        lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
        lblFecha2.Caption = Format(Now, "Long Date")
        lblHora2.Caption = Time()
        Me.lblMensaje2.Caption = "ENTRADA REGISTRADA"
        Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
                              
        ElseIf bFinSemana Then
        
        Me.adoSabados.Refresh
        
        Me.adoSabados.CommandType = adCmdText
        Me.adoSabados.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, FechaSalida, HoraSalida, bActivo, CodTurno " & _
                                     "FROM AsistenciaEmpleado WHERE (CodEmpleado = N'" & NumEmpleado & "') AND (FechaEntrada = CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102))"
        Me.adoSabados.Refresh
         
         
        If Not Me.adoSabados.Recordset.EOF Then
           
'           Me.adoSabados.CommandType = adCmdText
'           Me.adoSabados.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, HoraSalida, FechaSalida, bActivo " & _
'                              "FROM AsistenciaEmpleado WHERE bActivo=1"
'           Me.adoSabados.Refresh
           
           
           lblCodigo2.Caption = NumEmpleado
           lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
           lblFecha2.Caption = Format(Now, "Long Date")
           lblHora2.Caption = Time()
           Me.lblMensaje2.ForeColor = &HFF&
           
           If Me.adoSabados.Recordset.Fields("bActivo") Then
              Me.lblMensaje2.Caption = "YA FUE REGISTRADO!!!"
           Else
              Me.lblMensaje2.Caption = "SALIDA YA FUE REGISTRADA!!!!!"
           End If
           
           Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
           
           Exit Sub
           
        Else
        
          Me.adoSabados.CommandType = adCmdText
          Me.adoSabados.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, HoraSalida, FechaSalida, bActivo " & _
                              "FROM AsistenciaEmpleado WHERE bActivo=0 AND FechaSalida = CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND CodEmpleado =" & NumEmpleado
          Me.adoSabados.Refresh
          
          If Not Me.adoSabados.Recordset.EOF Then
          
             If DateDiff("n", dHMaximaEntrada, Time) > 0 Then
                Me.adoEntrada.Recordset.AddNew
                Me.adoEntrada.Recordset.Fields("CodEmpleado") = NumEmpleado
                Me.adoEntrada.Recordset.Fields("FechaEntrada") = Mid$(Now, 1, 10)
                Me.adoEntrada.Recordset.Fields("CodTurno") = sTurno
                Me.adoEntrada.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
                Me.adoEntrada.Recordset.Fields("HoraEntrada") = dHEntradaHoy
        
                Me.adoEntrada.Recordset.Fields("bActivo") = True
                Me.adoEntrada.Recordset.Update
                Me.adoEntrada.Refresh
        
                lblCodigo2.Caption = NumEmpleado
                lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
                lblFecha2.Caption = Format(Now, "Long Date")
                lblHora2.Caption = Time()
                Me.lblMensaje2.Caption = "ENTRADA REGISTRADA"
                Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
             
             Else
             
                lblCodigo2.Caption = NumEmpleado
                lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
                lblFecha2.Caption = Format(Now, "Long Date")
                lblHora2.Caption = Time()
                Me.lblMensaje2.ForeColor = &HFF&
                Me.lblMensaje2.Caption = "NO SE PUEDE ENTRAR ANTES DE SU HORA!!!!!"
                Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
             End If
             
             
          Else
             Me.adoEntrada.Recordset.AddNew
             Me.adoEntrada.Recordset.Fields("CodEmpleado") = NumEmpleado
             Me.adoEntrada.Recordset.Fields("FechaEntrada") = Mid$(Now, 1, 10)
             Me.adoEntrada.Recordset.Fields("CodTurno") = sTurno
             Me.adoEntrada.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
             Me.adoEntrada.Recordset.Fields("HoraEntrada") = dHEntradaHoy
        
             Me.adoEntrada.Recordset.Fields("bActivo") = True
             Me.adoEntrada.Recordset.Update
             Me.adoEntrada.Refresh
        
             lblCodigo2.Caption = NumEmpleado
             lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
             lblFecha2.Caption = Format(Now, "Long Date")
             lblHora2.Caption = Time()
             Me.lblMensaje2.Caption = "ENTRADA REGISTRADA"
             Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
          
          
             
                     
          End If
        
        End If
        
         
'        Me.adoEntrada.Recordset.AddNew
'        Me.adoEntrada.Recordset.Fields("CodEmpleado") = NumEmpleado
'        Me.adoEntrada.Recordset.Fields("FechaEntrada") = Mid$(Now, 1, 10)
'        Me.adoEntrada.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
'        Me.adoEntrada.Recordset.Fields("HoraEntrada") = Time
'        Me.adoEntrada.Recordset.Fields("bActivo") = True
'        Me.adoEntrada.Recordset.Fields("CodTurno") = sTurno
'        Me.adoEntrada.Recordset.Update
'        Me.adoEntrada.Refresh
'
'        lblCodigo2.Caption = NumCodigo
'        lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
'        lblFecha2.Caption = Format(Now, "Long Date")
'        lblHora2.Caption = Time()
'        Me.lblMensaje2.Caption = "ENTRADA REGISTRADA"
'        Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
'
          
        
        Else
        
        sFecha1 = Mid$(Fecha1, 7, 4) & "-" & Mid$(Fecha1, 4, 2) & "-" & Mid$(Fecha1, 1, 2)
        
        Me.adoSabados.CommandType = adCmdText
        Me.adoSabados.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, HoraSalida, FechaSalida, bActivo " & _
                              "FROM AsistenciaEmpleado WHERE bActivo=0 AND FechaSalida = CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND CodEmpleado =" & NumEmpleado
        Me.adoSabados.Refresh
        
        If Not Me.adoSabados.Recordset.EOF Then
          lblCodigo2.Caption = NumEmpleado
          lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
          lblFecha2.Caption = Format(Now, "Long Date")
          lblHora2.Caption = Time()
          Me.lblMensaje2.ForeColor = &HFF&
          Me.lblMensaje2.Caption = "YA FUE REGISTRADO!!"
          Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
          
        Else
          lblCodigo2.Caption = NumEmpleado
          lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
          lblFecha2.Caption = Format(Now, "Long Date")
          lblHora2.Caption = Time()
          Me.lblMensaje2.ForeColor = &HFF&
          Me.lblMensaje2.Caption = "PERDIO EL DIA POR NO MARCAR ENTRADA"
          Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
          
        End If
        
        
        End If
                              
                              
      Else
        lblCodigo2.Caption = NumEmpleado
        lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
        lblFecha2.Caption = Format(Now, "Long Date")
        lblHora2.Caption = Time()
        Me.lblMensaje2.ForeColor = &HFF&
        Me.lblMensaje2.Caption = "NO SE PUEDE ENTRAR ANTES DE SU HORA!!!!!"
        Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
        
                              
      End If
       
   Else
   
     lblCodigo2.Caption = NumEmpleado
     lblNombre2.Caption = "No Encontrado"
     lblFecha2.Caption = Format(Now, "Long Date")
     lblHora2.Caption = Time()
     Me.lblMensaje2.ForeColor = &H40&
     Me.lblMensaje2.Caption = "EMPLEADO NO ENCONTRADO"
     Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\NoDisponible.jpg")
  End If
'End If
  
  Me.adoEmpleado.Refresh
  Me.adoEntrada.Refresh
       
End Select
 

Exit Sub
  
TratarError:
  If Err.Number = 53 Then
     Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\NoDisponible.jpg")
  End If
  Me.adoEmpleado.Refresh
  Me.adoEntrada.Refresh
  Me.adoHorarios.Refresh

End Sub

Private Sub tmrReloj_Timer()

Me.lblHora.Caption = Time

Me.lblFecha.Caption = Format(Mid$(Now, 1, 10), "Long Date")
 
'Dim NumEmpleado As Variant
'Dim Longitud As Byte
'Dim Tiempo As Long
'Dim dHoraEntrada As Variant
'Dim dHoraSalida As Variant
'Dim dHEntradaHoy As Date
'Dim dHSalidaHoy As Date
'Dim bFinSemana As Boolean
'Dim sTurno As String
'
'On Error GoTo TratarError
'
' Me.lblMensaje2.ForeColor = &HC00000
'
''Select Case Me.mscReloj2.CommEvent
''
''   Case CommBreak
''
''
''   Case 2
'
''NumEmpleado = Me.txtCodigo.Text
'Longitud = Len(NumEmpleado)
'       'Me.List1.AddItem (Mid(NumEmpleado, 1, Longitud - 2)) & " Reloj1"
''       Me.mscReloj.Output = "ATDT" & vbCr
''NumEmpleado = Mid$(NumEmpleado, 1, Longitud - 2)
'
'bUbicacion = InStr(1, Format(Now, "Long Date"), ",")
'sDia = Mid$(Format(Now, "Long Date"), 1, bUbicacion - 1)
'
' Me.adoHorarios.Recordset.Find "CodEmpleado =" & NumEmpleado
'
'
' If Me.adoHorarios.Recordset.EOF Then
'
'   lblCodigo2.Caption = NumEmpleado
'   Me.lblNombre2.Caption = "No Encontrado"
'   Me.lblFecha2.Caption = Format(Now, "Long Date")
'   Me.lblHora2.Caption = Time()
'   Me.lblMensaje2.ForeColor = &H40&
'   Me.lblMensaje2.Caption = "EMPLEADO NO ENCONTRADO"
'   Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\NoDisponible.jpg")
'
'   Exit Sub
'
' End If
'
'
' Select Case sDia
'
' Case "Lunes":
'
'      If Me.adoHorarios.Recordset.Fields("MEntrada") > Time Then
'         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("MEntrada")
'      Else
'         dHEntradaHoy = Time
'      End If
'
'      sTurno = Me.adoHorarios.Recordset.Fields("TurnoLunes")
'      dHoraEntrada = TimeSerial(CInt(Mid$(Me.adoHorarios.Recordset.Fields("MEntrada"), 1, 2)) - 1, Mid$(Me.adoHorarios.Recordset.Fields("MEntrada"), 4, 2), 0)
'      dHoraSalida = Me.adoHorarios.Recordset.Fields("LSalida")
'
' Case "Martes":
'
'      If Me.adoHorarios.Recordset.Fields("MEntrada") > Time Then
'         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("MEntrada")
'      Else
'         dHEntradaHoy = Time
'      End If
'
'      sTurno = Me.adoHorarios.Recordset.Fields("TurnoMartes")
'      dHoraEntrada = TimeSerial(CInt(Mid$(Me.adoHorarios.Recordset.Fields("MEntrada"), 1, 2)) - 1, Mid$(Me.adoHorarios.Recordset.Fields("MEntrada"), 4, 2), 0)
'      dHoraSalida = Me.adoHorarios.Recordset.Fields("MSalida")
'
' Case "Miércoles":
'
'      If Me.adoHorarios.Recordset.Fields("MCEntrada") > Time Then
'         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("MCEntrada")
'      Else
'         dHEntradaHoy = Time
'      End If
'
'      sTurno = Me.adoHorarios.Recordset.Fields("TurnoMiercoles")
'      dHoraEntrada = TimeSerial(CInt(Mid$(Me.adoHorarios.Recordset.Fields("MCEntrada"), 1, 2)) - 1, Mid$(Me.adoHorarios.Recordset.Fields("MCEntrada"), 4, 2), 0)
'      dHoraSalida = Me.adoHorarios.Recordset.Fields("MCSalida")
'
' Case "Jueves":
'
'      If Me.adoHorarios.Recordset.Fields("MEntrada") > Time Then
'         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("JEntrada")
'      Else
'         dHEntradaHoy = Time
'      End If
'
'
'      sTurno = Me.adoHorarios.Recordset.Fields("TurnoJueves")
'      dHoraEntrada = TimeSerial(CInt(Mid$(Me.adoHorarios.Recordset.Fields("JEntrada"), 1, 2)) - 1, Mid$(Me.adoHorarios.Recordset.Fields("JEntrada"), 4, 2), 0)
'      dHoraSalida = Me.adoHorarios.Recordset.Fields("JSalida")
'
' Case "Viernes":
'
'      If Me.adoHorarios.Recordset.Fields("VEntrada") > Time Then
'         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("MEntrada")
'      Else
'         dHEntradaHoy = Time
'      End If
'
'
'      sTurno = Me.adoHorarios.Recordset.Fields("TurnoViernes")
'      dHoraEntrada = TimeSerial(CInt(Mid$(Me.adoHorarios.Recordset.Fields("VEntrada"), 1, 2)) - 1, Mid$(Me.adoHorarios.Recordset.Fields("VEntrada"), 4, 2), 0)
'      dHoraSalida = Me.adoHorarios.Recordset.Fields("VSalida")
'
' Case "Sábado":
'
'       If Me.adoHorarios.Recordset.Fields("DEntrada") > Time Then
'         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("DEntrada")
'      Else
'         dHEntradaHoy = Time
'      End If
'
'       sTurno = Me.adoHorarios.Recordset.Fields("TurnoSabado")
'       bFinSemana = True
'       dHoraEntrada = TimeSerial(CInt(Mid$(Time, 1, 2)) - 1, Mid$(Time, 4, 2), 0)
'       dHoraSalida = Time
'
' Case "Domingo":
'
'      If Me.adoHorarios.Recordset.Fields("DEntrada") > Time Then
'         dHEntradaHoy = Me.adoHorarios.Recordset.Fields("DEntrada")
'      Else
'         dHEntradaHoy = Time
'      End If
'
'       sTurno = Me.adoHorarios.Recordset.Fields("TurnoDomingo")
'       bFinSemana = True
'       dHoraEntrada = TimeSerial(CInt(Mid$(Time, 1, 2)) - 1, Mid$(Time, 4, 2), 0)
'       dHoraSalida = Time
'
'
'
'
' End Select
'
'
' Me.adoHorarios.Refresh
'
' 'NumReloj = Int((2 * Rnd) + 2)
' dHora = Time()
'
' 'Me.lblOtroMensaje.Caption = "Empleado: " & NumEmpleado & ", Reloj: " & NumReloj
'
'
'' If NumReloj = 2 Then
'
'    Me.adoEmpleado.Recordset.Find "CodEmpleado ='" & NumEmpleado & "'"
'
'    If Not Me.adoEmpleado.Recordset.EOF Then
'       Me.adoEntrada.Recordset.Find "CodEmpleado ='" & NumEmpleado & "'"
'
'       If Not Me.adoEntrada.Recordset.EOF Then
'         lblCodigo2.Caption = NumEmpleado
'
'         If Abs(DateDiff("n", Me.adoEntrada.Recordset.Fields("HoraEntrada"), dHora)) > 60 Or (bFinSemana And Abs(DateDiff("n", Me.adoEntrada.Recordset.Fields("HoraEntrada"), dHora)) > 60 And Me.adoEntrada.Recordset.Fields("HoraEntrada") = CDate(Mid$(Now, 1, 10))) Then
'            Me.adoEntrada.Recordset.Fields("FechaSalida") = Mid$(Now, 1, 10)
'            Me.adoEntrada.Recordset.Fields("HoraSalida") = Time()
'            Me.adoEntrada.Recordset.Fields("bActivo") = False
'            'Me.adoEntrada.Recordset.Fields("bSalida") = True
'            Me.adoEntrada.Recordset.Update
'            Me.adoEntrada.Refresh
'            lblCodigo2.Caption = NumEmpleado
'            lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
'            lblFecha2.Caption = Format(Now, "Long Date")
'            lblHora2.Caption = Time()
'            Me.lblMensaje2.Caption = "SALIDA REGISTRADA"
'            Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
'
'
'         ElseIf Me.adoEntrada.Recordset.Fields("FechaEntrada") <> CDate(Mid$(Now, 1, 10)) Then
'
'            Me.adoEntrada.Recordset.Fields("FechaSalida") = Mid$(Now, 1, 10)
'            Me.adoEntrada.Recordset.Fields("HoraSalida") = Time()
'            Me.adoEntrada.Recordset.Fields("bActivo") = False
'            'Me.adoEntrada.Recordset.Fields("bSalida") = True
'            Me.adoEntrada.Recordset.Update
'            Me.adoEntrada.Refresh
'            lblCodigo2.Caption = NumEmpleado
'            lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
'            lblFecha2.Caption = Format(Now, "Long Date")
'            lblHora2.Caption = Time()
'            Me.lblMensaje2.Caption = "SALIDA REGISTRADA"
'            Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
'
'
'         Else
'
'            lblCodigo2.Caption = NumEmpleado
'            lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
'            lblFecha2.Caption = Format(Now, "Long Date")
'            lblHora2.Caption = Time()
'            Me.lblMensaje2.ForeColor = &HFF&
'            Me.lblMensaje2.Caption = "SU ENTRADA YA FUE REGISTRADA!!"
'            Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
'            Me.adoEntrada.Refresh
'
'         End If
'
'
'      ElseIf Time >= dHoraEntrada Then
'
'        Me.adoSabados.Refresh
'
'        Me.adoSabados.CommandType = adCmdText
'        Me.adoSabados.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, FechaSalida, HoraSalida, bActivo, CodTurno " & _
'                                     "FROM AsistenciaEmpleado WHERE (CodEmpleado = N'" & NumEmpleado & "') AND (FechaSalida = '" & CDate(Mid$(Now, 1, 10)) & "') ORDER BY HoraSalida DESC"
'        Me.adoSabados.Refresh
'
'        If Not Me.adoSabados.Recordset.EOF Then
'           dHoraSalida = Me.adoSabados.Recordset.Fields("HoraSalida")
'        End If
'
'        If Abs(DateDiff("n", dHoraSalida, Time)) > 60 Then
'
'        Me.adoEntrada.Recordset.AddNew
'        Me.adoEntrada.Recordset.Fields("CodEmpleado") = NumEmpleado
'        Me.adoEntrada.Recordset.Fields("FechaEntrada") = Mid$(Now, 1, 10)
'        Me.adoEntrada.Recordset.Fields("CodTurno") = sTurno
'        Me.adoEntrada.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
'        Me.adoEntrada.Recordset.Fields("HoraEntrada") = dHEntradaHoy
'
'        Me.adoEntrada.Recordset.Fields("bActivo") = True
'        Me.adoEntrada.Recordset.Update
'        Me.adoEntrada.Refresh
'
'        lblCodigo2.Caption = NumCodigo
'        lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
'        lblFecha2.Caption = Format(Now, "Long Date")
'        lblHora2.Caption = Time()
'        Me.lblMensaje2.Caption = "ENTRADA REGISTRADA"
'        Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
'
'        ElseIf bFinSemana Then
'
'        Me.adoSabados.Refresh
'
'        Me.adoSabados.CommandType = adCmdText
'        Me.adoSabados.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, FechaSalida, HoraSalida, bActivo, CodTurno " & _
'                                     "FROM AsistenciaEmpleado WHERE (CodEmpleado = N'" & NumEmpleado & "') AND (FechaEntrada = '" & CDate(Mid$(Now, 1, 10)) & "')"
'        Me.adoSabados.Refresh
'
'
'        If Not Me.adoSabados.Recordset.EOF Then
'
''           Me.adoSabados.CommandType = adCmdText
''           Me.adoSabados.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, HoraSalida, FechaSalida, bActivo " & _
''                              "FROM AsistenciaEmpleado WHERE bActivo=1"
''           Me.adoSabados.Refresh
'
'
'           lblCodigo2.Caption = NumEmpleado
'           lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
'           lblFecha2.Caption = Format(Now, "Long Date")
'           lblHora2.Caption = Time()
'           Me.lblMensaje2.ForeColor = &HFF&
'
'           If Me.adoSabados.Recordset.Fields("bActivo") Then
'              Me.lblMensaje2.Caption = "ENTRADA YA FUE REGISTRADA!!!"
'           Else
'              Me.lblMensaje2.Caption = "SALIDA YA FUE REGISTRADA!!!!!"
'           End If
'
'           Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
'
'           Exit Sub
'
'        Else
'
'          Me.adoSabados.CommandType = adCmdText
'          Me.adoSabados.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, HoraSalida, FechaSalida, bActivo " & _
'                              "FROM AsistenciaEmpleado WHERE bActivo=0 AND FechaSalida = '" & CDate(Mid$(Now, 1, 10)) & "'"
'          Me.adoSabados.Refresh
'
'          If Not Me.adoSabados.Recordset.EOF Then
'
'             If Abs(DateDiff("n", Me.adoSabados.Recordset.Fields("HoraSalida"), Time)) > 60 Then
'                Me.adoEntrada.Recordset.AddNew
'                Me.adoEntrada.Recordset.Fields("CodEmpleado") = NumEmpleado
'                Me.adoEntrada.Recordset.Fields("FechaEntrada") = Mid$(Now, 1, 10)
'                Me.adoEntrada.Recordset.Fields("CodTurno") = sTurno
'                Me.adoEntrada.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
'                Me.adoEntrada.Recordset.Fields("HoraEntrada") = dHEntradaHoy
'
'                Me.adoEntrada.Recordset.Fields("bActivo") = True
'                Me.adoEntrada.Recordset.Update
'                Me.adoEntrada.Refresh
'
'                lblCodigo2.Caption = NumCodigo
'                lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
'                lblFecha2.Caption = Format(Now, "Long Date")
'                lblHora2.Caption = Time()
'                Me.lblMensaje2.Caption = "ENTRADA REGISTRADA"
'                Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
'
'             Else
'
'                lblCodigo2.Caption = NumEmpleado
'                lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
'                lblFecha2.Caption = Format(Now, "Long Date")
'                lblHora2.Caption = Time()
'                Me.lblMensaje2.ForeColor = &HFF&
'                Me.lblMensaje2.Caption = "NO SE PUEDE ENTRAR ANTES DE SU HORA!!!!!"
'                Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
'             End If
'
'          End If
'
'        End If
'
'
''        Me.adoEntrada.Recordset.AddNew
''        Me.adoEntrada.Recordset.Fields("CodEmpleado") = NumEmpleado
''        Me.adoEntrada.Recordset.Fields("FechaEntrada") = Mid$(Now, 1, 10)
''        Me.adoEntrada.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
''        Me.adoEntrada.Recordset.Fields("HoraEntrada") = Time
''        Me.adoEntrada.Recordset.Fields("bActivo") = True
''        Me.adoEntrada.Recordset.Fields("CodTurno") = sTurno
''        Me.adoEntrada.Recordset.Update
''        Me.adoEntrada.Refresh
''
''        lblCodigo2.Caption = NumCodigo
''        lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
''        lblFecha2.Caption = Format(Now, "Long Date")
''        lblHora2.Caption = Time()
''        Me.lblMensaje2.Caption = "ENTRADA REGISTRADA"
''        Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
''
'
'
'        Else
'
'        lblCodigo2.Caption = NumEmpleado
'        lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
'        lblFecha2.Caption = Format(Now, "Long Date")
'        lblHora2.Caption = Time()
'        Me.lblMensaje2.ForeColor = &HFF&
'        Me.lblMensaje2.Caption = "SALIDA YA FUE REGISTRADA!!!!!"
'        Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
'
'
'
'        End If
'
'
'      Else
'        lblCodigo2.Caption = NumEmpleado
'        lblNombre2.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
'        lblFecha2.Caption = Format(Now, "Long Date")
'        lblHora2.Caption = Time()
'        Me.lblMensaje2.ForeColor = &HFF&
'        Me.lblMensaje2.Caption = "NO SE PUEDE ENTRAR ANTES DE SU HORA!!!!!"
'        Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\" & NumEmpleado & ".jpg")
'
'
'      End If
'
'   Else
'
'     lblCodigo2.Caption = NumEmpleado
'     lblNombre2.Caption = "No Encontrado"
'     lblFecha2.Caption = Format(Now, "Long Date")
'     lblHora2.Caption = Time()
'     Me.lblMensaje2.ForeColor = &H40&
'     Me.lblMensaje2.Caption = "EMPLEADO NO ENCONTRADO"
'     Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\NoDisponible.jpg")
'  End If
''End If
'
'  Me.adoEmpleado.Refresh
'
'
'
'
'Exit Sub
'
'TratarError:
'  If Err.Number = 53 Then
'     Me.imgReloj2.Picture = LoadPicture(App.Path & "\Fotos\NoDisponible.jpg")
'  End If
'  Me.adoEmpleado.Refresh
'  Me.adoEntrada.Refresh
'  Me.adoHorarios.Refresh

End Sub
