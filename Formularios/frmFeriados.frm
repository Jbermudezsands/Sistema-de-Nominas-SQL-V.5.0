VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AF8CD3F4-666F-11D1-940D-000021A73813}#5.0#0"; "osProgress.ocx"
Begin VB.Form Form1 
   Caption         =   "Actualizar Feriado 08 de Diciembre"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin Progress.osProgress ospEmpleado 
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   5535
      _ExtentX        =   6694
      _ExtentY        =   873
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
   Begin VB.Data dtaServidor 
      Caption         =   "Servidor"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSAdodcLib.Adodc adoEmpleado 
      Height          =   375
      Left            =   1560
      Top             =   2760
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc adoAsistencia 
      Height          =   375
      Left            =   5040
      Top             =   2280
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSAdodcLib.Adodc adoTurno 
      Height          =   330
      Left            =   1560
      Top             =   2280
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
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   3720
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "&Calcular"
         Height          =   495
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalcular_Click()

Dim sDia As String
Dim bConta As Byte

Me.ospEmpleado.Max = Me.adoEmpleado.Recordset.RecordCount
Me.ospEmpleado.Min = 0


Do While Not Me.adoEmpleado.Recordset.EOF
    
   Me.ospEmpleado.Value = Me.ospEmpleado.Value + 1
   
   Me.adoTurno.CommandType = adCmdText
   Me.adoTurno.RecordSource = "SELECT CodEmpleado, MCEntrada, MCSalida, JEntrada, JSalida FROM HorarioEmpleado WHERE CodEmpleado ='" & Me.adoEmpleado.Recordset.Fields("CodEmpleado1") & "'"
   Me.adoTurno.Refresh
   
   Me.adoAsistencia.CommandType = adCmdText
   Me.adoAsistencia.RecordSource = "SELECT * FROM AsistenciaEmpleado WHERE FechaEntrada = CONVERT(DATETIME, '" & "2005-12-08" & " 00:00:00" & "', 102) AND CodEmpleado ='" & Me.adoEmpleado.Recordset.Fields("CodEmpleado1") & "'"
   Me.adoAsistencia.Refresh
   
   bConta = 1
   
  Do While bConta <= 1
   
   If Me.adoAsistencia.Recordset.EOF Then
         
      Me.adoAsistencia.Recordset.AddNew
      Me.adoAsistencia.Recordset.Fields("CodEmpleado") = Me.adoEmpleado.Recordset.Fields("CodEmpleado1")
      Me.adoAsistencia.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
      Me.adoAsistencia.Recordset.Fields("FechaEntrada") = "08/12/2005"
      Me.adoAsistencia.Recordset.Fields("HoraEntrada") = Me.adoTurno.Recordset.Fields("MCEntrada")
      Me.adoAsistencia.Recordset.Fields("FechaSalida") = "08/12/2005"
      Me.adoAsistencia.Recordset.Fields("HoraSalida") = Me.adoTurno.Recordset.Fields("MCSalida")
      Me.adoAsistencia.Recordset.Fields("bActivo") = 0
      Me.adoAsistencia.Recordset.Fields("CodTurno") = "Diurno"
      Me.adoAsistencia.Recordset.Update
   End If
      
      
'      Me.adoAsistencia.CommandType = adCmdText
'      Me.adoAsistencia.RecordSource = "SELECT * FROM AsistenciaEmpleado WHERE FechaEntrada = CONVERT(DATETIME, '" & "2005-09-15" & " 00:00:00" & "', 102) AND CodEmpleado ='" & Me.adoEmpleado.Recordset.Fields("CodEmpleado") & "'"
'      Me.adoAsistencia.Refresh
'
'   If Me.adoAsistencia.Recordset.EOF Then
'
'      Me.adoAsistencia.Recordset.AddNew
'      Me.adoAsistencia.Recordset.Fields("CodEmpleado") = Me.adoEmpleado.Recordset.Fields("CodEmpleado")
'      Me.adoAsistencia.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
'      Me.adoAsistencia.Recordset.Fields("FechaEntrada") = "15/09/2005"
'      Me.adoAsistencia.Recordset.Fields("HoraEntrada") = Me.adoTurno.Recordset.Fields("JEntrada")
'      Me.adoAsistencia.Recordset.Fields("FechaSalida") = "15/09/2005"
'      Me.adoAsistencia.Recordset.Fields("HoraSalida") = Me.adoTurno.Recordset.Fields("JSalida")
'      Me.adoAsistencia.Recordset.Fields("bActivo") = 0
'      Me.adoAsistencia.Recordset.Fields("CodTurno") = "Diurno"
'      Me.adoAsistencia.Recordset.Update
'   End If
   
   bConta = 2
   
 Loop


Me.adoEmpleado.Recordset.MoveNext

Loop

End Sub

Private Sub cmdSalir_Click()
Unload Me

End Sub

Private Sub Form_Load()

 Dim RutaServer As String
 Dim Server As String
 Dim Conexion As String
 
 RutaServer = App.Path + "\CntNominas.dll"

  With Me.dtaServidor
     .DatabaseName = RutaServer
     .RecordSource = "Servidor"
     .Refresh
  End With

  If Not IsNull(Me.dtaServidor.Recordset.Servidor) Then
   Server = Me.dtaServidor.Recordset.Servidor
  Else
   MsgBox "No se ha definido el Servidor", vbCritical, "Sistema de Nominas"
   Exit Sub
  End If

'Borrar despues la siguiente linea
'Server = "MOISES"

Conexion = "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=" & Server

Me.adoEmpleado.ConnectionString = Conexion
Me.adoTurno.ConnectionString = Conexion
Me.adoAsistencia.ConnectionString = Conexion

Me.adoTurno.CommandType = adCmdTable
Me.adoTurno.RecordSource = "HorarioEmpleado"
Me.adoTurno.Refresh


Me.adoAsistencia.CommandType = adCmdText
Me.adoAsistencia.RecordSource = "SELECT * FROM AsistenciaEmpleado"
Me.adoAsistencia.Refresh

Me.adoEmpleado.CommandType = adCmdText
Me.adoEmpleado.RecordSource = "SELECT CodEmpleado, CodEmpleado1, Activo, CodTipoNomina FROM Empleado WHERE Activo =1 AND (CodEmpleado1 <> N'IS NULL')"
Me.adoEmpleado.Refresh







End Sub
