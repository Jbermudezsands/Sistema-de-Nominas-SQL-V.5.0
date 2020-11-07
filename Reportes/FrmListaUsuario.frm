VERSION 5.00
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmListaUsuario 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios Registrados para el Acceso"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3285
   ForeColor       =   &H00EFEFEF&
   HelpContextID   =   1
   Icon            =   "FrmListaUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3285
   Begin VB.Data DtaServidor 
      Caption         =   "DtaServidor"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Width           =   3015
   End
   Begin MSDataListLib.DataList DBLUsuario 
      Bindings        =   "FrmListaUsuario.frx":0442
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3201
      _Version        =   393216
      ListField       =   "NombreUsuario"
   End
   Begin MSAdodcLib.Adodc DtaPassword 
      Height          =   375
      Left            =   600
      Top             =   4920
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
      Caption         =   "DtaPassword"
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
   Begin MSAdodcLib.Adodc DtaFecha 
      Height          =   375
      Left            =   600
      Top             =   4440
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "DtaFecha"
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
      Left            =   4680
      Top             =   2760
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.CommandButton CmdSeleccionar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Seleccionar"
      DownPicture     =   "FrmListaUsuario.frx":045C
      Height          =   375
      Left            =   240
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton CmdSalir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancelar"
      DownPicture     =   "FrmListaUsuario.frx":325E
      Height          =   375
      Left            =   1800
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
End
Attribute VB_Name = "FrmListaUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdSalir_Click()
Unload Me
End Sub
Private Sub CmdSeleccionar_Click()
FrmEntrada.TxtNombreUsuario.Text = DBLUsuario.Text
FrmEntrada.Show 1
End Sub

Private Sub DBLUsuario_DblClick()
FrmEntrada.TxtNombreUsuario.Text = DBLUsuario.Text
FrmEntrada.Show 1
End Sub

Private Sub DBLUsuario_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  FrmEntrada.TxtNombreUsuario.Text = DBLUsuario.Text
  FrmEntrada.Show 1
 End If
End Sub



Private Sub Form_Activate()
Dim Var As Double

 Me.DtaPassword.Refresh
 If DtaPassword.Recordset.EOF Then
    BaseEntrada = True
    MDIPrimero.Show
    FrmListaUsuario.CmdSalir.Value = True
  Else
    Var = DtaPassword.Recordset("CodUsuario")
 End If

End Sub

Private Sub Form_Load()

Dim TextFecha As String
Dim FechaSystem As Long
Dim unidad As String

Me.Top = 3500
Me.Left = 3500

Ruta = ""
If Dir(Ruta) <> "" Then
  RutaServer = App.Path + "\CntNominas.dll"
  
  With Me.DtaServidor
     .DatabaseName = RutaServer
     .RecordSource = "Servidor"
     .Refresh
  End With
  
  If Not IsNull(Me.DtaServidor.Recordset.Servidor) Then
   Server = Me.DtaServidor.Recordset.Servidor
  Else
   MsgBox "No se ha definido el Servidor", vbCritical, "Sistmea de Nominas"
   Exit Sub
  End If
  
  RutaLogo = App.Path + "\fotos\Zw.bmp"
  Ruta = App.Path + "\nominas.log"
  RutaFoto = App.Path + "\fotos\"
  RutaIconos = App.Path + "\Iconos\"
  
  Conexion = "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=" & Server
'  Conexion = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas;Data Source=" & Server
  'Conexion =   ";DATABASENAME=" + Ruta + ";UID=Administrador;PWD=15081977"
   'ConexionReporte = "DSN=MS Access Database;DBQ=" + Ruta + ";DefaultDir=" + Ruta + ";DriverId=281;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;PWD=15081977;UID=admin;"
  ConexionReporte = "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=" & Server
  unidad = App.Path + "\"
  unidad = Mid(unidad, 1, 3)
ElseIf Dir(Ruta) <> "" Then
'  RutaLogo = App.Path + "\fotos\Zw.bmp"
'  Ruta = App.Path + "\nominas.log"
'  RutaFoto = App.Path + "\fotos\"
'  RutaIconos = App.Path + "\Iconos\"
'  Conexion = ";DATABASENAME=" + Ruta + ";UID=Administrador;PWD=15081977"
'  ConexionReporte = "DSN=MS Access Database;DBQ=" + Ruta + ";DefaultDir=" + Ruta + ";DriverId=281;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;PWD=15081977;UID=admin;"
'  unidad = App.Path + "\"
'  unidad = Mid(unidad, 1, 3)
   'Unidad = "Z:\"
   'Ruta = "Z:\Nominas\Nominas.log"
   'RutaFoto = "Z:\Nominas\fotos\"
    Conexion = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas;Data Source=localhost"
   'ConexionReporte = "DSN=MS Access Database;DBQ=" + Ruta + ";DefaultDir=" + Ruta + ";DriverId=281;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;PWD=15081977;UID=admin;"
Else
  MsgBox "La base de Datos no Existe", vbCritical, "Sistema de Nominas"
  Exit Sub
End If

With Me.DtaPassword
    .ConnectionString = Conexion
    .RecordSource = "Usuarios"
    .Refresh
End With

With Me.DtaFecha
   .ConnectionString = Conexion
   .RecordSource = "Entrada"
   .Refresh
End With

DtaFecha.Refresh
TextFecha = DtaFecha.Recordset("Fentrada")
TextFecha = Decrypt(TextFecha)
If Not IsDate(TextFecha) Then End

FechaSystem = CDate(TextFecha)
Dim Vol As String * 256, FileSystem As String * 256
Dim Longitud As Long, NumSerie As Long, Flags As Long
Dim Serie As String

Call GetVolumeInformation(unidad, Vol, 256, NumSerie, Longitud, Flags, FileSystem, 256)
Serie = Str(NumSerie)
'MsgBox Format(FechaSystem, "dd/mm/yyyy")
'If FechaSystem + 15 < Now Or FechaSystem > Now Then
'  DtaFecha.Recordset.MoveNext
'  'MsgBox DtaFecha.Recordset.Fentrada
'  'MsgBox Decrypt(DtaFecha.Recordset.Fentrada)
'  If Trim(Decrypt(DtaFecha.Recordset("Fentrada"))) <> Trim(Serie) Then
'  Mensaje = "Esta Copia del Sistema Necesita la Licencia para poder seguir Funcionando," & vbCr
'  Mensaje = Mensaje + " Por Favor p�ngase en contacto con Juan G. S.A. al Tel�fono 8502372, Gracias"
'  MsgBox Mensaje
'  End
'  End If
'
'End If



End Sub