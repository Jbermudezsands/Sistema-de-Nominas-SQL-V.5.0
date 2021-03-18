VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FrmListaCompañia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Compañias"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data DtaConsulta 
      Caption         =   "DtaConsulta"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSDBCtls.DBList DBListCompañia 
      Bindings        =   "FrmListaCompañia.frx":0000
      Height          =   2400
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4233
      _Version        =   393216
      ListField       =   "NombreBD"
   End
   Begin VB.Data DtaServidor 
      Caption         =   "DtaServidor"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton CmdSeleccionar 
      Caption         =   "Seleccionar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   9015
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.Image Image2 
         Height          =   960
         Left            =   120
         Picture         =   "FrmListaCompañia.frx":001A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   9000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Top             =   120
         Width           =   645
      End
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Compañias"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   2745
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FrmListaCompañia.frx":1B5C
      Top             =   0
   End
End
Attribute VB_Name = "FrmListaCompañia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdSeleccionar_Click()
  Dim ConexionSTR1 As String, NombreBD As String
  
  NombreBD = Me.DBListCompañia.Text
  Me.DtaConsulta.RecordSource = "SELECT * FROM Servidor WHERE (NombreBD = '" & NombreBD & "')"
  Me.DtaConsulta.Refresh
  If Not Me.DtaConsulta.Recordset.EOF Then
   Conexion = Me.DtaConsulta.Recordset("Servidor")
   ConexionReporte = Me.DtaConsulta.Recordset("Servidor")
   Unload Me
   FrmListaUsuario.Show
  End If
End Sub

Private Sub Form_Load()
Dim NumeroCia As Double
Dim RutaServer As String
Me.Picture1.BackColor = RGB(239, 243, 255)
Me.Skin1.ApplySkin hWnd

RegistrarBitacora = True
RutaServer = App.Path + "\CntNominas.dll"
If Dir(RutaServer) <> "" Then

  
  With Me.dtaServidor
     .DatabaseName = RutaServer
     .RecordSource = "Servidor"
     .Refresh
  End With
  
    With Me.DtaConsulta
     .DatabaseName = RutaServer
   End With
   
 

  If Not Me.dtaServidor.Recordset.EOF Then
   Me.dtaServidor.Recordset.MoveLast
   NumeroCia = Me.dtaServidor.Recordset.RecordCount
    If NumeroCia = 1 Then
      Conexion = Me.dtaServidor.Recordset("Servidor")
      ConexionReporte = Me.dtaServidor.Recordset("Servidor")
      Unload Me
      FrmListaUsuario.Show
    End If
   End If
  
End If
End Sub
