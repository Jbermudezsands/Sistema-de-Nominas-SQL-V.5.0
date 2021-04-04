VERSION 5.00
Begin VB.Form FrmLicencia 
   Caption         =   "Form1"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   Icon            =   "FrmLicencia.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.Data DtaEntrada 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Nominas\Nominas.log"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Entrada"
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton CmdDias 
      Caption         =   "Temporal"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox TxtSerie 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3480
      Width           =   3375
   End
   Begin VB.CommandButton CmdLicencia 
      Caption         =   "Licencia"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   1935
      Index           =   6
      Left            =   0
      Picture         =   "FrmLicencia.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Nominas"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   2685
   End
End
Attribute VB_Name = "FrmLicencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdDias_Click()
TxtSerie.Text = Format(Now, "dd/mm/yyyy")
TxtSerie.Text = Encrypt(TxtSerie.Text)
DtaEntrada.Refresh
DtaEntrada.Recordset.Edit
DtaEntrada.Recordset.Fentrada = TxtSerie.Text
DtaEntrada.Recordset.Update
MsgBox "Programa Licenciado para 15 dias"
Unload Me
End Sub

Private Sub CmdLicencia_Click()
DtaEntrada.Refresh
DtaEntrada.Recordset.MoveNext
DtaEntrada.Recordset.Edit
DtaEntrada.Recordset.Fentrada = TxtSerie.Text
DtaEntrada.Recordset.Update

MsgBox "Programa Correctamente Licenciado"
Unload Me
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()

Directorio = App.Path + "\nominas.log"
Conexion = ";DATABASENAME=" + Directorio + ";UID=Administrador;PWD=15081977"
Me.DtaEntrada.DatabaseName = Directorio
Me.DtaEntrada.ConnectionString = Conexion

End Sub
