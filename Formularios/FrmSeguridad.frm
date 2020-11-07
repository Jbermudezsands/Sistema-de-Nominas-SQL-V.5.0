VERSION 5.00
Begin VB.Form FrmSeguridad 
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5640
   ControlBox      =   0   'False
   Icon            =   "FrmSeguridad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdClaves 
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   2880
      TabIndex        =   5
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton CmdClave 
         Caption         =   "Encriptar"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cmdesclave 
         Caption         =   "DesEncriptar"
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton CmdSerie 
         Caption         =   "Número de Serie"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.TextBox TxtClave 
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtDesclave 
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Clave"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Desclave"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub CmdClave_Click()
TxtDesclave.Text = Encrypt(TxtClave.Text)
End Sub

Private Sub CmdClaves_Click()
clave = InputBox("Digite su clave")
If clave = "1521080619771978" Then

   FrmLicencia.Show
   FrmLicencia.TxtSerie.Text = TxtDesclave.Text
  Unload Me
Else
Exit Sub
End If

End Sub

Private Sub Cmdesclave_Click()
TxtClave.Text = Decrypt(TxtDesclave.Text)
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSerie_Click()
Dim Vol As String * 256, FileSystem As String * 256, unidad As String
Dim longitud As Long, NumSerie As Long, Flags As Long

unidad = App.Path + "\"
unidad = Mid(unidad, 1, 3)
Call GetVolumeInformation(unidad, Vol, 256, NumSerie, longitud, Flags, FileSystem, 256)
TxtClave.Text = NumSerie
End Sub

Private Sub Form_Load()
Me.Top = 2200
Me.Left = 2200
End Sub

