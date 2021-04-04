VERSION 5.00
Begin VB.Form FrmDirectorio 
   Caption         =   "Directorio de Fotos"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      DownPicture     =   "FrmDirectorio.frx":0000
      Height          =   735
      Left            =   1320
      MouseIcon       =   "FrmDirectorio.frx":1AE2
      MousePointer    =   99  'Custom
      Picture         =   "FrmDirectorio.frx":1F24
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      DownPicture     =   "FrmDirectorio.frx":2B66
      Height          =   735
      Left            =   120
      MouseIcon       =   "FrmDirectorio.frx":4648
      MousePointer    =   99  'Custom
      Picture         =   "FrmDirectorio.frx":4A8A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.DirListBox Dir1 
      DragIcon        =   "FrmDirectorio.frx":53BC
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   2400
      TabIndex        =   3
      Top             =   1320
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      DragIcon        =   "FrmDirectorio.frx":56C6
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   5655
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Image Image2 
         Height          =   1020
         Left            =   0
         Picture         =   "FrmDirectorio.frx":59D0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1290
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   6720
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "RUTA DE FOTOS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   2400
      End
   End
End
Attribute VB_Name = "FrmDirectorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
 FrmControles.TxtRutaFoto.Text = Me.Dir1.Path
 Unload Me
End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub Drive1_Change()
  On Error GoTo DriveErrs
    Dir1.Path = Drive1.Drive
Exit Sub
DriveErrs:
   ControlErrores
End Sub

Private Sub Form_Load()
 Dir1.Path = App.Path
End Sub
