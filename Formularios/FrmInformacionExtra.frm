VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmInformacionExtra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informacio Empledo"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5160
      TabIndex        =   17
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6735
      Begin VB.PictureBox Picture1 
         Height          =   2655
         Left            =   3840
         ScaleHeight     =   2595
         ScaleWidth      =   2595
         TabIndex        =   16
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   15
         Top             =   2520
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmInformacionExtra.frx":0000
         Left            =   1680
         List            =   "FrmInformacionExtra.frx":000A
         TabIndex        =   9
         Text            =   "SOLTERO (A)"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   5
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox TxtNombre2 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo DBCodigoEmpleado 
         Bindings        =   "FrmInformacionExtra.frx":0027
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "CodEmpleado1"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "FrmInformacionExtra.frx":0042
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "CodEmpleado1"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "FrmInformacionExtra.frx":005D
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "CodEmpleado1"
         Text            =   ""
      End
      Begin VB.Label Label6 
         Caption         =   "Monto Incentivo:"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Turno:"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Jefe Inmediato:"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Estado Civil:"
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Profesion:"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Celular Emergencia;"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label54 
         Caption         =   "Numero Celular:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   1695
   End
End
Attribute VB_Name = "FrmInformacionExtra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
End Sub

