VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFechaIngresoBaja 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DtpFechaIngreso 
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      Format          =   61865985
      CurrentDate     =   38802
   End
   Begin VB.Label Label9 
      Caption         =   "Fecha Ingreso"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "FrmFechaIngresoBaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
FechaIngreso = Format(Me.DtpFechaIngreso.Value, "dd/mm/yyyy")
Unload Me
End Sub

Private Sub CmdCancelar_Click()
FechaIngreso = Format(Me.DtpFechaIngreso.Value, "dd/mm/yyyy")
Unload Me
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(219, 226, 242)
Me.DtpFechaIngreso.Value = Now
MDIPrimero.Skin1.ApplySkin hWnd
End Sub
