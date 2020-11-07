VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FrmInforme 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informacion de Usuarios"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   387
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4680
      Picture         =   "FrmInforme.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Usuario Actual"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin ACTIVESKINLibCtl.SkinLabel LblUsuario 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "FrmInforme.frx":0442
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNivel 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "FrmInforme.frx":04BC
         TabIndex        =   5
         Top             =   960
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmInforme.frx":0536
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmInforme.frx":05B2
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CmdAceptar_Click()
Unload Me
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
LblNivel.Caption = NivelAcceso
LblUsuario.Caption = NombreUsuario
End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub
