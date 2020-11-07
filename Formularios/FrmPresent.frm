VERSION 5.00
Begin VB.Form FrmPresent 
   BackColor       =   &H80000007&
   ClientHeight    =   6690
   ClientLeft      =   2070
   ClientTop       =   1245
   ClientWidth     =   9105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   6615
      Left            =   0
      Picture         =   "FrmPresent.frx":0000
      ScaleHeight     =   6555
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   0
      Width           =   9135
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   480
      Top             =   3240
   End
End
Attribute VB_Name = "FrmPresent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
Unload Me
FrmListaUsuario.Show
End Sub
