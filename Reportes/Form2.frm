VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "Form2"
   ScaleHeight     =   5415
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin MSCommLib.MSComm MSComm1 
      Left            =   360
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox TxtLista 
      Height          =   3135
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   7455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Me.MSComm1.CommPort = 1
Me.MSComm1.Settings = "9600,N,8,1"
Me.MSComm1.InputMode = comInputModeText
Me.MSComm1.Handshaking = comRTSXOnXOff
Me.MSComm1.SThreshold = 1
Me.MSComm1.RThreshold = 1
Me.MSComm1.PortOpen = True

End Sub

Private Sub mscReloj1_OnComm()

 
End Sub

Private Sub MSComm1_OnComm()
Dim Peso As Variant
Select Case MSComm1.CommEvent

Case 2

Me.TxtLista.Text = "Humberto Pesa" & Me.MSComm1.Input
DoEvents
End Select
End Sub
