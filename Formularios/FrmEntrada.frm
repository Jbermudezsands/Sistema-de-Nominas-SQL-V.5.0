VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Begin VB.Form FrmEntrada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control  Zeus Nóminas"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   HelpContextID   =   1
   Icon            =   "FrmEntrada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   3975
      TabIndex        =   9
      Top             =   0
      Width           =   3975
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Validacion de Usuarios"
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
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   2400
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   6720
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image2 
         Height          =   1020
         Left            =   0
         Picture         =   "FrmEntrada.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1290
      End
   End
   Begin MSAdodcLib.Adodc DtaUsuario 
      Height          =   375
      Left            =   120
      Top             =   3840
      Width           =   3135
      _ExtentX        =   5530
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
      Caption         =   "Adodc1"
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
      Left            =   120
      Top             =   3960
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.TextBox TxtCodEmpleado 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Text            =   "TxtCodEmpleado"
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox TxtNombreUsuario 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      DownPicture     =   "FrmEntrada.frx":6A94
      Height          =   735
      Left            =   360
      MouseIcon       =   "FrmEntrada.frx":8576
      MousePointer    =   99  'Custom
      Picture         =   "FrmEntrada.frx":89B8
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      DownPicture     =   "FrmEntrada.frx":92EA
      Height          =   735
      Left            =   2520
      MouseIcon       =   "FrmEntrada.frx":ADCC
      MousePointer    =   99  'Custom
      Picture         =   "FrmEntrada.frx":B20E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox TxtNivel 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      TabIndex        =   2
      Text            =   "TxtNivel"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox TxtClave 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Nivel del Usuario:"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label LblClave 
      Caption         =   "Clave de Acceso:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre del Usuario:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "FrmEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
 DtaUsuario.Refresh
       Do While Not DtaUsuario.Recordset.EOF
         If DtaUsuario.Recordset("NombreUsuario") = TxtNombreUsuario.Text Then
            If DtaUsuario.Recordset("Clave") = TxtClave.Text Then
               NivelAcceso = DtaUsuario.Recordset("NivelAcceso")
               CodPasword = DtaUsuario.Recordset("CodUsuario")
               CodigoUsuario = DtaUsuario.Recordset("CodUsuario")
               NombreUsuario = TxtNombreUsuario.Text
                     Unload Me
                     FrmListaUsuario.CmdSalir.Value = True
                     MDIPrimero.Show
                   
                    Exit Sub
                
                   Guia = 1
              Else
                Guia = 1
            End If 'Cierre del If Pasword
         Else
          Guia = 1
         End If 'Cierre del If NombreEmpleado
       DtaUsuario.Recordset.MoveNext
       Loop
    Select Case Guia
       Case 1: MsgBox "No Tiene Permiso", vbCritical, "Sistema de Nominas"
               TxtClave.Text = ""
               TxtClave.SetFocus
    End Select
End Sub

Private Sub CmdCancelar_Click()
 Unload Me
End Sub

Private Sub DBNombreUsuario_Change()
 'Al ejecutar algun cambio en el combo actualizo el nombre del Empleado
   DtaUsuario.Refresh
   Do While Not DtaUsuario.Recordset.EOF
     If DtaUsuario.Recordset("NombreUsuario") = DBNombreUsuario.Text Then
        'VAL(TxtNivel.Text) = DtaUsuario.Recordset("NivelAcceso")
        'TxtClave.SetFocus
        Exit Do
     End If
       DtaUsuario.Recordset.MoveNext
   Loop
End Sub

Private Sub DBNombreUsuario_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   TxtNivel.SetFocus
 End If
End Sub

Private Sub Form_Load()

With Me.DtaUsuario
   .ConnectionString = Conexion
   .RecordSource = "Usuarios"
   .Refresh
End With

End Sub


Private Sub TxtClave_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   CmdAceptar.Value = True
 End If
End Sub
Private Sub TxtNivel_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
   TxtClave.SetFocus
  End If
End Sub

