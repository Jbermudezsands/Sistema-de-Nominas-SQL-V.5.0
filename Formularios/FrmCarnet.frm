VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Begin VB.Form FrmCarnet 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3645
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5595
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5595
   Begin MSAdodcLib.Adodc DtaEmpleado 
      Height          =   375
      Left            =   360
      Top             =   4440
      Width           =   4095
      _ExtentX        =   7223
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
      Caption         =   "DtaEmpleado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   1080
      Top             =   3240
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.CommandButton CmdImprimir 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   5160
      Picture         =   "FrmCarnet.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton CmdSalir 
      BackColor       =   &H80000009&
      Caption         =   "X"
      Height          =   360
      Left            =   5280
      TabIndex        =   11
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   1695
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   3495
      Begin VB.Label LblCedula 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         Caption         =   "Cedula #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblCargo 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Nombres"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LblNombres 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Apellidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label LblApellidos 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Firma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   3240
      X2              =   5400
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label LblCodempleado2 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label LblCodempleado 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Ramses"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Carnet del Empleado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   120
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "FrmCarnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdImprimir_Click()
cmdSalir.Visible = False
cmdImprimir.Visible = False
FrmCarnet.PrintForm
Unload Me
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo TipoErr
Me.Top = 1500
Me.Left = 500

'DtaEmpleado '.DatabaseName = Ruta
DtaEmpleado.ConnectionString = Conexion

Dim SQlEmpleado As String
SQlEmpleado = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Cargo.Cargo FROM Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo WHERE Empleado.CodEmpleado='" & frmEmpleado.TxtCodEmpleado.Text & "'"
DtaEmpleado.RecordSource = SQlEmpleado
DtaEmpleado.Refresh

LblNombres.Caption = DtaEmpleado.Recordset("Nombre1") + " " + DtaEmpleado.Recordset("Nombre2")
LblApellidos.Caption = DtaEmpleado.Recordset("Apellido1") + " " + DtaEmpleado.Recordset("Apellido2")
LblCodEmpleado.Caption = "*" + DtaEmpleado.Recordset("CodEmpleado1") + "*"
LblCodempleado2.Caption = DtaEmpleado.Recordset("CodEmpleado1")
LblCargo.Caption = DtaEmpleado.Recordset("Cargo")
        'coloco la foto
        If Dir(RutaFoto & CodEmpleado & ".jpg") <> "" Then
           Destino = RutaFoto & CodEmpleado & ".jpg"
        ElseIf Dir(RutaFoto & CodEmpleado & ".gif") <> "" Then
           Destino = RutaFoto & CodEmpleado & ".gif"
        ElseIf Dir(RutaFoto & CodEmpleado & ".bmp") <> "" Then
           Destino = RutaFoto & CodEmpleado & ".bmp"
        End If
        
        If (Dir(Destino) <> "") Then
         Image1.Picture = LoadPicture(Destino)
        Else
         Destino = RutaFoto + "\Zw.bmp"
         Image1.Picture = LoadPicture(Destino)
        End If
        
Exit Sub
TipoErr:
ControlErrores
End Sub

Private Sub LblNombre1_Click()

End Sub

