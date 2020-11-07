VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Begin VB.Form FrmEditarPrestamo 
   Caption         =   "Editar Prestamo"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc DtaMovPrestamo 
      Height          =   375
      Left            =   480
      Top             =   5400
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "DtaMovPrestamo"
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
   Begin MSAdodcLib.Adodc DtaPrestamo 
      Height          =   495
      Left            =   600
      Top             =   6360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
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
      Caption         =   "DtaPrestamo"
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
      Left            =   480
      Top             =   3600
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Revalorando Cuota"
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   4695
      Begin VB.TextBox TxtNumPrestamo 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox TxtNumCuota 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox TxtMontoOld 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Número de Préstamo"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Numero de Cuota"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Monto Actual"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdCancelar 
      DownPicture     =   "FrmEditarPrestamo.frx":0000
      Height          =   375
      Left            =   3360
      Picture         =   "FrmEditarPrestamo.frx":1AE2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton CmdAceptar 
      DownPicture     =   "FrmEditarPrestamo.frx":35C4
      Height          =   375
      Left            =   1920
      Picture         =   "FrmEditarPrestamo.frx":50A6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox TxtMontoNew 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label LblNombre 
      Alignment       =   2  'Center
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "Monto a Pagar"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "FrmEditarPrestamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
On Error GoTo TipoErr
Dim CantCuotas As Integer
Dim Saldo As Integer
Dim MontoOld As Double
Dim MontoNew As Double
Dim CuotaIgual As Double
Dim SaldoPrestamo As Double

If Not IsNumeric(TxtMontoNew.Text) Then
   MsgBox "El Monto Digitado es erróneo"
   TxtMontoNew.SetFocus
   Exit Sub
End If

Dtaprestamo.Refresh
Do While Not Dtaprestamo.Recordset.EOF
  If Dtaprestamo.Recordset("NumPrestamo") = val(TxtNumPrestamo.Text) Then
     SaldoPrestamo = Dtaprestamo.Recordset("Saldo")
  End If
Dtaprestamo.Recordset.MoveNext
Loop

Saldo = 0
DtaMovprestamo.Refresh
DtaMovprestamo.Recordset.MoveLast
CantCuotas = DtaMovprestamo.Recordset.RecordCount

DtaMovprestamo.Refresh
Do While Not DtaMovprestamo.Recordset.EOF
Saldo = Saldo + DtaMovprestamo.Recordset("CuotaIgual")
DtaMovprestamo.Recordset.MoveNext
Loop


MontoOld = val(TxtMontoOld.Text)
MontoNew = val(TxtMontoNew.Text)
If Saldo < MontoNew Then
   MsgBox "El Nuevo Monto no puede ser mayor que el saldo del Préstamo"
   Exit Sub
End If


CuotaIgual = Format((Saldo - MontoNew) / (CantCuotas - 1), "###,##0.00")

DtaMovprestamo.Refresh
Do While Not DtaMovprestamo.Recordset.EOF
   If DtaMovprestamo.Recordset("numcuota") = val(txtNumCuota.Text) Then
   'DtaMovPrestamo.Recordset.Edit
   DtaMovprestamo.Recordset("CuotaIgual") = MontoNew
   DtaMovprestamo.Recordset("saldocuota") = SaldoPrestamo - MontoNew
   SaldoPrestamo = SaldoPrestamo - MontoNew
   DtaMovprestamo.Recordset.Update
   Else
   'DtaMovPrestamo.Recordset.Edit
   DtaMovprestamo.Recordset("CuotaIgual") = CuotaIgual
   DtaMovprestamo.Recordset("saldocuota") = SaldoPrestamo - CuotaIgual
   SaldoPrestamo = SaldoPrestamo - CuotaIgual
   DtaMovprestamo.Recordset.Update
   End If

DtaMovprestamo.Recordset.MoveNext
Loop

Unload Me
frmEmpleado.DtaMovprestamo.Refresh
frmEmpleado.DbgrLibreta.Columns(0).Visible = False
frmEmpleado.DbgrLibreta.Columns(1).Visible = False
frmEmpleado.DbgrLibreta.Columns(7).Visible = False

Exit Sub
TipoErr:
    ControlErrores
Unload Me


End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Activate()
On Error GoTo TipoErr
Dim SQLMovPrestamo As String
Dim NumPrestamo As Long

NumPrestamo = val(TxtNumPrestamo.Text)
SQLMovPrestamo = "SELECT MovPrestamo.NumPrestamo, MovPrestamo.NumCuota, MovPrestamo.Monto, MovPrestamo.Interes, MovPrestamo.CuotaIgual, MovPrestamo.SaldoCuota, MovPrestamo.Cancelado, MovPrestamo.NumNomina From MovPrestamo WHERE MovPrestamo.Cancelado=0 AND MovPrestamo.NumPrestamo=" & NumPrestamo & ""
DtaMovprestamo.RecordSource = SQLMovPrestamo
DtaMovprestamo.Refresh

Exit Sub
TipoErr:
    ControlErrores
Unload Me


End Sub

Private Sub Form_Load()
On Error GoTo TipoErr
With Me.DtaMovprestamo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion


End With

With Me.Dtaprestamo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Prestamo"
   .Refresh
End With

Exit Sub
TipoErr:
    ControlErrores
Unload Me
End Sub
