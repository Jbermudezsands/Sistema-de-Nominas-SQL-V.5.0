VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{EAD61168-CF37-11D1-A050-70D904C10000}#2.0#0"; "MacWin.ocx"
Begin VB.Form FrmSuspencion 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Suspenci�n"
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5205
   ForeColor       =   &H00EFEFEF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc DtaSubsidios 
      Height          =   375
      Left            =   1200
      Top             =   5760
      Width           =   3015
      _ExtentX        =   5318
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
      Caption         =   "DtaSubsidios"
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
   Begin MSAdodcLib.Adodc DtaEmpleados 
      Height          =   330
      Left            =   1200
      Top             =   5040
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
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
      Caption         =   "DtaEmpleados"
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
   Begin MacWindow.MacWin MacWin1 
      Height          =   300
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   529
      Caption         =   "  Suspenciones   "
   End
   Begin VB.CommandButton CmdCancelar 
      DownPicture     =   "FrmSuspencion.frx":0000
      Height          =   375
      Left            =   3360
      Picture         =   "FrmSuspencion.frx":1AE2
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton CmdAceptar 
      DownPicture     =   "FrmSuspencion.frx":35C4
      Height          =   375
      Left            =   1920
      Picture         =   "FrmSuspencion.frx":50A6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox TxtMotivo 
      Height          =   1095
      Left            =   840
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3000
      Width           =   4095
   End
   Begin MSComCtl2.DTPicker DTPFechaFin 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   17104897
      CurrentDate     =   37139
   End
   Begin MSComCtl2.DTPicker DTPFechaIni 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   17104897
      CurrentDate     =   37139
   End
   Begin VB.TextBox TxtNombre 
      Height          =   375
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox TxtCodEmpleado 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Motivo"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Finalizaci�n"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Inicial"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "FrmSuspencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
Dim NumSubsidio As Integer
Me.DtaSubsidios.Refresh
If DtaSubsidios.Recordset.EOF Then
 NumSubsidio = 1
Else
 Me.DtaSubsidios.Recordset.MoveLast
 NumSubsidio = DtaSubsidios.Recordset("ID") + 1
End If
DtaSubsidios.Recordset.AddNew
 dDtaSubsidios.Recordset("CodEmpleado") = frmEmpleado.TxtCodEmpleado.Text
DtaSubsidios.Recordset("Fechaini") = DTPFechaIni.Value
DtaSubsidios.Recordset("Fechafin") = dtpFechaFin.Value
DtaSubsidios.Recordset("ID") = NumSubsidio
If TxtMotivo.Text <> "" Then
    DtaSubsidios.Recordset("motivo") = TxtMotivo.Text
Else
    DtaSubsidios.Recordset("motivo") = "Sin Motivo"
End If
DtaSubsidios.Recordset("activo") = True
DtaSubsidios.Recordset.Update
'ubico al empleado y le coloco la suspencion
DtaEmpleados.Refresh
Do While Not DtaEmpleados.Recordset.EOF
   If DtaEmpleados.Recordset("CodEmpleado") = frmEmpleado.TxtCodEmpleado.Text Then
      'DtaEmpleados.Recordset.Edit
      DtaEmpleados.Recordset("ausente") = True
      DtaEmpleados.Recordset.Update
   Exit Do
   End If
DtaEmpleados.Recordset.MoveNext
Loop

Unload Me
End Sub

Private Sub CmdCancelar_Click()
    frmEmpleado.ChkSuspendido.Value = 0
    frmEmpleado.LblSuspendido.Visible = False
Unload Me
End Sub

Private Sub Form_Load()
With Me.DtaEmpleados
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Empleado"
   .Refresh
End With

With Me.DtaSubsidios
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Subsidios"
   .Refresh
End With
    
DTPFechaIni.Value = Now
dtpFechaFin.Value = Now
End Sub
