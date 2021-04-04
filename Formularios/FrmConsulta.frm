VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Zeus Nóminas.   Consultando........."
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   HelpContextID   =   38
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Restaura 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MaskColor       =   &H80000007&
      TabIndex        =   3
      ToolTipText     =   "Haga Click Para Restaurar Consulta"
      Top             =   120
      Width           =   255
   End
   Begin MSDBGrid.DBGrid DBGridConsulta 
      Bindings        =   "FrmConsulta.frx":0000
      Height          =   1095
      Left            =   120
      OleObjectBlob   =   "FrmConsulta.frx":001A
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton CmdSalir 
      DownPicture     =   "FrmConsulta.frx":09F4
      Height          =   375
      Left            =   2280
      MouseIcon       =   "FrmConsulta.frx":24D6
      MousePointer    =   99  'Custom
      Picture         =   "FrmConsulta.frx":2918
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton CmdPegar 
      DownPicture     =   "FrmConsulta.frx":43FA
      Height          =   375
      Left            =   840
      MouseIcon       =   "FrmConsulta.frx":5EDC
      MousePointer    =   99  'Custom
      Picture         =   "FrmConsulta.frx":631E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Data DtaConsulta 
      Caption         =   "DtaConsulta"
      Connect         =   "Access"
      DatabaseName    =   "C:\Zeus Nominas\Nominas.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CambioMoneda"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "FrmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdPegar_Click()
 On Error GoTo TipoErrs
 FrmCambioMoneda.MaskEdFecha.Text = DtaConsulta.Recordset.fechadia
 FrmCambioMoneda.MaskEdMonto.Text = Format((DtaConsulta.Recordset.montodia), "##,###0.00")
 'FrmCambioMoneda.CmdGrabar.SetFocus
 Unload Me
Exit Sub
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Private Sub CmdSalir_Click()
 Unload Me
End Sub

Private Sub DBGridConsulta_DblClick()
On Error GoTo TipoErrs
Dim SqlConsulta As Variant, Var As Variant
 FrmCambioMoneda.MaskEdFecha.Text = DtaConsulta.Recordset.fechadia
 'FrmCambioMoneda.MaskEdMonto.Text = Format((DtaConsulta.Recordset.MontoDia), "##,###0.00")

 Unload Me
Exit Sub
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Private Sub DBGridConsulta_KeyPress(KeyAscii As Integer)
 On Error GoTo TipoErrs
 Dim Consulta
  Lectura = KeyAscii
  LeeTecla
 If KeyAscii = 13 Then
  Dim SqlConsulta As Variant, Var As Variant
  FrmCambioMoneda.MaskEdFecha.Text = DtaConsulta.Recordset.fechadia
  'FrmCambioMoneda.MaskEdMonto.Text = Format((DtaConsulta.Recordset.MontoDia), "##,###0.00")
  Unload Me
 Else
   If Respuesta = "" Or Lectura = "*" Then
      Respuesta = Lectura
      DtaConsulta.RecordSource = "SELECT CambioMoneda.FechaDia,CambioMoneda.MontoDia From CambioMoneda  Where (((CambioMoneda.FechaDia) Like '" & Respuesta & "*'))"
      DtaConsulta.Refresh
     If Lectura = "*" Then
      Respuesta = ""
      Lectura = ""
     End If
   Else
    Respuesta = "" & Respuesta & Lectura & ""
    DtaConsulta.RecordSource = "SELECT CambioMoneda.FechaDia,CambioMoneda.MontoDia From CambioMoneda  Where (((CambioMoneda.FechaDia) Like '" & Respuesta & "*'))"
    DtaConsulta.Refresh
   End If
   
 End If
Exit Sub
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Private Sub Form_Activate()
FrmConsulta.DBGridConsulta.SetFocus
FrmConsulta.DBGridConsulta.Columns(1).Caption = "Monto del Dia"
FrmConsulta.DBGridConsulta.Columns(0).Caption = "Fecha de Cambio"
End Sub

Private Sub Form_Load()

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Respuesta = ""
 Lectura = ""
End Sub

Private Sub Restaura_Click()
Unload Me
FrmConsulta.Show
End Sub
