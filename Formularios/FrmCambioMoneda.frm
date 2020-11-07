VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmCambioMoneda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla Cambio Dólar"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   HelpContextID   =   37
   Icon            =   "FrmCambioMoneda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   7290
   Begin VB.CommandButton CmdUltimo 
      Caption         =   "Ultimo"
      Height          =   375
      Left            =   4080
      MouseIcon       =   "FrmCambioMoneda.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CmdPrimero 
      Caption         =   "Primero"
      Height          =   375
      Index           =   1
      Left            =   3120
      MouseIcon       =   "FrmCambioMoneda.frx":074C
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CmdConsulta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   2880
      MouseIcon       =   "FrmCambioMoneda.frx":0B8E
      Picture         =   "FrmCambioMoneda.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Consulta"
      Top             =   120
      Width           =   255
   End
   Begin MSMask.MaskEdBox MaskEdFecha 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdMonto 
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Data DtaCambio 
      Caption         =   "DtaCambio"
      Connect         =   "Access"
      DatabaseName    =   "C:\Zeus Nominas\Nominas.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CambioMoneda"
      Top             =   -120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6240
      MouseIcon       =   "FrmCambioMoneda.frx":1412
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   120
      MouseIcon       =   "FrmCambioMoneda.frx":1854
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdanterior 
      Caption         =   "Anterior"
      Height          =   375
      Index           =   0
      Left            =   1200
      MouseIcon       =   "FrmCambioMoneda.frx":1C96
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   2160
      MouseIcon       =   "FrmCambioMoneda.frx":20D8
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   5160
      MouseIcon       =   "FrmCambioMoneda.frx":251A
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Monto del Día:"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha de Cambio:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FrmCambioMoneda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnterior_Click(Index As Integer)
'On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Cambio Dolar")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaCambio.Recordset.MovePrevious

If DtaCambio.Recordset.BOF Then
 DtaCambio.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Zeus Nóminas"
Else
 MaskEdFecha.Text = Format(DtaCambio.Recordset.Fechadia, "dd/mm/yyyy")
 MaskEdMonto.Text = DtaCambio.Recordset.Montodia
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdBorrar_Click()
'On Error GoTo TipoErrs
 Dim Respuesta, Rsp
'Elimino el registro activo en la pantalla
  Set Rsp = DtaCambio.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando el Cambio de:" & MaskEdFecha.Text)
   If Respuesta = 6 Then
     Rsp.Delete
      Rsp.MovePrevious
      MaskEdFecha.Text = "__/__/____"
      DtaCambio.Recordset.MoveLast
      DtaCambio.Recordset.MovePrevious
      Salida = False
   End If
Exit Sub
TipoErrs:
    ControlErrores
    Unload Me
End Sub

Private Sub CmdConsulta_Click(Index As Integer)
On Error GoTo TipoErrs
frmTasa2.Show 1
Exit Sub
TipoErrs:
ControlErrores
Unload Me
End Sub

Private Sub CmdGrabar_Click()
'On Error GoTo TipoErrs
Evaluar = True
If MaskEdMonto.Text = "" Or MaskEdFecha = "__/__/____" Or MaskEdMonto.Text = "0.00" Then
 MsgBox "No se Pueden Dejar Campos en Blanco", vbCritical, "Error:Zeus Nominas"
 Exit Sub
End If
Salida = False

 'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
      MaskEdMonto.Text = Format((MaskEdMonto.Text), " ##,##0.00")
      DtaCambio.Refresh
      Do While Not DtaCambio.Recordset.EOF
       If DtaCambio.Recordset.Fechadia = MaskEdFecha.Text Then
          DtaCambio.Recordset.Edit
          DtaCambio.Recordset.Fields("MontoDia") = MaskEdMonto.Text
          DtaCambio.Recordset.Update
          MaskEdFecha.Text = "__/__/____"
          DtaCambio.Recordset.MoveLast
          DtaCambio.Recordset.MovePrevious
          Salida = False
          Exit Sub
       End If
      DtaCambio.Recordset.MoveNext
      Loop
  'Si despues de Buscar no exite el codigo grabo todos los cambios
         DtaCambio.Recordset.AddNew
         DtaCambio.Recordset.Fields("MontoDia") = MaskEdMonto.Text
         DtaCambio.Recordset.Fields("FechaDia") = MaskEdFecha.Text
         DtaCambio.Recordset.Update
         MaskEdFecha.Text = "__/__/____"
         DtaCambio.Recordset.MoveLast
         DtaCambio.Recordset.MovePrevious
         Salida = False
         Exit Sub

TipoErrs:
  ControlErrores
  If Error = 1 Then
    Exit Sub
  Else
  Unload Me
 End If
 End Sub

Private Sub CmdPrimero_Click(Index As Integer)
'On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Cambio Dolar")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaCambio.Recordset.MoveFirst

If DtaCambio.Recordset.BOF Then
 DtaCambio.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Zeus Nóminas"
Else
 MaskEdFecha.Text = Format(DtaCambio.Recordset.Fechadia, "dd/mm/yyyy")
 MaskEdMonto.Text = DtaCambio.Recordset.Montodia
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdSalir_Click()
 Unload Me
End Sub

Private Sub DBCFecha_Change()
FrmConsulta.Show
End Sub

Private Sub CmdSiguiente_Click()
' On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Cambio Dolar")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaCambio.Recordset.MoveNext

If DtaCambio.Recordset.EOF Then
 DtaCambio.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Zeus Nóminas"
Else
 MaskEdFecha.Text = Format(DtaCambio.Recordset.Fechadia, "dd/mm/yyyy")
 MaskEdMonto.Text = DtaCambio.Recordset.Montodia
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdUltimo_Click()
'  On Error GoTo TipoErrs
ValidaSalida ("en la Tabla Cambio Dolar")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaCambio.Recordset.MoveLast

If DtaCambio.Recordset.EOF Then
 DtaCambio.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Zeus Nóminas"
Else
 MaskEdFecha.Text = Format(DtaCambio.Recordset.Fechadia, "dd/mm/yyyy")
 MaskEdMonto.Text = DtaCambio.Recordset.Montodia
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub



Private Sub Form_Activate()
 If Not BRegistroMoneda = True Then
   CmdBorrar.Enabled = False
 End If
 If Not GRegistroMoneda = True Then
   CmdGrabar.Enabled = False
 End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If GRegistroMoneda = True Then
ValidaSalida ("en la Tabla Cambio Dolar")
 If Contesta Then
  CmdGrabar.Value = True
  Salida = False
  Unload Me
 Else
   Salida = False
   Unload Me
 End If
End If

End Sub

Private Sub MaskEdFecha_Change()
On Error GoTo TipoErrs
Evaluar = True
'Al ejecutar algun cambio en el combo actualizo el nombre del departamento
   DtaCambio.Refresh
   Do While Not DtaCambio.Recordset.EOF
     If DtaCambio.Recordset.Fechadia = MaskEdFecha.Text Then
        MaskEdMonto.Text = Format((DtaCambio.Recordset.Montodia), " ##,###0.00")
        Salida = False
        Exit Do
     Else
        Salida = True
        MaskEdMonto.Text = "0.00"
     End If
       DtaCambio.Recordset.MoveNext
   Loop
Exit Sub
TipoErrs:
ControlErrores
Unload Me
End Sub

Private Sub MaskEdFecha_KeyPress(KeyAscii As Integer)
 If KeyAscii = 42 Then
  FrmConsulta.Show
 ElseIf KeyAscii = 13 Then
   MaskEdMonto.SetFocus
 End If
End Sub


Private Sub MaskEdMonto_Change()
Salida = True
End Sub

Private Sub MaskEdMonto_GotFocus()
MaskEdMonto.Text = Format((DtaCambio.Recordset.Montodia), " ##,###0.00")
End Sub

Private Sub MaskEdMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 CmdGrabar.SetFocus
Else
   Evaluar = False
  End If
End Sub

Private Sub MaskEdMonto_LostFocus()
MaskEdMonto.Text = Format((MaskEdMonto.Text), "##,##0.00")
End Sub
