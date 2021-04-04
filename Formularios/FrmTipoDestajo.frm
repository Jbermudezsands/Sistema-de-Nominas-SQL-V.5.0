VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form FrmTipoDestajo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Destajo"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   HelpContextID   =   25
   Icon            =   "FrmTipoDestajo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   538
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   3135
      Begin XtremeSuiteControls.PushButton CmdAnterior 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Anterior"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmTipoDestajo.frx":030A
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdSiguiente 
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Siguiente"
         ForeColor       =   0
         TextAlignment   =   0
         Appearance      =   6
         Picture         =   "FrmTipoDestajo.frx":080C
         ImageAlignment  =   1
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton CmdPrimero 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Primero"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmTipoDestajo.frx":0D10
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdUltimo 
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   720
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Siguiente"
         ForeColor       =   0
         TextAlignment   =   0
         Appearance      =   6
         Picture         =   "FrmTipoDestajo.frx":1212
         ImageAlignment  =   1
         TextImageRelation=   4
      End
   End
   Begin VB.TextBox TxtDebito 
      Height          =   285
      Left            =   3720
      MaxLength       =   20
      TabIndex        =   6
      Text            =   "11111"
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox TxtDestajo 
      Height          =   285
      Left            =   3720
      MaxLength       =   35
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo DBCDestajo 
      Bindings        =   "FrmTipoDestajo.frx":1714
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "COdTipoDestajo"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaTipoDestajo 
      Height          =   375
      Left            =   600
      Top             =   3240
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
      Caption         =   "DtaTipoDestajo"
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
   Begin MSMask.MaskEdBox MaskEdMonto 
      Height          =   285
      Left            =   5880
      TabIndex        =   2
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      AllowPrompt     =   -1  'True
      PromptChar      =   "_"
   End
   Begin XtremeSuiteControls.PushButton CmdGrabar 
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   1320
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Grabar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmTipoDestajo.frx":1731
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdBorrar 
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   1800
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Borrar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmTipoDestajo.frx":3A95
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   1800
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmTipoDestajo.frx":3F49
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton Command1 
      Height          =   390
      Left            =   5880
      TabIndex        =   17
      Top             =   720
      Width           =   390
      _Version        =   786432
      _ExtentX        =   688
      _ExtentY        =   688
      _StockProps     =   79
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmTipoDestajo.frx":444D
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdRepetir 
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   1320
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Repetir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmTipoDestajo.frx":494F
      ImageAlignment  =   0
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta Contable:"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto:"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Destajo:"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Destajo:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FrmTipoDestajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Cargo")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaTipoDestajo.Recordset.MovePrevious

If DtaTipoDestajo.Recordset.BOF Then
 DtaTipoDestajo.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 TxtDestajo.Text = DtaTipoDestajo.Recordset("destajo")
 DBCDestajo.Text = DtaTipoDestajo.Recordset("Codtipodestajo")
 'MaskEdMonto.Text = DtaCargo.Recordset.Monto
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub cmdBorrar_Click()
 On Error GoTo TipoErrs
 Dim Respuesta, Rsp
'Elimino el registro activo en la pantalla
  Set Rsp = DtaTipoDestajo.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando el Destajo: " & TxtDestajo.Text)
   If Respuesta = 6 Then
     Rsp.Delete
      DBCDestajo.Text = ""
      TxtDestajo.Text = ""
      MaskEdMonto.Text = ""
      DtaTipoDestajo.Recordset.MoveLast
      DtaTipoDestajo.Recordset.MovePrevious
      Salida = False
   End If
 Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo TipoErrs
  Salida = False
  If TxtDestajo.Text = "" Or MaskEdMonto = "" Then
   MsgBox "Los Campos no Pueden quedar Vacios", vbCritical, "Error:Sistema de Nominas"
  End If
  
  
  'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
      DtaTipoDestajo.Refresh
      Do While Not DtaTipoDestajo.Recordset.EOF
       If DtaTipoDestajo.Recordset("Codtipodestajo") = DBCDestajo.Text Then
         'DtaTipoDestajo.Recordset.Edit
         DtaTipoDestajo.Recordset.Fields("Destajo") = TxtDestajo.Text
         DtaTipoDestajo.Recordset.Fields("Monto") = MaskEdMonto.Text
         
        If Me.TxtDebito.Text <> "" Then
         Me.DtaTipoDestajo.Recordset("CuentaContable") = Me.TxtDebito.Text
        End If
         DtaTipoDestajo.Recordset.Update
         DtaTipoDestajo.Recordset.MoveLast
         DtaTipoDestajo.Recordset.MovePrevious
         Salida = False
         Exit Sub
             
      End If
      DtaTipoDestajo.Recordset.MoveNext
      Loop
  'Si despues de Buscar no exite el codigo grabo todos los cambios
         DtaTipoDestajo.Recordset.AddNew
         DtaTipoDestajo.Recordset.Fields("CodTipoDestajo") = DBCDestajo.Text
         DtaTipoDestajo.Recordset.Fields("Destajo") = TxtDestajo.Text
         DtaTipoDestajo.Recordset.Fields("Monto") = MaskEdMonto.Text
        If Me.TxtDebito.Text <> "" Then
         Me.DtaTipoDestajo.Recordset("CuentaContable") = Me.TxtDebito.Text
        End If
         DtaTipoDestajo.Recordset.Update
         DtaTipoDestajo.Recordset.MoveLast
         DtaTipoDestajo.Recordset.MovePrevious
         Salida = False
         Exit Sub
         
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Private Sub CmdPirmero_Click()
 On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Cargo")
If Contesta Then
  CmdGrabar.Value = True
End If
DtaTipoDestajo.Recordset.MoveFirst
If DtaTipoDestajo.Recordset.BOF Then
 DtaTipoDestajo.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCDestajo.Text = DtaTipoDestajo.Recordset("Codtipodestajo")
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdRepetir_Click()
DtaTipoDestajo.Refresh
Do While Not DtaTipoDestajo.Recordset.EOF
 'DtaTipoDestajo.Recordset.Edit
  DtaTipoDestajo.Recordset("Monto") = val(MaskEdMonto.Text)
 DtaTipoDestajo.Recordset.Update

 DtaTipoDestajo.Recordset.MoveNext
Loop
 MsgBox "Proceso Terminado", vbInformation, "Sistema de Nominas"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSiguiente_Click()
On Error GoTo TipoErrs
 If GCargo = True Then
 ValidaSalida ("en la Tabla Cargo")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaTipoDestajo.Recordset.MoveNext
If DtaTipoDestajo.Recordset.EOF Then
  DtaTipoDestajo.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 TxtDestajo.Text = DtaTipoDestajo.Recordset("destajo")
 DBCDestajo.Text = DtaTipoDestajo.Recordset("Codtipodestajo")
End If
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdUltimo_Click()
 On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Cargo")
If Contesta Then
  CmdGrabar.Value = True
End If
   DtaTipoDestajo.Recordset.MoveLast
If DtaTipoDestajo.Recordset.EOF Then
  DtaTipoDestajo.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCDestajo.Text = DtaTipoDestajo.Recordset("Codtipodestajo")
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub DBCCargo_Change()

End Sub

Private Sub DBCCargo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  TxtCargo.SetFocus
 End If
End Sub

Private Sub Command1_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtDebito.Text = CuentaContable
End Sub

Private Sub DBCDestajo_Change()
On Error GoTo TipoErrs
Evaluar = True
'Al ejecutar algun cambio en el combo actualizo el nombre del departamento
   DtaTipoDestajo.Refresh
   Do While Not DtaTipoDestajo.Recordset.EOF
     If DtaTipoDestajo.Recordset("Codtipodestajo") = DBCDestajo.Text Then
        TxtDestajo.Text = DtaTipoDestajo.Recordset("destajo")
        MaskEdMonto.Text = Format((DtaTipoDestajo.Recordset("Monto")), "##,##0.00")
        If Not IsNull(Me.DtaTipoDestajo.Recordset("CuentaContable")) Then
         Me.TxtDebito.Text = Me.DtaTipoDestajo.Recordset("CuentaContable")
        End If
        Exit Do
     Else
        TxtDestajo.Text = ""
        MaskEdMonto.Text = "0.00"
     End If
       DtaTipoDestajo.Recordset.MoveNext
   Loop
Salida = False
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub Form_Activate()
'DBCDestajo.SetFocus
' If Not BCargo = True Then
'   CmdBorrar.Enabled = False
' End If
' If Not GCargo = True Then
'  CmdGrabar.Enabled = False
' End If
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(222, 227, 247)
Me.Frame1.BackColor = RGB(222, 227, 247)

With Me.DtaTipoDestajo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "TipoDestajo"
   .Refresh
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If GEmpleado = True Then
ValidaSalida ("en la Tabla Cargo")
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

Private Sub MaskEdMonto_Change()
Salida = True
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

Private Sub TxtCargo_Change()
Salida = True
End Sub

Private Sub TxtCargo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  MaskEdMonto.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub
