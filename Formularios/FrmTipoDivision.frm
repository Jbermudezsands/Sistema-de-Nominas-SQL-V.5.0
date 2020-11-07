VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form FrmGrupo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Grupos de Nominas"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7920
   HelpContextID   =   25
   Icon            =   "FrmTipoDivision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   143
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   528
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   3135
      Begin XtremeSuiteControls.PushButton CmdAnterior 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Anterior"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmTipoDivision.frx":030A
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdSiguiente 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
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
         Picture         =   "FrmTipoDivision.frx":080C
         ImageAlignment  =   1
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton CmdPirmero 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Primero"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmTipoDivision.frx":0D10
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdUltimo 
         Height          =   375
         Left            =   1560
         TabIndex        =   9
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
         Picture         =   "FrmTipoDivision.frx":1212
         ImageAlignment  =   1
         TextImageRelation=   4
      End
   End
   Begin VB.TextBox TxtGrupo 
      Height          =   285
      Left            =   3480
      MaxLength       =   50
      TabIndex        =   3
      Top             =   240
      Width           =   4215
   End
   Begin MSDataListLib.DataCombo DBCGrupo 
      Bindings        =   "FrmTipoDivision.frx":1714
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodGrupo"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaGrupo 
      Height          =   375
      Left            =   720
      Top             =   3600
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "DtaGrupo"
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
   Begin XtremeSuiteControls.PushButton CmdGrabar 
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   840
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Grabar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmTipoDivision.frx":172B
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdBorrar 
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   1320
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Borrar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmTipoDivision.frx":3A8F
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmTipoDivision.frx":3F43
      ImageAlignment  =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo:"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Grupo:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FrmGrupo"
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
 DtaGrupo.Recordset.MovePrevious

If DtaGrupo.Recordset.BOF Then
 DtaGrupo.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado. Está al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 txtGrupo.Text = DtaGrupo.Recordset("grupo")
 DBCGrupo.Text = DtaGrupo.Recordset("Codgrupo")
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
  Set Rsp = DtaGrupo.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando el grupo: " & txtGrupo.Text)
   If Respuesta = 6 Then
     Rsp.Delete
      DBCGrupo.Text = ""
      txtGrupo.Text = ""
      
      DtaGrupo.Recordset.MoveLast
      DtaGrupo.Recordset.MovePrevious
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
  If txtGrupo.Text = "" Then
   MsgBox "El nombre del Grupo No puede quedar Vacío", vbCritical, "Error:Sistema de Nominas"
  End If
  
  
  'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
      DtaGrupo.Refresh
      Do While Not DtaGrupo.Recordset.EOF
       If DtaGrupo.Recordset("Codgrupo") = DBCGrupo.Text Then
         'DtaGrupo.Recordset.Edit
         DtaGrupo.Recordset.Fields("Grupo") = txtGrupo.Text
         DtaGrupo.Recordset.Update
         DtaGrupo.Recordset.MoveLast
         DtaGrupo.Recordset.MovePrevious
         Salida = False
         Exit Sub
             
      End If
      DtaGrupo.Recordset.MoveNext
      Loop
  'Si despues de Buscar no exite el codigo grabo todos los cambios
         DtaGrupo.Recordset.AddNew
         DtaGrupo.Recordset.Fields("CodGrupo") = DBCGrupo.Text
         DtaGrupo.Recordset.Fields("Grupo") = txtGrupo.Text
         DtaGrupo.Recordset.Update
         DtaGrupo.Recordset.MoveLast
         DtaGrupo.Recordset.MovePrevious
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
DtaGrupo.Recordset.MoveFirst
If DtaGrupo.Recordset.BOF Then
 DtaGrupo.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado. Está al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCGrupo.Text = DtaGrupo.Recordset("Codgrupo")
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdPrimero_Click()

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
 DtaGrupo.Recordset.MoveNext
If DtaGrupo.Recordset.EOF Then
  DtaGrupo.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado. Está al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 txtGrupo.Text = DtaGrupo.Recordset("grupo")
 DBCGrupo.Text = DtaGrupo.Recordset("Codgrupo")
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
   DtaGrupo.Recordset.MoveLast
If DtaGrupo.Recordset.EOF Then
  DtaGrupo.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado. Está al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCGrupo.Text = DtaGrupo.Recordset("Codgrupo")
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

Private Sub DBCDestajo_Change()
End Sub

Private Sub DBCDivision_Change()


End Sub

Private Sub DBCDivision_Click(Area As Integer)

End Sub

Private Sub DBCGrupo_Change()
On Error GoTo TipoErrs
Evaluar = True
'Al ejecutar algun cambio en el combo actualizo el nombre del departamento
   DtaGrupo.Refresh
   Do While Not DtaGrupo.Recordset.EOF
     If DtaGrupo.Recordset("Codgrupo") = DBCGrupo.Text Then
        txtGrupo.Text = DtaGrupo.Recordset("grupo")
        Exit Do
     Else
     txtGrupo.Text = ""
     End If
       DtaGrupo.Recordset.MoveNext
   Loop
Salida = False
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub Form_Activate()
'DBCGrupo.SetFocus
' If Not BCargo = True Then
'   cmdBorrar.Enabled = False
' End If
' If Not GCargo = True Then
'  CmdGrabar.Enabled = False
' End If
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(222, 227, 247)
Me.Frame1.BackColor = RGB(222, 227, 247)

With Me.DtaGrupo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Grupo"
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
