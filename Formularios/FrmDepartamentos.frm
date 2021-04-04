VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmDepartamentos 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Departamentos"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   ForeColor       =   &H00EFEFEF&
   HelpContextID   =   24
   Icon            =   "FrmDepartamentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   Begin XtremeSuiteControls.PushButton CmdGrabar 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   720
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Grabar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmDepartamentos.frx":030A
      ImageAlignment  =   0
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   3135
      Begin XtremeSuiteControls.PushButton CmdAnterior 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Anterior"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmDepartamentos.frx":266E
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdSiguiente 
         Height          =   375
         Left            =   1560
         TabIndex        =   9
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
         Picture         =   "FrmDepartamentos.frx":2B70
         ImageAlignment  =   1
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton CmdPrimero 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Primero"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmDepartamentos.frx":3074
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdUltimo 
         Height          =   375
         Left            =   1560
         TabIndex        =   11
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
         Picture         =   "FrmDepartamentos.frx":3576
         ImageAlignment  =   1
         TextImageRelation=   4
      End
   End
   Begin VB.TextBox TxtNombre 
      Height          =   285
      Left            =   3600
      MaxLength       =   25
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin MSAdodcLib.Adodc DtaDepartamento 
      Height          =   375
      Left            =   1080
      Top             =   2640
      Width           =   3975
      _ExtentX        =   7011
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
      Caption         =   "DtaDepartamento"
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
   Begin MSDataListLib.DataCombo DBCodigo 
      Bindings        =   "FrmDepartamentos.frx":3A78
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodDepartamento"
      Text            =   ""
   End
   Begin XtremeSuiteControls.PushButton CmdBorrar 
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Borrar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmDepartamentos.frx":3A96
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmDepartamentos.frx":3F4A
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   390
      Left            =   2310
      TabIndex        =   12
      Top             =   180
      Width           =   390
      _Version        =   786432
      _ExtentX        =   688
      _ExtentY        =   688
      _StockProps     =   79
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmDepartamentos.frx":444E
      ImageAlignment  =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "FrmDepartamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Departamento")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaDepartamento.Recordset.MovePrevious

If DtaDepartamento.Recordset.BOF Then
 DtaDepartamento.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 txtNombre.Text = DtaDepartamento.Recordset("departamento")
 DBCodigo.Text = DtaDepartamento.Recordset("CodDepartamento")
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub cmdBorrar_Click()
On Error GoTo TipoErrs
Dim Respuesta, Rsp
If IsNull(DtaDepartamento.Recordset("CodDepartamento")) = True Then
        MsgBox "No Existe Registro"
        Exit Sub
End If
If DBCodigo.Text = "" Then
 MsgBox "No se Puede Eliminar este Registro", vbInformation, "Sistema de Nominas"
 Exit Sub
End If
'Elimino el registro activo en la pantalla

  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando el Departamento: " & txtNombre.Text)
   If Respuesta = 6 Then
     Me.DtaDepartamento.Recordset.Delete
     If IsNull(DtaDepartamento.Recordset("CodDepartamento")) = True Then
        MsgBox "No Existe Registro"
        Exit Sub
     Else
      DBCodigo.Text = ""
      txtNombre.Text = ""
      DtaDepartamento.Refresh
      DtaDepartamento.Recordset.MoveLast
      DtaDepartamento.Recordset.MovePrevious
      CmdAnterior.Enabled = True
     End If
 End If
Exit Sub
TipoErrs:
   ControlErrores
   'MsgBox Err.Number
 Unload Me
  End Sub
Private Sub cmdGrabar_Click()
On Error GoTo TipoErrs
CmdAnterior.Enabled = False
CmdSiguiente.Enabled = False
CmdPrimero.Enabled = False
CmdUltimo.Enabled = False
CmdBorrar.Enabled = False
Salida = False
 If txtNombre.Text = "" Then
  MsgBox "No se Puede dejar la Tabla sin Descripcion", vbCritical, "Error:Sistema de Nominas"
  Exit Sub
 End If
 
 'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
      DtaDepartamento.Refresh
      Do While Not DtaDepartamento.Recordset.EOF
       If DtaDepartamento.Recordset("CodDepartamento") = DBCodigo.Text Then
         'DtaDepartamento.Recordset.Edit
         DtaDepartamento.Recordset("Departamento") = txtNombre.Text
         DtaDepartamento.Recordset.Update
         'DBCodigo.Text = ""
         'TxtNombre.Text = ""
         UbicaDepartamento
         Salida = False
         Exit Sub
             
      End If
      DtaDepartamento.Recordset.MoveNext
      Loop
  'Si despues de Buscar no exite el codigo grabo todos los cambios
         DtaDepartamento.Recordset.AddNew
         DtaDepartamento.Recordset("CodDepartamento") = DBCodigo.Text
         DtaDepartamento.Recordset("Departamento") = txtNombre.Text
         DtaDepartamento.Recordset.Update
         DtaDepartamento.Refresh
         UbicaDepartamento
         Salida = False
         Exit Sub
         
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Private Sub CmdPrimero_Click()
 On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Departamento")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaDepartamento.Recordset.MoveFirst

If DtaDepartamento.Recordset.BOF Then
 DtaDepartamento.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 txtNombre.Text = DtaDepartamento.Recordset("departamento")
 DBCodigo.Text = DtaDepartamento.Recordset("CodDepartamento")
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSiguiente_Click()
On Error GoTo TipoErrs
ValidaSalida ("en la Tabla Departamento")
If Contesta Then
  CmdGrabar.Value = True
End If
DtaDepartamento.Recordset.MoveNext

If DtaDepartamento.Recordset.EOF Then
 DtaDepartamento.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Es esta al final del conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 txtNombre.Text = DtaDepartamento.Recordset("departamento")
 DBCodigo.Text = DtaDepartamento.Recordset("CodDepartamento")
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdUltimo_Click()
 On Error GoTo TipoErrs
ValidaSalida ("en la Tabla Departamento")
If Contesta Then
  CmdGrabar.Value = True
End If
DtaDepartamento.Recordset.MoveLast

If DtaDepartamento.Recordset.EOF Then
 DtaDepartamento.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Es esta al final del conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 txtNombre.Text = DtaDepartamento.Recordset("departamento")
 DBCodigo.Text = DtaDepartamento.Recordset("CodDepartamento")
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub DBCodigo_Change()

On Error GoTo TipoErrs

Evaluar = True
'Al ejecutar algun cambio en el combo actualizo el nombre del departamento
   DtaDepartamento.Refresh
   Do While Not DtaDepartamento.Recordset.EOF
     If DtaDepartamento.Recordset("CodDepartamento") = DBCodigo.Text Then
        txtNombre.Text = DtaDepartamento.Recordset("departamento")
        Salida = False
        Exit Do
     Else
        txtNombre.Text = ""
     End If
       DtaDepartamento.Recordset.MoveNext
   Loop
 Salida = False
 Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub DBCodigo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  txtNombre.SetFocus
  End If
End Sub

Private Sub Form_Activate()
'  If Not BDepartamento = True Then
'  CmdBorrar.Enabled = False
' End If
' If Not GDepartamento = True Then
'   CmdGrabar.Enabled = False
' End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If GDepartamento = True Then
 ValidaSalida ("en la Tabla Departamento")
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

Private Sub TxtNombre_Change()
Salida = True
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
   
  If KeyAscii = 13 Then
   CmdGrabar.SetFocus
  Else
   Evaluar = False
  End If
End Sub

Private Sub Form_Load()
'MDIPrimero.Skin1.ApplySkin
With Me.DtaDepartamento
   .ConnectionString = Conexion
   .RecordSource = "Departamento"
   .Refresh
End With

Me.BackColor = RGB(222, 227, 247)
Me.Frame1.BackColor = RGB(222, 227, 247)

'FrmDepartamentos.CmdSalir.MousePointer = 99
'FrmDepartamentos.cmdGrabar.MousePointer = 99
'FrmDepartamentos.MousePointer = 99
'FrmDepartamentos.DBCodigo.MousePointer = 0
'FrmDepartamentos.txtNombre.MousePointer = 0
End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub
