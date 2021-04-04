VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmRegistroJustifica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Justificaciones"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   7305
   Begin VB.TextBox TxtNombre 
      Height          =   285
      Left            =   3480
      MaxLength       =   25
      TabIndex        =   6
      Top             =   60
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   3135
      Begin XtremeSuiteControls.PushButton CmdAnterior 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Anterior"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmRegistroJustifica.frx":0000
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdSiguiente 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
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
         Picture         =   "FrmRegistroJustifica.frx":0502
         ImageAlignment  =   1
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton CmdPrimero 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Primero"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmRegistroJustifica.frx":0A06
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdUltimo 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
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
         Picture         =   "FrmRegistroJustifica.frx":0F08
         ImageAlignment  =   1
         TextImageRelation=   4
      End
   End
   Begin XtremeSuiteControls.PushButton CmdGrabar 
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   840
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Grabar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmRegistroJustifica.frx":140A
      ImageAlignment  =   0
   End
   Begin MSDataListLib.DataCombo DBCodigo 
      Bindings        =   "FrmRegistroJustifica.frx":376E
      Height          =   315
      Left            =   840
      TabIndex        =   7
      Top             =   60
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Classid"
      Text            =   ""
   End
   Begin XtremeSuiteControls.PushButton CmdBorrar 
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Borrar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmRegistroJustifica.frx":3791
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmRegistroJustifica.frx":3C45
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   390
      Left            =   2190
      TabIndex        =   10
      Top             =   0
      Width           =   390
      _Version        =   786432
      _ExtentX        =   688
      _ExtentY        =   688
      _StockProps     =   79
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmRegistroJustifica.frx":4149
      ImageAlignment  =   0
   End
   Begin MSAdodcLib.Adodc AdoRegistroJustifica 
      Height          =   375
      Left            =   1320
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
      Caption         =   "AdoRegistroJustifica"
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
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   480
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Nuevo"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmRegistroJustifica.frx":464B
      ImageAlignment  =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "FrmRegistroJustifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSiguiente_Click()
On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Registro")
If Contesta Then
  CmdGrabar.Value = True
End If
 AdoRegistroJustifica.Recordset.MoveNext

If AdoRegistroJustifica.Recordset.BOF Then
 AdoRegistroJustifica.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 TxtNombre.Text = AdoRegistroJustifica.Recordset("Classname")
 DBCodigo.Text = AdoRegistroJustifica.Recordset("Classid")
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
 AdoRegistroJustifica.Recordset.MoveLast

If AdoRegistroJustifica.Recordset.BOF Then
 AdoRegistroJustifica.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 TxtNombre.Text = AdoRegistroJustifica.Recordset("Classname")
 DBCodigo.Text = AdoRegistroJustifica.Recordset("Classid")
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub DBCodigo_Change()
On Error GoTo TipoErrs


   Me.AdoRegistroJustifica.Refresh
   Do While Not Me.AdoRegistroJustifica.Recordset.EOF
     If AdoRegistroJustifica.Recordset("Classid") = DBCodigo.Text Then
        TxtNombre.Text = AdoRegistroJustifica.Recordset("Classname")
        Salida = False
        Exit Do
     Else
        TxtNombre.Text = ""
     End If
       AdoRegistroJustifica.Recordset.MoveNext
   Loop
 Salida = False
 Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub Form_Load()
With Me.AdoRegistroJustifica
   .ConnectionString = Conexion
   .RecordSource = "LeaveClass"
   .Refresh
End With

Me.BackColor = RGB(222, 227, 247)
Me.Frame1.BackColor = RGB(222, 227, 247)
End Sub


Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Registro")
If Contesta Then
  CmdGrabar.Value = True
End If
 AdoRegistroJustifica.Recordset.MovePrevious

If AdoRegistroJustifica.Recordset.BOF Then
 AdoRegistroJustifica.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 TxtNombre.Text = AdoRegistroJustifica.Recordset("Classname")
 DBCodigo.Text = AdoRegistroJustifica.Recordset("Classid")
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub cmdborrar_Click()
On Error GoTo TipoErrs
Dim Respuesta, Rsp
If IsNull(AdoRegistroJustifica.Recordset("Classid")) = True Then
        MsgBox "No Existe Registro"
        Exit Sub
End If
If DBCodigo.Text = "" Then
 MsgBox "No se Puede Eliminar este Registro", vbInformation, "Sistema de Nominas"
 Exit Sub
End If
'Elimino el registro activo en la pantalla

  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando el Registro: " & TxtNombre.Text)
   If Respuesta = 6 Then
     Me.AdoRegistroJustifica.Recordset.Delete
     If IsNull(AdoRegistroJustifica.Recordset("Classid")) = True Then
        MsgBox "No Existe Registro"
        Exit Sub
     Else
      DBCodigo.Text = ""
      TxtNombre.Text = ""
      AdoRegistroJustifica.Refresh
      AdoRegistroJustifica.Recordset.MoveLast
      AdoRegistroJustifica.Recordset.MovePrevious
      CmdAnterior.Enabled = True
     End If
 End If
 
 Me.DBCodigo.Text = ""
Me.TxtNombre.Text = ""
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
cmdborrar.Enabled = False
Salida = False
 If TxtNombre.Text = "" Then
  MsgBox "No se Puede dejar la Tabla sin Descripcion", vbCritical, "Error:Sistema de Nominas"
  Exit Sub
 End If
 
 'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
      AdoRegistroJustifica.Refresh
      Do While Not AdoRegistroJustifica.Recordset.EOF
       If AdoRegistroJustifica.Recordset("Classid") = DBCodigo.Text Then
         'DtaDepartamento.Recordset.Edit
         AdoRegistroJustifica.Recordset("Classname") = TxtNombre.Text
         AdoRegistroJustifica.Recordset.Update
         'DBCodigo.Text = ""
         'TxtNombre.Text = ""
         UbicaDepartamento
         Salida = False
         Exit Sub
             
      End If
      AdoRegistroJustifica.Recordset.MoveNext
      Loop
  'Si despues de Buscar no exite el codigo grabo todos los cambios
         AdoRegistroJustifica.Recordset.AddNew
'         AdoRegistroJustifica.Recordset("Classid") = DBCodigo.Text
         AdoRegistroJustifica.Recordset("Classname") = TxtNombre.Text
         AdoRegistroJustifica.Recordset("ViewColor") = "16777215"
         AdoRegistroJustifica.Recordset.Update
         AdoRegistroJustifica.Refresh
'         UbicaDepartamento
         Salida = False
         
         Me.DBCodigo.Text = ""
         Me.TxtNombre.Text = ""

         Exit Sub
         
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Private Sub CmdPrimero_Click()
 On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Justificaciones")
If Contesta Then
  CmdGrabar.Value = True
End If
 AdoRegistroJustifica.Recordset.MoveFirst

If AdoRegistroJustifica.Recordset.BOF Then
 AdoRegistroJustifica.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 TxtNombre.Text = AdoRegistroJustifica.Recordset("Classname")
 DBCodigo.Text = AdoRegistroJustifica.Recordset("Classid")
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub PushButton1_Click()
QueProducto = "Justifica"
FrmConsulta.Show 1
Me.DBCodigo.Text = FrmConsulta.CodigoEmpleado1
End Sub

Private Sub PushButton2_Click()
Me.DBCodigo.Text = ""
Me.TxtNombre.Text = ""

End Sub
