VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmSubsidio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Subsidio"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   HelpContextID   =   25
   Icon            =   "FrmSubsidio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   156
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   533
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   3135
      Begin XtremeSuiteControls.PushButton CmdAnterior 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Anterior"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmSubsidio.frx":030A
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdSiguiente 
         Height          =   375
         Left            =   1560
         TabIndex        =   11
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
         Picture         =   "FrmSubsidio.frx":080C
         ImageAlignment  =   1
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton CmdPrimero 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Primero"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmSubsidio.frx":0D10
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdUltimo 
         Height          =   375
         Left            =   1560
         TabIndex        =   13
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
         Picture         =   "FrmSubsidio.frx":1212
         ImageAlignment  =   1
         TextImageRelation=   4
      End
   End
   Begin VB.TextBox TxtDebito 
      Height          =   315
      Left            =   4560
      MaxLength       =   20
      TabIndex        =   7
      Text            =   "11111"
      Top             =   720
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   2175
      Begin VB.OptionButton OPtVariable 
         Caption         =   "Variable"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptFijo 
         Caption         =   "Fijo"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox TxtSubsidio 
      Height          =   285
      Left            =   4560
      MaxLength       =   35
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin MSDataListLib.DataCombo DBCSubsidio 
      Bindings        =   "FrmSubsidio.frx":1714
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodTipoSubsidio"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaSubsidios 
      Height          =   375
      Left            =   480
      Top             =   4320
      Width           =   2775
      _ExtentX        =   4895
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
   Begin XtremeSuiteControls.PushButton CmdGrabar 
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   1320
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Grabar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmSubsidio.frx":172F
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdBorrar 
      Height          =   375
      Left            =   3840
      TabIndex        =   15
      Top             =   1800
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Borrar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmSubsidio.frx":3A93
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   1800
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmSubsidio.frx":3F47
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton Command1 
      Height          =   390
      Left            =   6720
      TabIndex        =   17
      Top             =   720
      Width           =   390
      _Version        =   786432
      _ExtentX        =   688
      _ExtentY        =   688
      _StockProps     =   79
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmSubsidio.frx":444B
      ImageAlignment  =   0
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta Contable:"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Subsidio:"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Subsidio:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "FrmSubsidio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Subsidios")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaSubsidios.Recordset.MovePrevious

If DtaSubsidios.Recordset.BOF Then
 DtaSubsidios.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
       DBCSubsidio.Text = DtaSubsidios.Recordset("CodtipoSubsidio")
        If DtaSubsidios.Recordset("tipo") = "F" Then
           OptFijo.Value = True
        Else
           OptFijo.Value = False
        End If
 
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub cmdborrar_Click()
 On Error GoTo TipoErrs
 Dim Respuesta, Rsp
'Elimino el registro activo en la pantalla
  Set Rsp = DtaSubsidios.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando el Subsidio: " & txtSubsidio.Text)
   If Respuesta = 6 Then
     Rsp.Delete
      DBCSubsidio.Text = ""
      txtSubsidio.Text = ""
      DtaSubsidios.Recordset.MoveLast
      DtaSubsidios.Recordset.MovePrevious
      Salida = False
   End If
 Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo TipoErrs
  Salida = False
 
  If txtSubsidio.Text = "" Then
   MsgBox "Los Campos no Pueden quedar Vacios", vbCritical, "Error:Sistema de Nominas"
   Exit Sub
  End If
  If OptFijo.Value = False And OPtVariable.Value = False Then
    OptFijo.Value = True
  End If
    
  'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
      DtaSubsidios.Refresh
      Do While Not DtaSubsidios.Recordset.EOF
       If DtaSubsidios.Recordset("CodtipoSubsidio") = DBCSubsidio.Text Then
         'DtaSubsidios.Recordset.Edit
         DtaSubsidios.Recordset.Fields("Subsidio") = txtSubsidio.Text
         
        If Me.TxtDebito.Text <> "" Then
         Me.DtaSubsidios.Recordset("CuentaContable") = Me.TxtDebito.Text
        End If
        
         If OptFijo.Value = True Then
            DtaSubsidios.Recordset("tipo") = "F"
        Else
            DtaSubsidios.Recordset("tipo") = "V"
        End If
         DtaSubsidios.Recordset.Update
         DtaSubsidios.Recordset.MoveLast
         DtaSubsidios.Recordset.MovePrevious
         Salida = False
         Exit Sub
             
      End If
      DtaSubsidios.Recordset.MoveNext
      Loop
  'Si despues de Buscar no exite el codigo grabo todos los cambios
         DtaSubsidios.Recordset.AddNew
         DtaSubsidios.Recordset("CodtipoSubsidio") = DBCSubsidio.Text
           DtaSubsidios.Recordset.Fields("Subsidio") = txtSubsidio.Text
           
        If Me.TxtDebito.Text <> "" Then
         Me.DtaSubsidios.Recordset("CuentaContable") = Me.TxtDebito.Text
        End If
        
         If OptFijo.Value = True Then
            DtaSubsidios.Recordset("tipo") = "F"
        Else
            DtaSubsidios.Recordset("tipo") = "V"
        End If
         DtaSubsidios.Recordset.Update
         DtaSubsidios.Recordset.MoveLast
         DtaSubsidios.Recordset.MovePrevious
         Salida = False
         Exit Sub
         
TipoErrs:
  ControlErrores
End Sub

Private Sub CmdPirmero_Click()
 On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Subsidio")
If Contesta Then
  CmdGrabar.Value = True
End If
DtaSubsidios.Recordset.MoveFirst
If DtaSubsidios.Recordset.BOF Then
 DtaSubsidios.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCSubsidio.Text = DtaSubsidios.Recordset("CodtipoSubsidio")
End If
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSiguiente_Click()
On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Subsidios")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaSubsidios.Recordset.MoveNext

If DtaSubsidios.Recordset.EOF Then
 DtaSubsidios.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCSubsidio.Text = DtaSubsidios.Recordset("CodtipoSubsidio")
        If DtaSubsidios.Recordset("tipo") = "F" Then
           OptFijo.Value = True
        Else
           OptFijo.Value = False
        End If
 
End If
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub CmdUltimo_Click()
On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Subsidios")
If Contesta Then
  CmdGrabar.Value = True
End If
   DtaSubsidios.Recordset.MoveLast
If DtaSubsidios.Recordset.EOF Then
  DtaSubsidios.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCSubsidio.Text = DtaSubsidios.Recordset("CodtipoSubsidio")
End If
Exit Sub
TipoErrs:
 ControlErrores

End Sub

Private Sub DBCCargo_Change()
End Sub

Private Sub DBCCargo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  txtCargo.SetFocus
 End If
End Sub

Private Sub DBCIncentivo_Change()
End Sub

Private Sub DBCDeduccion_Change()


End Sub

Private Sub Command1_Click()
'QueProducto = "CuentaContable"
'FrmConsulta.Show 1
'Me.TxtDebito.Text = CuentaContable
FrmIncentivos.Show 1
End Sub

Private Sub DBCSubsidio_Change()
On Error GoTo TipoErrs
Evaluar = True
'Al ejecutar algun cambio en el combo actualizo el nombre del departamento
   DtaSubsidios.Refresh
   Do While Not DtaSubsidios.Recordset.EOF
     If DtaSubsidios.Recordset("CodtipoSubsidio") = DBCSubsidio.Text Then
        txtSubsidio.Text = DtaSubsidios.Recordset("Subsidio")
        
        If Not IsNull(Me.DtaSubsidios.Recordset("CuentaContable")) Then
         Me.TxtDebito.Text = Me.DtaSubsidios.Recordset("CuentaContable")
        End If
        
        If DtaSubsidios.Recordset("tipo") = "F" Then
           OptFijo.Value = True
        Else
           OPtVariable.Value = True
        End If
        Exit Do
     Else
        txtSubsidio.Text = ""
        OptFijo.Value = False
        OPtVariable.Value = False
     End If
       DtaSubsidios.Recordset.MoveNext
   Loop
Salida = False
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me

End Sub

Private Sub Form_Activate()
DBCSubsidio.SetFocus
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

With Me.DtaSubsidios
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "TipoSubsidio"
   .Refresh
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If GEmpleado = True Then
ValidaSalida ("en la Tabla Deduccion")
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

Private Sub Label3_Click()

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
