VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmDeduccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deducciones"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8100
   HelpContextID   =   25
   Icon            =   "FrmDeduccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   170
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   540
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   1200
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
         Picture         =   "FrmDeduccion.frx":030A
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
         Picture         =   "FrmDeduccion.frx":080C
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
         Picture         =   "FrmDeduccion.frx":0D10
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
         Picture         =   "FrmDeduccion.frx":1212
         ImageAlignment  =   1
         TextImageRelation=   4
      End
   End
   Begin VB.TextBox TxtDebito 
      Height          =   315
      Left            =   4080
      MaxLength       =   20
      TabIndex        =   7
      Text            =   "11111"
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox TxtDeduccion 
      Height          =   525
      Left            =   4080
      MaxLength       =   35
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   3855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   2175
      Begin VB.OptionButton OPtVariable 
         Caption         =   "Variable"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptFijo 
         Caption         =   "Fijo"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin MSDataListLib.DataCombo DBCDeduccion 
      Bindings        =   "FrmDeduccion.frx":1714
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodTipoDeduccion"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaDeducciones 
      Height          =   375
      Left            =   720
      Top             =   4320
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "DtaDeducciones"
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
      Left            =   3600
      TabIndex        =   14
      Top             =   1440
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Grabar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmDeduccion.frx":1731
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdBorrar 
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   1920
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Borrar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmDeduccion.frx":3A95
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      Top             =   1920
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmDeduccion.frx":3F49
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton Command1 
      Height          =   390
      Left            =   6240
      TabIndex        =   17
      Top             =   840
      Width           =   390
      _Version        =   786432
      _ExtentX        =   688
      _ExtentY        =   688
      _StockProps     =   79
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmDeduccion.frx":444D
      ImageAlignment  =   0
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta Contable:"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Deducción:"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Deducción:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "FrmDeduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Incentivos")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaDeducciones.Recordset.MovePrevious

If DtaDeducciones.Recordset.BOF Then
 DtaDeducciones.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
       DBCDeduccion.Text = DtaDeducciones.Recordset("codtipodeduccion")
        If DtaDeducciones.Recordset("tipo") = "F" Then
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

Private Sub cmdBorrar_Click()
 On Error GoTo TipoErrs
 Dim Respuesta, Rsp
 
 If DBCDeduccion.Text = "01" Or DBCDeduccion.Text = "02" Or DBCDeduccion.Text = "03" Or DBCDeduccion.Text = "04" Or DBCDeduccion.Text = "05" Then
    MsgBox "Este tipo de Deducción no puede ser modificado"
    Exit Sub
 End If
 
'Elimino el registro activo en la pantalla
  Set Rsp = DtaDeducciones.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando la Deducción: " & txtDeduccion.Text)
   If Respuesta = 6 Then
     Rsp.Delete
      DBCDeduccion.Text = ""
      txtDeduccion.Text = ""
      DtaDeducciones.Recordset.MoveLast
      DtaDeducciones.Recordset.MovePrevious
      Salida = False
   End If
 Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo TipoErrs
  Salida = False
 

 
  If txtDeduccion.Text = "" Then
   MsgBox "Los Campos no Pueden quedar Vacios", vbCritical, "Error:Sistema de Nominas"
   Exit Sub
  End If
  If OptFijo.Value = False And OPtVariable.Value = False Then
    OptFijo.Value = True
  End If
    
  'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
      DtaDeducciones.Refresh
      Do While Not DtaDeducciones.Recordset.EOF
       If DtaDeducciones.Recordset("codtipodeduccion") = DBCDeduccion.Text Then
         'DtaDeducciones.Recordset.Edit

        
        If DBCDeduccion.Text = "01" Or DBCDeduccion.Text = "02" Or DBCDeduccion.Text = "03" Or DBCDeduccion.Text = "04" Or DBCDeduccion.Text = "05" Then
            If Me.TxtDebito.Text <> "" Then
              DtaDeducciones.Recordset("CuentaContable") = Me.TxtDebito.Text
            End If
        Else
            DtaDeducciones.Recordset.Fields("Deduccion") = txtDeduccion.Text
            If OptFijo.Value = True Then
                DtaDeducciones.Recordset("tipo") = "F"
            Else
                DtaDeducciones.Recordset("tipo") = "V"
            End If
            
            If Me.TxtDebito.Text <> "" Then
              DtaDeducciones.Recordset("CuentaContable") = Me.TxtDebito.Text
            End If
        End If
        
         DtaDeducciones.Recordset.Update
         DtaDeducciones.Recordset.MoveLast
         DtaDeducciones.Recordset.MovePrevious
         Salida = False
         Exit Sub
             
      End If
      DtaDeducciones.Recordset.MoveNext
      Loop
      
      
 If DBCDeduccion.Text = "01" Or DBCDeduccion.Text = "02" Then
    MsgBox "Este tipo de Deducción no puede ser modificado"
    Exit Sub
 End If
  'Si despues de Buscar no exite el codigo grabo todos los cambios
         DtaDeducciones.Recordset.AddNew
         DtaDeducciones.Recordset("codtipodeduccion") = DBCDeduccion.Text
           DtaDeducciones.Recordset.Fields("Deduccion") = txtDeduccion.Text
         If OptFijo.Value = True Then
            DtaDeducciones.Recordset("tipo") = "F"
        Else
            DtaDeducciones.Recordset("tipo") = "V"
        End If
        
         If Me.TxtDebito.Text <> "" Then
          DtaDeducciones.Recordset("CuentaContable") = Me.TxtDebito.Text
         End If
         
         DtaDeducciones.Recordset.Update
         DtaDeducciones.Recordset.MoveLast
         DtaDeducciones.Recordset.MovePrevious
         Salida = False
         Exit Sub
         
TipoErrs:
  ControlErrores
End Sub

Private Sub CmdPirmero_Click()
 On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Incentivo")
If Contesta Then
  CmdGrabar.Value = True
End If
DtaDeducciones.Recordset.MoveFirst
If DtaDeducciones.Recordset.BOF Then
 DtaDeducciones.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCDeduccion.Text = DtaDeducciones.Recordset("codtipodeduccion")
End If
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSiguiente_Click()
On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Incentivos")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaDeducciones.Recordset.MoveNext

If DtaDeducciones.Recordset.EOF Then
 DtaDeducciones.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCDeduccion.Text = DtaDeducciones.Recordset("codtipodeduccion")
        If DtaDeducciones.Recordset("tipo") = "F" Then
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
 ValidaSalida ("en la Tabla Incentivos")
If Contesta Then
  CmdGrabar.Value = True
End If
   DtaDeducciones.Recordset.MoveLast
If DtaDeducciones.Recordset.EOF Then
  DtaDeducciones.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCDeduccion.Text = DtaDeducciones.Recordset("codtipodeduccion")
End If
Exit Sub
TipoErrs:
 ControlErrores

End Sub

Private Sub DBCCargo_Change()
End Sub

Private Sub DBCCargo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  TxtCargo.SetFocus
 End If
End Sub

Private Sub DBCIncentivo_Change()
End Sub

Private Sub Command1_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtDebito.Text = CuentaContable
End Sub

Private Sub DBCDeduccion_Change()
On Error GoTo TipoErrs
Evaluar = True
'Al ejecutar algun cambio en el combo actualizo el nombre del departamento
   DtaDeducciones.Refresh
   Do While Not DtaDeducciones.Recordset.EOF
     If DtaDeducciones.Recordset("codtipodeduccion") = DBCDeduccion.Text Then
        txtDeduccion.Text = DtaDeducciones.Recordset("deduccion")
        If Not IsNull(DtaDeducciones.Recordset("CuentaContable")) Then
          Me.TxtDebito.Text = DtaDeducciones.Recordset("CuentaContable")
        End If
        
        If DtaDeducciones.Recordset("tipo") = "F" Then
           OptFijo.Value = True
        Else
           OPtVariable.Value = True
        End If
        Exit Do
     Else
        txtDeduccion.Text = ""
        OptFijo.Value = False
        OPtVariable.Value = False
     End If
       DtaDeducciones.Recordset.MoveNext
   Loop
Salida = False
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me


End Sub

Private Sub DtaDeducciones_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Form_Activate()
'DBCDeduccion.SetFocus
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

With Me.DtaDeducciones
   .ConnectionString = Conexion
   .RecordSource = "TipoDeduccion"
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
