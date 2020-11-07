VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmCargo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargo"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   HelpContextID   =   25
   Icon            =   "FrmCargo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   840
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
         Picture         =   "FrmCargo.frx":030A
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
         Picture         =   "FrmCargo.frx":080C
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
         Picture         =   "FrmCargo.frx":0D10
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
         Picture         =   "FrmCargo.frx":1212
         ImageAlignment  =   1
         TextImageRelation=   4
      End
   End
   Begin MSDataListLib.DataCombo DBCCargo 
      Bindings        =   "FrmCargo.frx":1714
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodCargo"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaCargo 
      Height          =   375
      Left            =   1080
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
      Caption         =   "DtaCargo"
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
      Left            =   5760
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      AllowPrompt     =   -1  'True
      PromptChar      =   "_"
   End
   Begin VB.TextBox TxtCargo 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3600
      MaxLength       =   35
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin XtremeSuiteControls.PushButton CmdGrabar 
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Grabar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmCargo.frx":172B
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdBorrar 
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Borrar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmCargo.frx":3A8F
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmCargo.frx":3F43
      ImageAlignment  =   0
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto"
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo:"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo Cargo"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FrmCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CloseButton_Click()
Unload Me
End Sub

Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Cargo")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaCargo.Recordset.MovePrevious

If DtaCargo.Recordset.BOF Then
 DtaCargo.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 TxtCargo.Text = DtaCargo.Recordset("Cargo")
 DBCCargo.Text = DtaCargo.Recordset("CodCargo")
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

  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando el Cargo: " & TxtCargo.Text)
   If Respuesta = 6 Then
     DtaCargo.Recordset.Delete
      DBCCargo.Text = ""
      TxtCargo.Text = ""
      MaskEdMonto.Text = ""
      DtaCargo.Recordset.MoveLast
      DtaCargo.Recordset.MovePrevious
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
  
'  If TxtCargo.Text = "" Then
'   MsgBox "Los Campos no Pueden quedar Vacios", vbCritical, "Error:Sistema de Nominas"
'  End If
'
'  If MaskEdMonto.Text = "" Then
'   MsgBox "Los Campos no Pueden quedar Vacios", vbCritical, "Error:Sistema de Nominas"
'  End If
  
  
  'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
      DtaCargo.Refresh
      Do While Not DtaCargo.Recordset.EOF
       If DtaCargo.Recordset("CodCargo") = DBCCargo.Text Then
         'DtaCargo.Recordset.Edit
         DtaCargo.Recordset.Fields("Cargo") = TxtCargo.Text
         DtaCargo.Recordset.Fields("Monto") = MaskEdMonto.Text
         DtaCargo.Recordset.Update
         DtaCargo.Recordset.MoveLast
         DtaCargo.Recordset.MovePrevious
         Salida = False
         Exit Sub
             
      End If
      DtaCargo.Recordset.MoveNext
      Loop
  'Si despues de Buscar no exite el codigo grabo todos los cambios
         DtaCargo.Recordset.AddNew
         DtaCargo.Recordset.Fields("CodCargo") = DBCCargo.Text
         DtaCargo.Recordset.Fields("Cargo") = TxtCargo.Text
         DtaCargo.Recordset.Fields("Monto") = MaskEdMonto.Text
         DtaCargo.Recordset.Update
         DtaCargo.Recordset.MoveLast
         DtaCargo.Recordset.MovePrevious
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
DtaCargo.Recordset.MoveFirst
If DtaCargo.Recordset.BOF Then
 DtaCargo.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCCargo.Text = DtaCargo.Recordset("CodCargo")
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
 If GCargo = True Then
 ValidaSalida ("en la Tabla Cargo")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaCargo.Recordset.MoveNext
If DtaCargo.Recordset.EOF Then
  DtaCargo.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 TxtCargo.Text = DtaCargo.Recordset("Cargo")
 DBCCargo.Text = DtaCargo.Recordset("CodCargo")
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
   DtaCargo.Recordset.MoveLast
If DtaCargo.Recordset.EOF Then
  DtaCargo.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCCargo.Text = DtaCargo.Recordset("CodCargo")
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub DBCCargo_Change()
On Error GoTo TipoErrs
Evaluar = True
'Al ejecutar algun cambio en el combo actualizo el nombre del departamento
   DtaCargo.Refresh
   Do While Not DtaCargo.Recordset.EOF
     If DtaCargo.Recordset("CodCargo") = DBCCargo.Text Then
        TxtCargo.Text = DtaCargo.Recordset("Cargo")
        MaskEdMonto.Text = Format((DtaCargo.Recordset("Monto")), "##,##0.00")
        Exit Do
     Else
        TxtCargo.Text = ""
        MaskEdMonto.Text = "0.00"
     End If
       DtaCargo.Recordset.MoveNext
   Loop
Salida = False
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub DBCCargo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  TxtCargo.SetFocus
 End If
End Sub

Private Sub Form_Activate()
'FrmCargo.DBCCargo.SetFocus
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

With Me.DtaCargo
  .ConnectionString = Conexion
  .RecordSource = "Cargo"
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

Private Sub MaxButton_Click()
   If Me.WindowState = 0 Then
        xpmr.Value = RestoreB
        Me.WindowState = 2
        xpmr.ToolTipText = "Restore"
    ElseIf Me.WindowState = 2 Then
        xpmr.Value = MaxB
        Me.WindowState = 0
        xpmr.ToolTipText = "Maximize"
    End If
End Sub

Private Sub MinButton_Click()
 Me.WindowState = 1
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

End Sub

Private Sub xptopbuttons3_Click()

End Sub

Private Sub xptopbuttons4_Click()

End Sub

Private Sub xp_canvas1_Click()

End Sub
