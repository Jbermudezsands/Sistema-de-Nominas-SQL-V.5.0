VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmTipoComision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Comisiones"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   HelpContextID   =   25
   Icon            =   "FrmTipoComision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   133
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   537
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6120
      Picture         =   "FrmTipoComision.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox TxtDebito 
      Height          =   315
      Left            =   3960
      MaxLength       =   20
      TabIndex        =   14
      Text            =   "11111"
      Top             =   600
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   3255
      Begin VB.CommandButton CmdUltimo 
         DownPicture     =   "FrmTipoComision.frx":0458
         Height          =   375
         Left            =   1560
         MouseIcon       =   "FrmTipoComision.frx":1F3A
         MousePointer    =   99  'Custom
         Picture         =   "FrmTipoComision.frx":237C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton CmdPirmero 
         DownPicture     =   "FrmTipoComision.frx":3E5E
         Height          =   375
         Left            =   120
         MouseIcon       =   "FrmTipoComision.frx":5940
         MousePointer    =   99  'Custom
         Picture         =   "FrmTipoComision.frx":5D82
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton CmdSiguiente 
         DownPicture     =   "FrmTipoComision.frx":7864
         Height          =   375
         Left            =   1560
         MouseIcon       =   "FrmTipoComision.frx":9346
         MousePointer    =   99  'Custom
         Picture         =   "FrmTipoComision.frx":9788
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdAnterior 
         DownPicture     =   "FrmTipoComision.frx":B26A
         Height          =   375
         Left            =   120
         MouseIcon       =   "FrmTipoComision.frx":CD4C
         MousePointer    =   99  'Custom
         Picture         =   "FrmTipoComision.frx":D18E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdGrabar 
      DownPicture     =   "FrmTipoComision.frx":EC70
      Height          =   375
      Left            =   3720
      MouseIcon       =   "FrmTipoComision.frx":10752
      MousePointer    =   99  'Custom
      Picture         =   "FrmTipoComision.frx":10B94
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalir 
      DownPicture     =   "FrmTipoComision.frx":12676
      Height          =   375
      Left            =   6000
      MouseIcon       =   "FrmTipoComision.frx":14158
      MousePointer    =   99  'Custom
      Picture         =   "FrmTipoComision.frx":1459A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton CmdBorrar 
      DownPicture     =   "FrmTipoComision.frx":1607C
      Height          =   375
      Left            =   3720
      MouseIcon       =   "FrmTipoComision.frx":17B5E
      MousePointer    =   99  'Custom
      Picture         =   "FrmTipoComision.frx":17FA0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox TxtComision 
      Height          =   285
      Left            =   3960
      MaxLength       =   35
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo DBCComision 
      Bindings        =   "FrmTipoComision.frx":19A82
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodTipoComision"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaTipoComision 
      Height          =   375
      Left            =   360
      Top             =   3720
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
      Caption         =   "DtaTipoComision"
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
      Left            =   6120
      TabIndex        =   4
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      AllowPrompt     =   -1  'True
      Format          =   "0%"
      PromptChar      =   "_"
   End
   Begin VB.Label Label13 
      Caption         =   "Cuenta Contable:"
      Height          =   255
      Left            =   2520
      TabIndex        =   16
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "%"
      Height          =   255
      Left            =   5760
      TabIndex        =   13
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Comision:"
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Código Comision:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FrmTipoComision"
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
 DtaTipoComision.Recordset.MovePrevious

If DtaTipoComision.Recordset.BOF Then
 DtaTipoComision.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 TxtComision.Text = DtaTipoComision.Recordset("Comision")
 DBCComision.Text = DtaTipoComision.Recordset("Codtipocomision")
 'MaskEdMonto.Text = DtaCargo.Recordset.Monto
End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdBorrar_Click()
 On Error GoTo TipoErrs
 Dim Respuesta, Rsp
'Elimino el registro activo en la pantalla
  Set Rsp = DtaTipoComision.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando la Comisión: " & TxtComision.Text)
   If Respuesta = 6 Then
     Rsp.Delete
      DBCComision.Text = ""
      TxtComision.Text = ""
      MaskEdMonto.Text = ""
      DtaTipoComision.Recordset.MoveLast
      DtaTipoComision.Recordset.MovePrevious
      Salida = False
   End If
 Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdGrabar_Click()
'On Error GoTo TipoErrs
 Dim Comision As Double
  Salida = False
  If TxtComision.Text = "" Or MaskEdMonto = "" Then
   MsgBox "Los Campos no Pueden quedar Vacios", vbCritical, "Error:Sistema de Nominas"
  End If
  
  Comision = Val(MaskEdMonto.Text) / 100
  'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
      DtaTipoComision.Refresh
      Do While Not DtaTipoComision.Recordset.EOF
       If DtaTipoComision.Recordset("Codtipocomision") = DBCComision.Text Then
         'DtaTipoComision.Recordset.Edit
         DtaTipoComision.Recordset.Fields("Comision") = TxtComision.Text
         DtaTipoComision.Recordset.Fields("Porcentaje") = Comision
         If Me.TxtDebito.Text <> "" Then
          Me.DtaTipoComision.Recordset("CuentaContable") = Me.TxtDebito.Text
         End If
         DtaTipoComision.Recordset.Update
         DtaTipoComision.Recordset.MoveLast
         DtaTipoComision.Recordset.MovePrevious
         Salida = False
         Exit Sub
             
      End If
      DtaTipoComision.Recordset.MoveNext
      Loop
  'Si despues de Buscar no exite el codigo grabo todos los cambios
         DtaTipoComision.Recordset.AddNew
         DtaTipoComision.Recordset.Fields("CodTipoComision") = DBCComision.Text
         DtaTipoComision.Recordset.Fields("Comision") = TxtComision.Text
         DtaTipoComision.Recordset.Fields("Porcentaje") = Comision
         If Me.TxtDebito.Text <> "" Then
          Me.DtaTipoComision.Recordset("CuentaContable") = Me.TxtDebito.Text
         End If
         DtaTipoComision.Recordset.Update
         DtaTipoComision.Recordset.MoveLast
         DtaTipoComision.Recordset.MovePrevious
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
DtaTipoComision.Recordset.MoveFirst
If DtaTipoComision.Recordset.BOF Then
 DtaTipoComision.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCComision.Text = DtaTipoComision.Recordset("Codtipocomision")
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
 DtaTipoComision.Recordset.MoveNext
If DtaTipoComision.Recordset.EOF Then
  DtaTipoComision.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 TxtComision.Text = DtaTipoComision.Recordset("Comision")
 DBCComision.Text = DtaTipoComision.Recordset("Codtipocomision")
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
   DtaTipoComision.Recordset.MoveLast
If DtaTipoComision.Recordset.EOF Then
  DtaTipoComision.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCComision.Text = DtaTipoComision.Recordset("Codtipocomision")
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
  txtCargo.SetFocus
 End If
End Sub

Private Sub DBCDestajo_Change()

End Sub

Private Sub Command1_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtDebito.Text = CuentaContable
End Sub

Private Sub DBCComision_Change()
On Error GoTo TipoErrs
Evaluar = True
'Al ejecutar algun cambio en el combo actualizo el nombre del departamento
   DtaTipoComision.Refresh
   Do While Not DtaTipoComision.Recordset.EOF
     If DtaTipoComision.Recordset("Codtipocomision") = DBCComision.Text Then
        TxtComision.Text = DtaTipoComision.Recordset("Comision")
        MaskEdMonto.Text = Format((DtaTipoComision.Recordset("porcentaje")), "0%")
        
        If Not IsNull(Me.DtaTipoComision.Recordset("CuentaContable")) Then
        Me.TxtDebito.Text = Me.DtaTipoComision.Recordset("CuentaContable")
        End If
        
        Exit Do
     Else
        TxtComision.Text = ""
        MaskEdMonto.Text = "0.00"
     End If
       DtaTipoComision.Recordset.MoveNext
   Loop
Salida = False
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub Form_Activate()
'DBCComision.SetFocus
' If Not BCargo = True Then
'   CmdBorrar.Enabled = False
' End If
' If Not GCargo = True Then
'  CmdGrabar.Enabled = False
' End If
End Sub

Private Sub Form_Load()

With Me.DtaTipoComision
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "TipoComision"
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
