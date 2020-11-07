VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmIncentivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Incentivos"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8130
   HelpContextID   =   25
   Icon            =   "FrmIncentivo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   152
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   542
   Begin VB.TextBox TxtDebito 
      Height          =   285
      Left            =   4200
      MaxLength       =   20
      TabIndex        =   16
      Text            =   "11111"
      Top             =   720
      Width           =   2055
   End
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
      Left            =   6360
      Picture         =   "FrmIncentivo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox TxtIncentivo 
      Height          =   285
      Left            =   4200
      MaxLength       =   35
      TabIndex        =   14
      Top             =   360
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   2175
      Begin VB.OptionButton OPtVariable 
         Caption         =   "Variable"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptFijo 
         Caption         =   "Fijo"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   3135
      Begin VB.CommandButton CmdSiguiente 
         DownPicture     =   "FrmIncentivo.frx":0458
         Height          =   375
         Left            =   1560
         MouseIcon       =   "FrmIncentivo.frx":1F3A
         MousePointer    =   99  'Custom
         Picture         =   "FrmIncentivo.frx":237C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdAnterior 
         DownPicture     =   "FrmIncentivo.frx":3E5E
         Height          =   375
         Left            =   120
         MouseIcon       =   "FrmIncentivo.frx":5940
         MousePointer    =   99  'Custom
         Picture         =   "FrmIncentivo.frx":5D82
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdPirmero 
         DownPicture     =   "FrmIncentivo.frx":7864
         Height          =   375
         Left            =   120
         MouseIcon       =   "FrmIncentivo.frx":9346
         MousePointer    =   99  'Custom
         Picture         =   "FrmIncentivo.frx":9788
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton CmdUltimo 
         DownPicture     =   "FrmIncentivo.frx":B26A
         Height          =   375
         Left            =   1560
         MouseIcon       =   "FrmIncentivo.frx":CD4C
         MousePointer    =   99  'Custom
         Picture         =   "FrmIncentivo.frx":D18E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdGrabar 
      DownPicture     =   "FrmIncentivo.frx":EC70
      Height          =   375
      Left            =   3840
      MouseIcon       =   "FrmIncentivo.frx":10752
      MousePointer    =   99  'Custom
      Picture         =   "FrmIncentivo.frx":10B94
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalir 
      DownPicture     =   "FrmIncentivo.frx":12676
      Height          =   375
      Left            =   6360
      MouseIcon       =   "FrmIncentivo.frx":14158
      MousePointer    =   99  'Custom
      Picture         =   "FrmIncentivo.frx":1459A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton CmdBorrar 
      DownPicture     =   "FrmIncentivo.frx":1607C
      Height          =   375
      Left            =   3840
      MouseIcon       =   "FrmIncentivo.frx":17B5E
      MousePointer    =   99  'Custom
      Picture         =   "FrmIncentivo.frx":17FA0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo DBCIncentivo 
      Bindings        =   "FrmIncentivo.frx":19A82
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodTipoIncentivo"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaIncentivos 
      Height          =   375
      Left            =   1080
      Top             =   3960
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
      Caption         =   "DtaIncentivos"
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
   Begin VB.Label Label13 
      Caption         =   "Cuenta Contable:"
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Incentivo:"
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Código Incentivo:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "FrmIncentivo"
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
 DtaIncentivos.Recordset.MovePrevious

If DtaIncentivos.Recordset.BOF Then
 DtaIncentivos.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
       DBCIncentivo.Text = DtaIncentivos.Recordset("CodtipoIncentivo")
        If DtaIncentivos.Recordset("tipo") = "F" Then
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

Private Sub CmdBorrar_Click()
 On Error GoTo TipoErrs
'If Me.DBCIncentivo.Text = "01" Or Me.DBCIncentivo.Text = "02" Then
 'Exit Sub
'End If
 Dim Respuesta, Rsp
'Elimino el registro activo en la pantalla
  Set Rsp = DtaIncentivos.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando el Incentivo: " & TxtIncentivo.Text)
   If Respuesta = 6 Then
     Rsp.Delete
      DBCIncentivo.Text = ""
      TxtIncentivo.Text = ""
      DtaIncentivos.Recordset.MoveLast
      DtaIncentivos.Recordset.MovePrevious
      Salida = False
   End If
 Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo TipoErrs

'If Me.DBCIncentivo.Text = "01" Or Me.DBCIncentivo.Text = "02" Then
 'Exit Sub
'End If

  Salida = False
 
  If TxtIncentivo.Text = "" Then
   MsgBox "Los Campos no Pueden quedar Vacios", vbCritical, "Error:Sistema de Nominas"
   Exit Sub
  End If
  If OptFijo.Value = False And OPtVariable.Value = False Then
    OptFijo.Value = True
  End If
    
  'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
      DtaIncentivos.Refresh
      Do While Not DtaIncentivos.Recordset.EOF
       If DtaIncentivos.Recordset("CodtipoIncentivo") = DBCIncentivo.Text Then
         'DtaIncentivos.Recordset.Edit
         DtaIncentivos.Recordset.Fields("Incentivo") = TxtIncentivo.Text
         If OptFijo.Value = True Then
            DtaIncentivos.Recordset("tipo") = "F"
        Else
            DtaIncentivos.Recordset("tipo") = "V"
        End If
         If Me.TxtDebito.Text <> "" Then
          Me.DtaIncentivos.Recordset("CuentaContable") = Me.TxtDebito.Text
         End If
         DtaIncentivos.Recordset.Update
         DtaIncentivos.Recordset.MoveLast
         DtaIncentivos.Recordset.MovePrevious
         Salida = False
         Exit Sub
             
      End If
      DtaIncentivos.Recordset.MoveNext
      Loop
  'Si despues de Buscar no exite el codigo grabo todos los cambios
         DtaIncentivos.Recordset.AddNew
         DtaIncentivos.Recordset("CodtipoIncentivo") = DBCIncentivo.Text
           DtaIncentivos.Recordset.Fields("Incentivo") = TxtIncentivo.Text
         If OptFijo.Value = True Then
            DtaIncentivos.Recordset("tipo") = "F"
        Else
            DtaIncentivos.Recordset("tipo") = "V"
        End If
         DtaIncentivos.Recordset.Update
         DtaIncentivos.Recordset.MoveLast
         DtaIncentivos.Recordset.MovePrevious
         Salida = False
         Exit Sub
         
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Private Sub CmdPirmero_Click()
 On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Incentivo")
If Contesta Then
  CmdGrabar.Value = True
End If
DtaIncentivos.Recordset.MoveFirst
If DtaIncentivos.Recordset.BOF Then
 DtaIncentivos.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCIncentivo.Text = DtaIncentivos.Recordset("CodtipoIncentivo")
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
 ValidaSalida ("en la Tabla Incentivos")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaIncentivos.Recordset.MoveNext

If DtaIncentivos.Recordset.EOF Then
 DtaIncentivos.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCIncentivo.Text = DtaIncentivos.Recordset("CodtipoIncentivo")
        If DtaIncentivos.Recordset("tipo") = "F" Then
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
   DtaIncentivos.Recordset.MoveLast
If DtaIncentivos.Recordset.EOF Then
  DtaIncentivos.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCIncentivo.Text = DtaIncentivos.Recordset("CodtipoIncentivo")
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

Private Sub Command1_Click()
QueProducto = "CuentaContable"
FrmConsulta.Show 1
Me.TxtDebito.Text = CuentaContable



End Sub

Private Sub DBCIncentivo_Change()
'On Error GoTo TipoErrs
Evaluar = True
'Al ejecutar algun cambio en el combo actualizo el nombre del departamento
   DtaIncentivos.Refresh
   Do While Not DtaIncentivos.Recordset.EOF
     If DtaIncentivos.Recordset("CodtipoIncentivo") = DBCIncentivo.Text Then
        TxtIncentivo.Text = DtaIncentivos.Recordset("incentivo")
        If Not IsNull(Me.DtaIncentivos.Recordset("CuentaContable")) Then
         Me.TxtDebito.Text = Me.DtaIncentivos.Recordset("CuentaContable")
        Else
         Me.TxtDebito.Text = ""
        End If
        
        If DtaIncentivos.Recordset("tipo") = "F" Then
           OptFijo.Value = True
        Else
           OPtVariable.Value = True
        End If
        Exit Do
     Else
        Me.TxtDebito.Text = ""
        TxtIncentivo.Text = ""
        OptFijo.Value = False
        OPtVariable.Value = False
     End If
       DtaIncentivos.Recordset.MoveNext
   Loop
Salida = False
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me

End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub Form_Activate()
FrmIncentivo.DBCIncentivo.SetFocus
 If Not BCargo = True Then
'   CmdBorrar.Enabled = False
 End If
 If Not GCargo = True Then
'  CmdGrabar.Enabled = False
 End If
End Sub

Private Sub Form_Load()
With Me.DtaIncentivos
   .ConnectionString = Conexion
   .RecordSource = "TipoIncentivo"
   .Refresh
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If GEmpleado = True Then
ValidaSalida ("en la Tabla Incentivo")
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
