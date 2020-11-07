VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form FrmTipoIncapacidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Incapacidades"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8715
   HelpContextID   =   30
   Icon            =   "FrmTipoIncapacidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   132
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   581
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   3135
      Begin XtremeSuiteControls.PushButton CmdAnterior 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Anterior"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmTipoIncapacidad.frx":030A
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdSiguiente 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
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
         Picture         =   "FrmTipoIncapacidad.frx":080C
         ImageAlignment  =   1
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton CmdPrimero 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Primero"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmTipoIncapacidad.frx":0D10
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdUltimo 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
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
         Picture         =   "FrmTipoIncapacidad.frx":1212
         ImageAlignment  =   1
         TextImageRelation=   4
      End
   End
   Begin VB.TextBox TxtIncapacidad 
      Height          =   285
      Left            =   4560
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
   Begin MSDataListLib.DataCombo DBCIncapacidad 
      Bindings        =   "FrmTipoIncapacidad.frx":1714
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodIncapacidad"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaTipoIncapacidad 
      Height          =   495
      Left            =   600
      Top             =   3720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   873
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
      Caption         =   "DtaTipoIncapacidad"
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
      TabIndex        =   9
      Top             =   960
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Grabar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmTipoIncapacidad.frx":1735
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdBorrar 
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   1440
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Borrar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmTipoIncapacidad.frx":3A99
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmTipoIncapacidad.frx":3F4D
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton Command1 
      Height          =   390
      Left            =   3030
      TabIndex        =   12
      Top             =   180
      Width           =   390
      _Version        =   786432
      _ExtentX        =   688
      _ExtentY        =   688
      _StockProps     =   79
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmTipoIncapacidad.frx":4451
      ImageAlignment  =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Incapcidad:"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Incapacidad:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "FrmTipoIncapacidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAnterior_Click()
 On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Tipo Incapacidad")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaTipoIncapacidad.Recordset.MovePrevious

If DtaTipoIncapacidad.Recordset.BOF Then
 DtaTipoIncapacidad.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCIncapacidad.Text = DtaTipoIncapacidad.Recordset("CodIncapacidad")
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

  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando el Tipo de Incapacidad: " & TxtIncapacidad.Text)
   If Respuesta = 6 Then
     DtaTipoIncapacidad.Recordset.Delete
       DBCIncapacidad.Text = ""
       TxtIncapacidad.Text = ""
       DtaTipoIncapacidad.Recordset.MoveLast
       DtaTipoIncapacidad.Recordset.MovePrevious
       Salida = False
   End If
Exit Sub
TipoErrs:
 ControlErrores
 
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo TipoErrs
 Salida = False
 If TxtIncapacidad.Text = "" Then
  MsgBox "No se Puede Dejar El Campo Vacio", vbCritical, "Error:Sistema de Nominas"
  Exit Sub
 End If
 'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
      DtaTipoIncapacidad.Refresh
      Do While Not DtaTipoIncapacidad.Recordset.EOF
       If DtaTipoIncapacidad.Recordset("CodIncapacidad") = DBCIncapacidad.Text Then
          'DtaTipoIncapacidad.Recordset.Edit
          DtaTipoIncapacidad.Recordset.Fields("Incapacidad") = TxtIncapacidad.Text
          DtaTipoIncapacidad.Recordset.Update
          DBCIncapacidad.Text = ""
          DtaTipoIncapacidad.Recordset.MoveLast
          DtaTipoIncapacidad.Recordset.MovePrevious
          Salida = False
         Exit Sub
             
      End If
      DtaTipoIncapacidad.Recordset.MoveNext
      Loop
  'Si despues de Buscar no exite el codigo grabo todos los cambios
         DtaTipoIncapacidad.Recordset.AddNew
         DtaTipoIncapacidad.Recordset.Fields("Incapacidad") = TxtIncapacidad.Text
         DtaTipoIncapacidad.Recordset.Fields("CodIncapacidad") = DBCIncapacidad.Text
         DtaTipoIncapacidad.Recordset.Update
         DBCIncapacidad.Text = ""
          DtaTipoIncapacidad.Recordset.MoveLast
          DtaTipoIncapacidad.Recordset.MovePrevious
          Salida = False
         Exit Sub
         
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Private Sub CmdPrimero_Click()
 On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Tipo Incapacidad")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaTipoIncapacidad.Recordset.MoveFirst

If DtaTipoIncapacidad.Recordset.BOF Then
 DtaTipoIncapacidad.Recordset.MoveNext
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCIncapacidad.Text = DtaTipoIncapacidad.Recordset("CodIncapacidad")
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
ValidaSalida ("en la Tabla Tipo Incapacidad")
If Contesta Then
  CmdGrabar.Value = True
End If
DtaTipoIncapacidad.Recordset.MoveNext

If DtaTipoIncapacidad.Recordset.EOF Then
 DtaTipoIncapacidad.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCIncapacidad.Text = DtaTipoIncapacidad.Recordset("CodIncapacidad")
 End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdUltimo_Click()
On Error GoTo TipoErrs
ValidaSalida ("en la Tabla Tipo Incapacidad")
If Contesta Then
  CmdGrabar.Value = True
End If
DtaTipoIncapacidad.Recordset.MoveLast

If DtaTipoIncapacidad.Recordset.EOF Then
 DtaTipoIncapacidad.Recordset.MovePrevious
 MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
Else
 DBCIncapacidad.Text = DtaTipoIncapacidad.Recordset("CodIncapacidad")
 End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub DBCIncapacidad_Change()
 
Evaluar = True
'Al ejecutar algun cambio en el combo actualizo el nombre del departamento
   DtaTipoIncapacidad.Refresh
   Do While Not Me.DtaTipoIncapacidad.Recordset.EOF
     If DtaTipoIncapacidad.Recordset("CodIncapacidad") = DBCIncapacidad.Text Then
        TxtIncapacidad.Text = DtaTipoIncapacidad.Recordset("incapacidad")
        Exit Do
     Else
        TxtIncapacidad.Text = ""
     End If
       DtaTipoIncapacidad.Recordset.MoveNext
   Loop
Salida = False
End Sub

Private Sub DBCIncapacidad_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  TxtIncapacidad.SetFocus
  End If
End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub Form_Activate()
 If Not BTipoIncapacidad = True Then
   CmdBorrar.Enabled = False
 End If
 If Not GTipoIncapacidad = True Then
   CmdGrabar.Enabled = False
 End If
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(222, 227, 247)
Me.Frame1.BackColor = RGB(222, 227, 247)

With Me.DtaTipoIncapacidad
   .ConnectionString = Conexion
   .RecordSource = "TipoIncapacidad"
   .Refresh
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If GTipoIncapacidad = True Then
ValidaSalida ("en la Tabla Tipo Incapacidad")
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

Private Sub TxtIncapacidad_Change()
Salida = True
End Sub

Private Sub TxtIncapacidad_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
 CmdGrabar.SetFocus
  Else
   Evaluar = False
  End If
End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub
