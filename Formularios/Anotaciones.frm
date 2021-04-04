VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmAnotaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anotaciones de la Nomina"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   HelpContextID   =   23
   Icon            =   "Anotaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   582
   Begin MSComCtl2.DTPicker MaskFechaFinaliza 
      Height          =   300
      Left            =   2160
      TabIndex        =   34
      Top             =   4560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      Format          =   51576833
      CurrentDate     =   38465
   End
   Begin MSComCtl2.DTPicker MaskFechaContratacion 
      Height          =   300
      Left            =   2160
      TabIndex        =   33
      Top             =   4080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      Format          =   51576833
      CurrentDate     =   38465
   End
   Begin MSDataListLib.DataCombo DBEmpleado 
      Bindings        =   "Anotaciones.frx":030A
      Height          =   315
      Left            =   2160
      TabIndex        =   32
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodEmpleado"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaCurriculum 
      Height          =   375
      Left            =   4680
      Top             =   6360
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "DtaCurriculum"
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
   Begin MSAdodcLib.Adodc DtaEmpleado 
      Height          =   375
      Left            =   840
      Top             =   6240
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
      Caption         =   "DtaEmpleado"
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
   Begin VB.CommandButton CmdPrint 
      DownPicture     =   "Anotaciones.frx":0324
      Height          =   375
      Left            =   5520
      Picture         =   "Anotaciones.frx":1E06
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton CmdBuscarEmpleado 
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
      Left            =   3840
      Picture         =   "Anotaciones.frx":38E8
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtSalida 
      Height          =   615
      Left            =   6600
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   4320
      Width           =   1935
   End
   Begin VB.ComboBox ComboFaltas 
      Height          =   315
      ItemData        =   "Anotaciones.frx":3A36
      Left            =   2160
      List            =   "Anotaciones.frx":3A40
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox TxtRecomendaciones 
      Height          =   615
      Left            =   6600
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox TxtTrabAnteriores 
      Height          =   615
      Left            =   6600
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox TxtRazones 
      Height          =   615
      Left            =   6600
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox TxtCursos 
      Height          =   645
      Left            =   6600
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox TxtTelEmergencia 
      Height          =   525
      Left            =   6600
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox TxtIdiomas 
      Height          =   615
      Left            =   2160
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton CmdBorrar 
      DownPicture     =   "Anotaciones.frx":3A4C
      Height          =   375
      Left            =   2400
      MouseIcon       =   "Anotaciones.frx":552E
      MousePointer    =   99  'Custom
      Picture         =   "Anotaciones.frx":5838
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton CmdGrabar 
      DownPicture     =   "Anotaciones.frx":731A
      Height          =   375
      Left            =   960
      MouseIcon       =   "Anotaciones.frx":8DFC
      MousePointer    =   99  'Custom
      Picture         =   "Anotaciones.frx":9106
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "Anotaciones.frx":ABE8
      Height          =   375
      Left            =   6960
      MouseIcon       =   "Anotaciones.frx":C6CA
      MousePointer    =   99  'Custom
      Picture         =   "Anotaciones.frx":C9D4
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox TxtDatosRecord 
      Height          =   615
      Left            =   2160
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox TxtJustificaFalta 
      Height          =   645
      Left            =   2160
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox TxtNombre 
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Soluciones Informaticas"
      Height          =   255
      Left            =   5760
      TabIndex        =   14
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Comentarios de Comportamientos de Empleados"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Line Line4 
      X1              =   512
      X2              =   584
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line3 
      X1              =   312
      X2              =   400
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line2 
      X1              =   272
      X2              =   288
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line1 
      X1              =   16
      X2              =   40
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Causa de la Salida:"
      Height          =   255
      Left            =   5040
      TabIndex        =   29
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fecha Finalizacion:"
      Height          =   375
      Left            =   480
      TabIndex        =   28
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fecha Contratacion:"
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Idiomas que Domina:"
      Height          =   375
      Left            =   480
      TabIndex        =   26
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Recomendaciones:"
      Height          =   375
      Left            =   5040
      TabIndex        =   25
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Trab. Anteriores"
      Height          =   375
      Left            =   5040
      TabIndex        =   24
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Razones Contratacion"
      Height          =   375
      Left            =   4920
      TabIndex        =   23
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Conocimientos o Cusos que tiene."
      Height          =   615
      Left            =   5040
      TabIndex        =   22
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Telefono para caso de emergencia ."
      Height          =   615
      Left            =   4920
      TabIndex        =   21
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Datos Record Polcia:"
      Height          =   375
      Left            =   480
      TabIndex        =   20
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Justificacion Falta"
      Height          =   375
      Left            =   480
      TabIndex        =   19
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Faltas "
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nombre Empleado"
      Height          =   375
      Left            =   480
      TabIndex        =   16
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Código Empleado"
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   840
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   5610
      Left            =   0
      Picture         =   "Anotaciones.frx":E4B6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "FrmAnotaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
 
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaCurriculum.Recordset.MovePrevious
       If DtaCurriculum.Recordset.BOF Then
           DtaCurriculum.Recordset.MoveNext
           MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
       Else
           DBEmpleado.Text = DtaCurriculum.Recordset("CodEmpleado")
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
  Set Rsp = DtaCurriculum.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando el Curriculum de: " & txtNombre.Text)
   If Respuesta = 6 Then
     Rsp.Delete
      Limpia
       DtaCurriculum.Recordset.MoveLast
       DtaCurriculum.Recordset.MovePrevious
   End If
Salida = False
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdBuscarEmpleado_Click()
FrmBuscaEmpleado.Show 1
End Sub

Private Sub CmdGrabar_Click()
On Error GoTo TipoErrs
If MaskFechaContratacion.Value = "__/__/____" Then
  MaskFechaContratacion.Value = "01/12/2000"
  MaskFechaFinaliza.Value = "01/12/2001"
End If
Salida = False

'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
     DtaCurriculum.Refresh
      Do While Not DtaCurriculum.Recordset.EOF
        If DtaCurriculum.Recordset("CodEmpleado") = DBEmpleado.Text Then
           'DtaCurriculum.Recordset.Edit
           GuardaRegistro
           Limpia
           DtaCurriculum.Recordset.MoveLast
           DtaCurriculum.Recordset.MovePrevious
           Salida = False
           Exit Sub
        End If
       DtaCurriculum.Recordset.MoveNext
      Loop
    DtaCurriculum.Recordset.AddNew
    GuardaRegistro
    Salida = False
    Exit Sub
TipoErrs:
  ControlErrores
  Unload Me
End Sub

Private Sub CmdPrimero_Click()
 On Error GoTo TipoErrs

If Contesta Then
  CmdGrabar.Value = True
End If
 DtaCurriculum.Recordset.MoveFirst
       If DtaCurriculum.Recordset.BOF Then
           DtaCurriculum.Recordset.MoveNext
           MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
       Else
           DBEmpleado.Text = DtaCurriculum.Recordset("CodEmpleado")
       End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdPrint_Click()
FrmAnotaciones.PrintForm
End Sub

Private Sub CmdSiguiente_Click()
 On Error GoTo TipoErrs

If Contesta Then
  CmdGrabar.Value = True
End If
 DtaCurriculum.Recordset.MoveNext
       If DtaCurriculum.Recordset.EOF Then
           DtaCurriculum.Recordset.MovePrevious
           MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
       Else
           DBEmpleado.Text = DtaCurriculum.Recordset("CodEmpleado")
       End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdUltimo_Click()
 On Error GoTo TipoErrs

If Contesta Then
  CmdGrabar.Value = True
End If
 DtaCurriculum.Recordset.MoveLast
       If DtaCurriculum.Recordset.EOF Then
           DtaCurriculum.Recordset.MovePrevious
           MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
       Else
           DBEmpleado.Text = DtaCurriculum.Recordset("CodEmpleado")
       End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub ComboFaltas_Click()
Salida = True
End Sub

Private Sub ComboFaltas_Change()
PreparaSalida
End Sub

Private Sub ComboFaltas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
justificaFaltas.SetFocus
 End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub


Private Sub DBEmpleado_Change()
On Error GoTo TipoErrs
Dim SqlCurriculum As Variant, Temporal As Variant
Evaluar = True
'Al ejecutar algun cambio en el combo actualizo el nombre del Empleado
   DtaEmpleado.Refresh
'Busco el codigo del empleado para que automaticamente ubique el nombre
 'aunque no existe en la data consulta
    Do While Not DtaEmpleado.Recordset.EOF
     If DtaEmpleado.Recordset("CodEmpleado") = DBEmpleado.Text Then
        txtNombre.Text = DtaEmpleado.Recordset("Nombre1")
        Exit Do
     End If
       DtaEmpleado.Recordset.MoveNext
   Loop
   
 SqlCurriculum = "SELECT Curriculum.CodEmpleado, Empleado.Nombre1, Curriculum.FechaContratacion, Curriculum.FechaFinalizacion, Curriculum.CausaSalida, Curriculum.Faltas, Curriculum.JustificacionFaltas, Curriculum.DatosRecord, Curriculum.Idiomas, Curriculum.TelefonoCasoEmergencia, Curriculum.Cursos, Curriculum.RazonesContratacion, Curriculum.TrabajoAnterior, Curriculum.Recomendaciones FROM Empleado INNER JOIN Curriculum ON Empleado.CodEmpleado = Curriculum.CodEmpleado"
 DtaCurriculum.RecordSource = SqlCurriculum
 DtaCurriculum.Refresh
   
   If DtaCurriculum.Recordset.EOF Then
     MsgBox "El historial esta vacío"
     Exit Sub
   End If
   
   DtaCurriculum.Refresh
   Do While Not DtaCurriculum.Recordset.EOF
     If DtaCurriculum.Recordset("CodEmpleado") = DBEmpleado.Text Then
        txtNombre.Text = DtaCurriculum.Recordset("Nombre1")
        ComboFaltas.Text = DtaCurriculum.Recordset("Faltas")
        TxtJustificaFalta.Text = DtaCurriculum.Recordset("JustificacionFaltas")
        TxtDatosRecord.Text = DtaCurriculum.Recordset("DatosRecord")
        TxtIdiomas.Text = DtaCurriculum.Recordset("Idiomas")
        TxtTelEmergencia.Text = DtaCurriculum.Recordset("TelefonoCasoEmergencia")
        TxtCursos.Text = DtaCurriculum.Recordset("Cursos")
        TxtRazones.Text = DtaCurriculum.Recordset("RazonesContratacion")
        TxtTrabAnteriores.Text = DtaCurriculum.Recordset("TrabajoAnterior")
        TxtRecomendaciones.Text = DtaCurriculum.Recordset("Recomendaciones")
        TxtSalida.Text = DtaCurriculum.Recordset("CausaSalida")
        MaskFechaContratacion.Value = DtaCurriculum.Recordset("FechaContratacion")
        MaskFechaFinaliza.Value = DtaCurriculum.Recordset("FechaFinalizacion")
        Exit Sub
     End If
       DtaCurriculum.Recordset.MoveNext
   Loop
  ComboFaltas.Text = ""
        TxtJustificaFalta.Text = ""
        TxtDatosRecord.Text = ""
        TxtIdiomas.Text = ""
        TxtTelEmergencia.Text = ""
        TxtCursos.Text = ""
        TxtRazones.Text = ""
        TxtTrabAnteriores.Text = ""
        TxtRecomendaciones.Text = ""
        TxtSalida.Text = ""
        MaskFechaContratacion.Value = Format(Now, "dd/mm/yyyy")
        MaskFechaFinaliza.Value = Format(Now, "dd/mm/yyyy")
  
 Salida = False
 Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub DBEmpleado_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  txtNombre.SetFocus
 End If
End Sub

Private Sub Form_Activate()
 On Error GoTo TipoErrs
 Dim SqlCurriculum As String, Temporal As Variant
 DBEmpleado.Text = CodEmpleado
' If Not BAnotaciones = True Then
'   CmdBorrar.Enabled = False
' End If
' If Not GAnotaciones = True Then
'   CmdGrabar.Enabled = False
' End If
 
 SqlCurriculum = "SELECT Curriculum.CodEmpleado, Empleado.Nombre1, Curriculum.FechaContratacion, Curriculum.FechaFinalizacion, Curriculum.CausaSalida, Curriculum.Faltas, Curriculum.JustificacionFaltas, Curriculum.DatosRecord, Curriculum.Idiomas, Curriculum.TelefonoCasoEmergencia, Curriculum.Cursos, Curriculum.RazonesContratacion, Curriculum.TrabajoAnterior, Curriculum.Recomendaciones FROM Empleado INNER JOIN Curriculum ON Empleado.CodEmpleado = Curriculum.CodEmpleado"
 DtaCurriculum.RecordSource = SqlCurriculum
 DtaCurriculum.Refresh
 Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
 End Sub

Private Sub Form_Load()
On Error GoTo TipoErrs
 
 Dim SqlCurriculum As String, Temporal As Variant
With Me.DtaCurriculum
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaEmpleado
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Empleado"
   .Refresh
End With


DBEmpleado.Text = CodEmpleado
 If Not BAnotaciones = True Then
   CmdBorrar.Enabled = False
 End If
 If Not GAnotaciones = True Then
   CmdGrabar.Enabled = False
 End If
 
 SqlCurriculum = "SELECT Curriculum.CodEmpleado, Empleado.Nombre1, Curriculum.FechaContratacion, Curriculum.FechaFinalizacion, Curriculum.CausaSalida, Curriculum.Faltas, Curriculum.JustificacionFaltas, Curriculum.DatosRecord, Curriculum.Idiomas, Curriculum.TelefonoCasoEmergencia, Curriculum.Cursos, Curriculum.RazonesContratacion, Curriculum.TrabajoAnterior, Curriculum.Recomendaciones FROM Empleado INNER JOIN Curriculum ON Empleado.CodEmpleado = Curriculum.CodEmpleado"
 DtaCurriculum.RecordSource = SqlCurriculum
 DtaCurriculum.Refresh
 Exit Sub
TipoErrs:
 ControlErrores
 Unload Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If GAnotaciones = True Then
 ValidaSalida ("en la Tabla Anotaciones")
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

Private Sub MaskFechaContratacion_Change()
Salida = True
End Sub

Private Sub MaskFechaContratacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
MaskFechaFinaliza.SetFocus
Else
   Evaluar = False
   End If
End Sub

Private Sub MaskFechaFinaliza_Change()
Salida = True
End Sub

Private Sub MaskFechaFinaliza_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtTelEmergencia.SetFocus
 Else
   Evaluar = False
   End If
End Sub

Private Sub TxtCursos_Change()
Salida = True
End Sub

Private Sub TxtCursos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtRazones.SetFocus
  Else
   Evaluar = False
   End If
End Sub

Private Sub TxtDatosRecord_Change()
Salida = True
End Sub

Private Sub TxtDatosRecord_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtIdiomas.SetFocus
 Else
   Evaluar = False
   End If
End Sub

Private Sub TxtIdiomas_Change()
Salida = True
End Sub

Private Sub TxtIdiomas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
MaskFechaContratacion.SetFocus
 Else
   Evaluar = False
   End If
End Sub

Private Sub TxtJustificaFalta_Change()
Salida = True
End Sub



Private Sub TxtJustificaFalta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtDatosRecord.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ComboFaltas.SetFocus
 End If

End Sub

Private Sub TxtRazones_Change()
Salida = True
End Sub

Private Sub TxtRazones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtTrabAnteriores.SetFocus
Else
   Evaluar = False
     End If
End Sub

Private Sub TxtRecomendaciones_Change()
Salida = True
End Sub

Private Sub TxtRecomendaciones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtSalida.SetFocus
 Else
   Evaluar = False
     End If
End Sub

Private Sub TxtSalida_Change()
Salida = True
End Sub

Private Sub TxtSalida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CmdGrabar.SetFocus
 Else
   Evaluar = False
   End If
End Sub

Private Sub TxtTelEmergencia_Change()
Salida = True
End Sub

Private Sub TxtTelEmergencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtCursos.SetFocus
 Else
   Evaluar = False
   End If
End Sub

Private Sub TxtTrabAnteriores_Change()
Salida = True
End Sub

Private Sub TxtTrabAnteriores_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtRecomendaciones.SetFocus
 Else
   Evaluar = False
   End If
End Sub
