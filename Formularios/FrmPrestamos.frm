VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FrmPrestamos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Prestamos"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   HelpContextID   =   27
   Icon            =   "FrmPrestamos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6945
   Begin VB.TextBox TxtMontoPagar 
      Height          =   285
      Left            =   4920
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox TxtSaldoPrestamo 
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox TxtCodPrestamo 
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton CmdUltimo 
      Caption         =   "Ultimo"
      Height          =   375
      Left            =   3960
      MouseIcon       =   "FrmPrestamos.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton CmdPrimero 
      Caption         =   "Primero"
      Height          =   375
      Left            =   3000
      MouseIcon       =   "FrmPrestamos.frx":0884
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox TxtNombre 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Data DtaEmpleado 
      Caption         =   "DtaEmpleado"
      Connect         =   "Access"
      DatabaseName    =   "C:\Zeus Nominas\Nominas.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Empleado"
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSMask.MaskEdBox MaskEdFecha 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdMonto 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Data DtaPrestamos 
      Caption         =   "DtaPrestamos"
      Connect         =   "Access"
      DatabaseName    =   "C:\Zeus Nominas\Nominas.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Prestamos"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBCtls.DBCombo DBCCodEmpleado 
      Bindings        =   "FrmPrestamos.frx":0CC6
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodEmpleado"
      Text            =   ""
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5880
      MouseIcon       =   "FrmPrestamos.frx":0CE0
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   120
      MouseIcon       =   "FrmPrestamos.frx":1122
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   1080
      MouseIcon       =   "FrmPrestamos.frx":1564
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   2040
      MouseIcon       =   "FrmPrestamos.frx":19A6
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   4920
      MouseIcon       =   "FrmPrestamos.frx":1DE8
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox TxtFrecuencia 
      Height          =   285
      Left            =   4920
      TabIndex        =   7
      Top             =   1680
      Width           =   1695
   End
   Begin MSMask.MaskEdBox MaskEdfinal 
      Height          =   285
      Left            =   4920
      TabIndex        =   5
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Prestamo #:"
      Height          =   255
      Left            =   1440
      TabIndex        =   25
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Saldo Prestamos:"
      Height          =   255
      Left            =   3480
      TabIndex        =   23
      Top             =   2160
      Width           =   1455
   End
   Begin VB.OLE OLE1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Class           =   "Package"
      DisplayType     =   1  'Icon
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "FrmPrestamos.frx":222A
      SourceDoc       =   "C:\Icon\Smileys.exe"
      TabIndex        =   22
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre Empleado"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cuota a Pagar:"
      Height          =   375
      Left            =   3480
      TabIndex        =   20
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Inicio:"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frecuencia Cuota:"
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Cuota Final:"
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Monto Prestamo:"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Código Empleado:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "FrmPrestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAnterior_Click()
On Error GoTo TipoErrs
ValidaSalida ("en la Tabla Prestamos")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaPrestamos.Recordset.MovePrevious
       If DtaPrestamos.Recordset.BOF Then
           DtaPrestamos.Recordset.MoveNext
           MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Zeus Nóminas"
       Else
           DBCCodEmpleado.Text = DtaPrestamos.Recordset.CodEmpleado
       End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdBorrar_Click()
On Error GoTo TipoErrs
 Dim Respuesta, Rsp
CmdAnterior.Enabled = False
CmdSiguiente.Enabled = False
CmdPrimero.Enabled = False
CmdUltimo.Enabled = False
CmdBorrar.Enabled = False
Salida = False
'Elimino el registro activo en la pantalla
  Set Rsp = DtaPrestamos.Recordset
  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando El Prestamo de: " & TxtNombre.Text)
   If Respuesta = 6 Then
     Rsp.Delete
      Rsp.MovePrevious
      DBCCodEmpleado.Text = ""
      Sql = "SELECT Prestamos.Codprestamos, Prestamos.CodEmpleado, Prestamos.MontoPrestamo, Prestamos.FechaFinal, Prestamos.FrecuenciaCuota, Prestamos.FechaInicio, Prestamos.CuotaPagar, Prestamos.SaldoPrestamo From prestamos ORDER BY Prestamos.CodEmpleado"
      DtaPrestamos.RecordSource = Sql
      DtaPrestamos.Refresh
      DtaPrestamos.Recordset.MoveLast
      DtaPrestamos.Recordset.MovePrevious
      LimpiaPrestamos
   Salida = False
   End If
CmdAnterior.Enabled = True
CmdSiguiente.Enabled = True
CmdPrimero.Enabled = True
CmdUltimo.Enabled = True
CmdBorrar.Enabled = True
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub CmdGrabar_Click()
Dim Sql As String
CmdAnterior.Enabled = False
CmdSiguiente.Enabled = False
CmdPrimero.Enabled = False
CmdUltimo.Enabled = False
CmdBorrar.Enabled = False
On Error GoTo TipoErrs
  Salida = False
  
  
  If MaskEdMonto.Text = "" Then
   MsgBox "No se Puede Dejar Sin Monto", vbCritical, "Error:Zeus Nóminas"
   Exit Sub
  End If
  If MaskEdfinal.Text = "__/__/____" Then
   MsgBox "No se Puede Dejar Sin Meses", vbCritical, "Error:Zeus Nóminas"
   Exit Sub
  End If
  If TxtFrecuencia.Text = "" Then
   MsgBox "Se necesita frecuencia de Ahorro", vbCritical, "Error:Zeus Nóminas"
   Exit Sub
  End If
  If MaskEdFecha.Text = "__/__/____" Then
   MsgBox "Se necesita Fecha en que Inicia Ahorro", vbCritical, "Error:Zeus Nóminas"
   Exit Sub
  End If
  If TxtMontoPagar.Text = "" Then
   MsgBox "Se necesita el monto a pagar", vbCritical, "Error:Zeus Nóminas"
   Exit Sub
  End If
  'Busco si el codigo esta repetido si se repite solo se guarda la descripcion
      DtaPrestamos.Refresh
      Do While Not DtaPrestamos.Recordset.EOF
       If DtaPrestamos.Recordset.Codprestamos = TxtCodPrestamo.Text Then
         DtaPrestamos.Recordset.Edit
         GrabarPrestamos
         FrmPrestamos.DtaPrestamos.Recordset.Update
         Salida = False
         CmdAnterior.Enabled = True
        CmdSiguiente.Enabled = True
        CmdPrimero.Enabled = True
        CmdUltimo.Enabled = True
        CmdBorrar.Enabled = True
         Exit Sub
       End If
      DtaPrestamos.Recordset.MoveNext
      Loop
  'Si despues de Buscar no exite el codigo grabo todos los cambios
         
         DtaPrestamos.Recordset.AddNew
         GrabarPrestamos
         FrmPrestamos.DtaPrestamos.Recordset.Update
        Sql = "SELECT Prestamos.Codprestamos, Prestamos.CodEmpleado, Prestamos.MontoPrestamo, Prestamos.FechaFinal, Prestamos.FrecuenciaCuota, Prestamos.FechaInicio, Prestamos.CuotaPagar, Prestamos.SaldoPrestamo From prestamos ORDER BY Prestamos.CodEmpleado"
        DtaPrestamos.RecordSource = Sql
        DtaPrestamos.Refresh
        DtaPrestamos.Recordset.MoveLast
        DtaPrestamos.Recordset.MovePrevious
        'LimpiaPrestamos
        CmdAnterior.Enabled = True
        CmdSiguiente.Enabled = True
        CmdPrimero.Enabled = True
        CmdUltimo.Enabled = True
        CmdBorrar.Enabled = True
         Salida = False
         Exit Sub
TipoErrs:
  ControlErrores
  Unload Me
    
End Sub

Private Sub CmdPrimero_Click()
On Error GoTo TipoErrs
ValidaSalida ("en la Tabla Prestamos")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaPrestamos.Recordset.MoveFirst
       If DtaPrestamos.Recordset.BOF Then
           DtaPrestamos.Recordset.MoveNext
           MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Zeus Nóminas"
       Else
           DBCCodEmpleado.Text = DtaPrestamos.Recordset.CodEmpleado
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
 ValidaSalida ("en la Tabla Prestamos")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaPrestamos.Recordset.MoveNext
       If DtaPrestamos.Recordset.EOF Then
           DtaPrestamos.Recordset.MovePrevious
           MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Zeus Nóminas"
       Else
           DBCCodEmpleado.Text = DtaPrestamos.Recordset.CodEmpleado
       End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdUltimo_Click()
 On Error GoTo TipoErrs
 ValidaSalida ("en la Tabla Prestamos")
If Contesta Then
  CmdGrabar.Value = True
End If
 DtaPrestamos.Recordset.MoveLast
       If DtaPrestamos.Recordset.EOF Then
           DtaPrestamos.Recordset.MovePrevious
           MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Zeus Nóminas"
       Else
           DBCCodEmpleado.Text = DtaPrestamos.Recordset.CodEmpleado
       End If
Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub



Private Sub DBCCodEmpleado_Change()

Evaluar = True
  DtaEmpleado.Refresh
'Busco el codigo del empleado para que automaticamente ubique el nombre
 'aunque no existe en la data consulta
    Do While Not DtaEmpleado.Recordset.EOF
     If DtaEmpleado.Recordset.CodEmpleado = DBCCodEmpleado.Text Then
        TxtNombre.Text = DtaEmpleado.Recordset.Nombre1
        Exit Do
     Else
       FrmPrestamos.TxtNombre = ""
       FrmPrestamos.MaskEdMonto.Text = "0.00"
       FrmPrestamos.TxtSaldoPrestamo = "0.00"
       FrmPrestamos.TxtFrecuencia.Text = ""
       FrmPrestamos.MaskEdFecha.Text = "__/__/____"
       FrmPrestamos.TxtMontoPagar.Text = "0.00"
       FrmPrestamos.MaskEdfinal.Text = "__/__/____"
       TxtCodPrestamo.Text = ""
     End If
       DtaEmpleado.Recordset.MoveNext
   Loop

 DtaPrestamos.Refresh
   Do While Not DtaPrestamos.Recordset.EOF
     If DtaPrestamos.Recordset.CodEmpleado = DBCCodEmpleado.Text Then
        TxtCodPrestamo.Text = DtaPrestamos.Recordset.Codprestamos
        DBCCodEmpleado.Text = DtaPrestamos.Recordset.CodEmpleado
        MaskEdMonto.Text = Format((DtaPrestamos.Recordset.MontoPrestamo), "##,##0.00")
        TxtFrecuencia.Text = DtaPrestamos.Recordset.FrecuenciaCuota
        MaskEdFecha.Text = DtaPrestamos.Recordset.FechaInicio
        MaskEdfinal.Text = DtaPrestamos.Recordset.FechaFinal
        TxtMontoPagar.Text = Format((DtaPrestamos.Recordset.CuotaPagar), "##,##0.00")
        TxtSaldoPrestamo.Text = Format((DtaPrestamos.Recordset.SaldoPrestamo), "##,##0.00")
        MontoPagar = TxtMontoPagar.Text
        SaldoPrestamo = TxtSaldoPrestamo.Text
        Monto = MaskEdMonto.Text
        TxtFrecuencia.Text = TxtFrecuencia.Text & " " & "Dias"
        TxtSaldoPrestamo.Text = "C$" & " " & TxtSaldoPrestamo.Text
        TxtMontoPagar.Text = "C$" & " " & TxtMontoPagar.Text
        MaskEdMonto.Text = "C$" & " " & MaskEdMonto.Text
        Exit Do
      End If
       DtaPrestamos.Recordset.MoveNext
   Loop

Salida = False

End Sub

Private Sub DBCCodEmpleado_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   MaskEdMonto.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub DBCCodPrestamo_Change()

End Sub

Private Sub DBCCodPrestamo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   DBCCodEmpleado.SetFocus
 End If
End Sub


Private Sub Form_Activate()
Dim Sql As String
DBCCodEmpleado.SetFocus
Salida = False
Sql = "SELECT Prestamos.Codprestamos, Prestamos.CodEmpleado, Prestamos.MontoPrestamo, Prestamos.FechaFinal, Prestamos.FrecuenciaCuota, Prestamos.FechaInicio, Prestamos.CuotaPagar, Prestamos.SaldoPrestamo From prestamos ORDER BY Prestamos.CodEmpleado"
DtaPrestamos.RecordSource = Sql
DtaPrestamos.Refresh
 If Not BPrestamos = True Then
   CmdBorrar.Enabled = False
 End If
 If Not GPrestamos = True Then
   CmdGrabar.Enabled = False
 End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If GPrestamos = True Then
ValidaSalida ("en la Tabla Prestamos")
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

Private Sub MaskEdFecha_Change()
PreparaSalida
End Sub

Private Sub MaskEdFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
   MaskEdfinal.SetFocus
  Else
   Evaluar = False
  End If
End Sub

Private Sub MaskEdfinal_Change()
PreparaSalida
End Sub

Private Sub MaskEdfinal_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   TxtMontoPagar.SetFocus
  Else
   Evaluar = False
  End If
End Sub

Private Sub MaskEdMonto_Change()
PreparaSalida
End Sub

Private Sub MaskEdMonto_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   MaskEdFecha.SetFocus
  Else
   Evaluar = False
  End If
End Sub

Private Sub MaskEdMonto_LostFocus()
 MaskEdMonto.Text = Format((MaskEdMonto.Text), "##,##0.00")
 Monto = MaskEdMonto.Text
 MaskEdMonto.Text = "C$" & " " & MaskEdMonto.Text
End Sub

Private Sub TxtFrecuencia_Change()
PreparaSalida
End Sub

Private Sub TxtFrecuencia_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  CmdGrabar.SetFocus
  Else
   Evaluar = False
  End If
End Sub

Private Sub TxtMeses_Change()
PreparaSalida
End Sub

Private Sub TxtMeses_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  TxtFrecuencia.SetFocus
  Else
   Evaluar = False
  End If
End Sub

Private Sub TxtFrecuencia_LostFocus()
TxtFrecuencia.Text = TxtFrecuencia.Text & " " & "Dias"
End Sub

Private Sub TxtMontoPagar_Change()
PreparaSalida
End Sub

Private Sub TxtMontoPagar_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  TxtFrecuencia.SetFocus
 Else
   Evaluar = False
  End If
End Sub

Private Sub TxtMontoPagar_LostFocus()
 TxtMontoPagar.Text = Format((TxtMontoPagar.Text), "##,##0.00")
 MontoPagar = TxtMontoPagar.Text
 TxtMontoPagar.Text = "C$" & " " & TxtMontoPagar.Text
End Sub

Private Sub TxtNombre_Change()
Salida = False
End Sub
