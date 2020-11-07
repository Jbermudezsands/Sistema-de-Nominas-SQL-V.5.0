VERSION 5.00
Begin VB.Form FrmExporta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportacion de Registros"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6465
   HelpContextID   =   41
   Icon            =   "FrmExporta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   431
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
      Left            =   5640
      Picture         =   "FrmExporta.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   375
   End
   Begin VB.CheckBox ChkGrupo 
      Caption         =   "Agregar Departamento"
      Height          =   195
      Left            =   3480
      TabIndex        =   7
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton CmdProcesar 
      DownPicture     =   "FrmExporta.frx":0458
      Height          =   375
      Left            =   2640
      MouseIcon       =   "FrmExporta.frx":1F3A
      MousePointer    =   99  'Custom
      Picture         =   "FrmExporta.frx":237C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton CmdCerrar 
      DownPicture     =   "FrmExporta.frx":3E5E
      Height          =   375
      Left            =   4080
      MouseIcon       =   "FrmExporta.frx":5940
      MousePointer    =   99  'Custom
      Picture         =   "FrmExporta.frx":5D82
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox TxtRuta 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de Exporacion"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      Begin VB.OptionButton OptPrestamo 
         Caption         =   "Movimientos de Préstamo"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2415
      End
      Begin VB.OptionButton OptTransaciones 
         Caption         =   "Transaciones de la Nomina."
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton OptCuentas 
         Caption         =   "Cuentas y Contactos"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.Data DtaExporta 
      Caption         =   "DtaExporta"
      Connect         =   "Access"
      DatabaseName    =   "C:\Zeus Nominas\Nominas.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DetalleNomina"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data DtaHistorico 
      Caption         =   "DtaHistorico"
      Connect         =   "Access"
      DatabaseName    =   "C:\Zeus Nominas\Nominas.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Historico"
      Top             =   4440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data DtaEmpleado 
      Caption         =   "DtaEmpleado"
      Connect         =   "Access"
      DatabaseName    =   "C:\Zeus Nominas\Nominas.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Empleado"
      Top             =   3720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label LblTitulo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Ruta Exportacion:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
   End
End
Attribute VB_Name = "FrmExporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Sub CmdBuscarEmpleado_Click()
'QuienLlama = "Exporta Nomina"
'FrmRuta.Caption = "Buscando Ruta de Base de Datos"
'FrmRuta.Show 1
'End Sub
'
'Private Sub CmdCerrar_Click()
' Unload Me
'End Sub
'
'Private Sub CmdConsulta_Click(Index As Integer)
' FrmRuta.Show 1
'End Sub
'
'Private Sub CmdProcesar_Click()
'On Error GoTo TipoErrs
'Dim SQLExporta As String
'Dim Cadena As String
'Dim TextoMonto
'
'If OptCuentas Then
'    Ruta2 = TxtRuta.Text
'    If (Dir(Ruta2) <> "") Then
'       R% = MsgBox("Reescribir el Archivo?", vbYesNo, "Sistema de Nominas")
'        If R% = 6 Then
'          Open Ruta2 For Output As #1
'          CreaArchivo
'          Close #1
'        Else
'          Exit Sub
'        End If
'    Else
'         Open Ruta2 For Output As #1
'           CreaArchivo
'          Close #1
'    End If
'End If
'If OptTransaciones Then
'                Ruta2 = TxtRuta.Text
'                Open Ruta2 For Output As #1
'                SQLExporta = "SELECT Empleado.CodEmpleado, Empleado.CodDepartamento, Historico.CuentaDebito, Historico.CuentaCredito, DetalleNomina.NumNomina, Nomina.FechaNomina, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[HorasExtras]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]-[DetalleNomina].[Deducciones]-[DetalleNomina].[Prestamo]-[DetalleNomina].[MontoINSS]-[DetalleNomina].[MontoIR]+[DetalleNomina].[TotalSubsidio] AS GranTotal FROM Nomina INNER JOIN ((Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado) ON Nomina.NumNomina = DetalleNomina.NumNomina Where DetalleNomina.NumNomina = " & NumNomina & " ORDER BY Empleado.CodEmpleado"
'                DtaExporta.RecordSource = SQLExporta
'                DtaExporta.Refresh
'                 Do While Not DtaExporta.Recordset.EOF
'                    TextoMonto = Format(DtaExporta.Recordset.grantotal, "####0.00")
'                    For I = 1 To 15 - Len(TextoMonto)
'                       TextoMonto = " " + TextoMonto
'                    Next I
'
'                    Cadena = Trim(Str(Month(DtaExporta.Recordset.FechaNomina)))
'                    Cadena = Cadena + Trim(Str(Day(DtaExporta.Recordset.FechaNomina)))
'                    Cadena = Cadena + Trim(Str(Year(DtaExporta.Recordset.FechaNomina)))
'                    Cadena = Cadena + "        "
'                    Cadena = Cadena + "ZEUS"
'                    Cadena = Cadena + Trim(Str(DtaExporta.Recordset.CuentaDebito))
'                    For I = 1 To 36 - Len(Cadena)
'                    Cadena = Cadena + " "
'                    Next I
'                    If ChkGrupo.Value = 1 Then
'                        Cadena = Cadena + Trim(DtaExporta.Recordset.CodDepartamento)
'                    Else
'                        Cadena = Cadena + "  "
'                    End If
'                    Cadena = Cadena + "                "
'                    Cadena = Cadena + "               "
'                    Cadena = Cadena + "          "
'                    Cadena = Cadena + "03"
'                    Cadena = Cadena + "      "
'                    Cadena = Cadena + "Pago de la Nómina " + Trim(Str(NumNomina)) + "               "
'                    Cadena = Cadena + "                "
'                    Cadena = Cadena + "   " + TextoMonto
'                    For I = 1 To 34
'                        Cadena = Cadena + " "
'                    Next I
'                    Cadena = Cadena + "00"
'                    Print #1, Cadena
'
'                    Cadena = Trim(Str(Month(DtaExporta.Recordset.FechaNomina)))
'                    Cadena = Cadena + Trim(Str(Day(DtaExporta.Recordset.FechaNomina)))
'                    Cadena = Cadena + Trim(Str(Year(DtaExporta.Recordset.FechaNomina)))
'                    Cadena = Cadena + "        "
'                    Cadena = Cadena + "ZEUS"
'                    Cadena = Cadena + Trim(Str(DtaExporta.Recordset.cuentacredito))
'                    For I = 1 To 36 - Len(Cadena)
'                    Cadena = Cadena + " "
'                    Next I
'                    If ChkGrupo.Value = 1 Then
'                        Cadena = Cadena + Trim(DtaExporta.Recordset.CodDepartamento)
'                    Else
'                        Cadena = Cadena + "  "
'                    End If
'                    Cadena = Cadena + "                "
'                    Cadena = Cadena + "               "
'                    Cadena = Cadena + "          "
'                    Cadena = Cadena + "07"
'                    Cadena = Cadena + "      "
'                    Cadena = Cadena + "Pago de la Nómina " + Trim(Str(NumNomina)) + "               "
'                    Cadena = Cadena + "                "
'                    Cadena = Cadena + "   " + TextoMonto
'                        For I = 1 To 34
'                        Cadena = Cadena + " "
'                    Next I
'                    Cadena = Cadena + "00"
'                    Print #1, Cadena
'
'                  DtaExporta.Recordset.MoveNext
'                  Loop
'                 Close #1
'End If
'If OptPrestamo Then
'
'
'                Ruta2 = TxtRuta.Text
'                Open Ruta2 For Output As #1
'                SQLExporta = "SELECT Prestamo.NumPrestamo, Prestamo.CodEmpleado, Prestamo.CuentaDebito, Prestamo.CuentaCredito, MovPrestamo.Monto, MovPrestamo.NumCuota, MovPrestamo.CuotaIgual, MovPrestamo.NumNomina, Nomina.FechaNomina FROM (Prestamo INNER JOIN MovPrestamo ON Prestamo.NumPrestamo = MovPrestamo.NumPrestamo) INNER JOIN Nomina ON MovPrestamo.NumNomina = Nomina.NumNomina WHERE MovPrestamo.NumNomina= " & NumNomina & ""
'                DtaExporta.RecordSource = SQLExporta
'                DtaExporta.Refresh
'                 Do While Not DtaExporta.Recordset.EOF
'                    TextoMonto = Format(DtaExporta.Recordset.CuotaIgual, "####0.00")
'                    For I = 1 To 15 - Len(TextoMonto)
'                       TextoMonto = " " + TextoMonto
'                    Next I
'
'                    Cadena = Trim(Str(Month(DtaExporta.Recordset.FechaNomina)))
'                    Cadena = Cadena + Trim(Str(Day(DtaExporta.Recordset.FechaNomina)))
'                    Cadena = Cadena + Trim(Str(Year(DtaExporta.Recordset.FechaNomina)))
'                    Cadena = Cadena + "        "
'                    Cadena = Cadena + "ZEUS"
'                    Cadena = Cadena + Trim(Str(DtaExporta.Recordset.CuentaDebito))
'                    For I = 1 To 36 - Len(Cadena)
'                    Cadena = Cadena + " "
'                    Next I
'                    Cadena = Cadena + "  "
'                    Cadena = Cadena + "                "
'                    Cadena = Cadena + "               "
'                    Cadena = Cadena + "          "
'                    Cadena = Cadena + "03"
'                    Cadena = Cadena + "      "
'                    Cadena = Cadena + "Pago de la Prestamo " + Trim(Str(DtaExporta.Recordset.NumPrestamo)) + " Cuota:"
'                    Cadena = Cadena + Str(DtaExporta.Recordset.numcuota)
'
'                    If Len(Trim(Str(DtaExporta.Recordset.numcuota))) > 1 Then
'                        Cadena = Cadena + "                   "
'                    Else
'                        Cadena = Cadena + "                    "
'                    End If
'
'                    Cadena = Cadena + "   " + TextoMonto
'                    For I = 1 To 34
'                        Cadena = Cadena + " "
'                    Next I
'                    Cadena = Cadena + "00"
'                    Print #1, Cadena
'
'                    TextoMonto = Format(DtaExporta.Recordset.CuotaIgual, "####0.00")
'                    For I = 1 To 15 - Len(TextoMonto)
'                       TextoMonto = " " + TextoMonto
'                    Next I
'
'                    Cadena = Trim(Str(Month(DtaExporta.Recordset.FechaNomina)))
'                    Cadena = Cadena + Trim(Str(Day(DtaExporta.Recordset.FechaNomina)))
'                    Cadena = Cadena + Trim(Str(Year(DtaExporta.Recordset.FechaNomina)))
'                    Cadena = Cadena + "        "
'                    Cadena = Cadena + "ZEUS"
'                    Cadena = Cadena + Trim(Str(DtaExporta.Recordset.cuentacredito))
'                    For I = 1 To 36 - Len(Cadena)
'                    Cadena = Cadena + " "
'                    Next I
'                    Cadena = Cadena + "  "
'                    Cadena = Cadena + "                "
'                    Cadena = Cadena + "               "
'                    Cadena = Cadena + "          "
'                    Cadena = Cadena + "07"
'                    Cadena = Cadena + "      "
'                    Cadena = Cadena + "Pago de la Prestamo " + Trim(Str(DtaExporta.Recordset.NumPrestamo)) + " Cuota:"
'                    Cadena = Cadena + Str(DtaExporta.Recordset.numcuota)
'                    If Len(Trim(Str(DtaExporta.Recordset.numcuota))) > 1 Then
'                        Cadena = Cadena + "                   "
'                    Else
'                        Cadena = Cadena + "                    "
'                    End If
'                    Cadena = Cadena + "   " + TextoMonto
'                    For I = 1 To 34
'                        Cadena = Cadena + " "
'                    Next I
'                    Cadena = Cadena + "00"
'                    Print #1, Cadena
'
'                  DtaExporta.Recordset.MoveNext
'                  Loop
'                 Close #1
'End If
'Unload Me
'Exit Sub
'TipoErrs:
'ControlErrores
' TxtRuta.Text = ""
' End Sub
'
'Private Sub Form_Load()
'
'With Me.DtaEmpleado
'   '.DatabaseName = Ruta
'   .ConnectionString = Conexion
'End With
'
'With Me.DtaExporta
'   '.DatabaseName = Ruta
'   .ConnectionString = Conexion
'End With
'
'With Me.DtaHistorico
'   '.DatabaseName = Ruta
'   .ConnectionString = Conexion
'End With
'
'End Sub
'
'Private Sub xptopbuttons1_Click()
'Unload Me
'End Sub
Private Sub xp_canvas1_Click()

End Sub
