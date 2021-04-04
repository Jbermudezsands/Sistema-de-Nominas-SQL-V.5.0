VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmReportesNominas 
   Caption         =   "Reportes Varios"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7350
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmdlgReportes 
      Left            =   480
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data dtaServidor 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc adoTipoNomina 
      Height          =   375
      Left            =   6600
      Top             =   1920
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
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
      Caption         =   "Tipo de Nomina"
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
   Begin VB.CommandButton cmdReporte 
      Caption         =   "Reporte"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CheckBox chkExportar 
         Caption         =   "Exportar a Excel"
         Height          =   255
         Left            =   3720
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox cboTipoNomina 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cboPeriodo 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmReportesNominas.frx":0000
         Left            =   2520
         List            =   "frmReportesNominas.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtAno 
         Height          =   285
         Left            =   960
         MaxLength       =   5
         TabIndex        =   2
         Text            =   " "
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nomina:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "Periodo"
         Height          =   255
         Left            =   4560
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.Label txtMes 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   2040
         TabIndex        =   7
         Top             =   720
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmReportesNominas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ConexionRep As String






Private Sub cboMes_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = 13 Then
   
   
   Me.adoTipoNomina.ConnectionString = ConexionRep
   Me.adoTipoNomina.CommandType = adCmdText
   Me.adoTipoNomina.RecordSource = "SELECT TipoNomina.Nomina, Nomina.Mes, Nomina.Ano, Nomina.Periodo, Nomina.Activa " & _
                                   "FROM Nomina INNER JOIN TipoNomina ON Nomina.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                                   "WHERE (Nomina.Mes =" & Trim(Mid$(Me.cboMes.Text, 1, InStr(1, Me.cboMes.Text, " ", vbTextCompare))) & " AND Nomina.Ano =" & Me.txtAno.Text & " AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "') ORDER BY Nomina.Periodo ASC"
                                   
   Me.adoTipoNomina.Refresh
 
   Me.cboPeriodo.Clear
   Me.cboPeriodo.AddItem "Todos"
   
     
   Do While Not Me.adoTipoNomina.Recordset.EOF

      Me.cboPeriodo.AddItem Me.adoTipoNomina.Recordset.Fields("Periodo")
      Me.adoTipoNomina.Recordset.MoveNext

   Loop
   
   Me.cboPeriodo.SetFocus
   
End If


End Sub





Private Sub cboPeriodo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
  
  Me.cmdReporte.SetFocus

End If


End Sub

Private Sub cboTipoNomina_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   
   Me.txtAno.SetFocus

End If


End Sub

Private Sub cmdReporte_Click()

Dim sSQl As String
Dim rptTotalNominasxDepto As New arepTotalNominasxDepto


If Me.cboMes.Text <> "" And Me.cboTipoNomina.Text <> "" And IsNumeric(Me.txtAno.Text) Then
   
   If Me.chkExportar.Value Then
      Me.cmdlgReportes.ShowSave
      Directorio = ""
      Directorio = Me.cmdlgReportes.FileName + ".xls"
      bExportar = True
   End If
   
   
   
   If Me.cboPeriodo.Text <> "Todos" Then

      sSQl = "SELECT dbo.TipoNomina.Nomina, dbo.TipoNomina.CodTipoNomina, dbo.Nomina.NumNomina, dbo.Nomina.Activa, dbo.Nomina.FechaNominaINI, " & _
                      "dbo.Nomina.FechaNomina, dbo.Nomina.Mes, dbo.Nomina.Ano, dbo.Nomina.Periodo, dbo.Empleado.CodEmpleado1, dbo.Empleado.Nombre1, " & _
                      "dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2, dbo.Empleado.Sexo, dbo.Departamento.Departamento, dbo.Cargo.Cargo, " & _
                      "dbo.DetalleNomina.SalarioBasico, dbo.DetalleNomina.HE, dbo.DetalleNomina.DD, dbo.DetalleNomina.HorasExtras, dbo.DetalleNomina.OtrosIngresos, " & _
                      "dbo.DetalleNomina.Incentivos, dbo.DetalleNomina.Deducciones, dbo.DetalleNomina.MontoINSS, dbo.DetalleNomina.MontoIR, " & _
                      "dbo.DetalleNomina.INSSPatronal, dbo.DetalleNomina.IRPatronal, dbo.DetalleNomina.INATEC, dbo.DetalleNomina.HTrabajada, " & _
                      "dbo.DetalleNomina.SeptimoDia, dbo.DetalleNomina.IncetivoProduccion, dbo.DetalleNomina.TarifaHoraria, dbo.DetalleNomina.Destajo, " & _
                      "dbo.Nomina.TotalHorasExtras, dbo.Nomina.TotalIncentivos, dbo.Nomina.TotalDeducciones, dbo.Nomina.TotalOtrosIngresos, " & _
                      "dbo.Nomina.TotalMontoINSS, dbo.Nomina.TotalMontoIR, dbo.Nomina.TotalINSSPatronal, dbo.Nomina.TotalIRPatronal, dbo.Nomina.TotalINATEC, " & _
                      "dbo.Empleado.Activo , dbo.Nomina.TotalDestajo, dbo.Nomina.TotalSalarioBasico " & _
            "FROM dbo.Cargo INNER JOIN " & _
                      "dbo.Empleado ON dbo.Cargo.CodCargo = dbo.Empleado.CodCargo INNER JOIN " & _
                      "dbo.Departamento ON dbo.Empleado.CodDepartamento = dbo.Departamento.CodDepartamento INNER JOIN " & _
                      "dbo.DetalleNomina ON dbo.Empleado.CodEmpleado = dbo.DetalleNomina.CodEmpleado INNER JOIN " & _
                      "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina INNER JOIN " & _
                      "dbo.TipoNomina ON dbo.Nomina.CodTipoNomina = dbo.TipoNomina.CodTipoNomina " & _
            "WHERE (dbo.TipoNomina.Nomina = '" & Me.cboTipoNomina.Text & "') AND (dbo.Nomina.Periodo = " & Me.cboPeriodo.Text & ") AND dbo.Nomina.Mes =" & Trim(Mid$(Me.cboMes.Text, 1, InStr(1, Me.cboMes.Text, " ", vbTextCompare))) & " AND " & _
                "(dbo.DetalleNomina.MontoINSS <> 0) ORDER BY dbo.Departamento.Departamento ASC"

   
   Else
   
   
        sSQl = "SELECT dbo.TipoNomina.Nomina, dbo.TipoNomina.CodTipoNomina, dbo.Nomina.NumNomina, dbo.Nomina.Activa, dbo.Nomina.FechaNominaINI, " & _
                      "dbo.Nomina.FechaNomina, dbo.Nomina.Mes, dbo.Nomina.Ano, dbo.Nomina.Periodo, dbo.Empleado.CodEmpleado1, dbo.Empleado.Nombre1, " & _
                      "dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2, dbo.Empleado.Sexo, dbo.Departamento.Departamento, dbo.Cargo.Cargo, " & _
                      "dbo.DetalleNomina.SalarioBasico, dbo.DetalleNomina.HE, dbo.DetalleNomina.DD, dbo.DetalleNomina.HorasExtras, dbo.DetalleNomina.OtrosIngresos, " & _
                      "dbo.DetalleNomina.Incentivos, dbo.DetalleNomina.Deducciones, dbo.DetalleNomina.MontoINSS, dbo.DetalleNomina.MontoIR, " & _
                      "dbo.DetalleNomina.INSSPatronal, dbo.DetalleNomina.IRPatronal, dbo.DetalleNomina.INATEC, dbo.DetalleNomina.HTrabajada, " & _
                      "dbo.DetalleNomina.SeptimoDia, dbo.DetalleNomina.IncetivoProduccion, dbo.DetalleNomina.TarifaHoraria, dbo.DetalleNomina.Destajo, " & _
                      "dbo.Nomina.TotalHorasExtras, dbo.Nomina.TotalIncentivos, dbo.Nomina.TotalDeducciones, dbo.Nomina.TotalOtrosIngresos, " & _
                      "dbo.Nomina.TotalMontoINSS, dbo.Nomina.TotalMontoIR, dbo.Nomina.TotalINSSPatronal, dbo.Nomina.TotalIRPatronal, dbo.Nomina.TotalINATEC, " & _
                      "dbo.Empleado.Activo , dbo.Nomina.TotalDestajo, dbo.Nomina.TotalSalarioBasico " & _
            "FROM dbo.Cargo INNER JOIN " & _
                      "dbo.Empleado ON dbo.Cargo.CodCargo = dbo.Empleado.CodCargo INNER JOIN " & _
                      "dbo.Departamento ON dbo.Empleado.CodDepartamento = dbo.Departamento.CodDepartamento INNER JOIN " & _
                      "dbo.DetalleNomina ON dbo.Empleado.CodEmpleado = dbo.DetalleNomina.CodEmpleado INNER JOIN " & _
                      "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina INNER JOIN " & _
                      "dbo.TipoNomina ON dbo.Nomina.CodTipoNomina = dbo.TipoNomina.CodTipoNomina " & _
            "WHERE (dbo.Empleado.Activo = 1) AND (dbo.TipoNomina.Nomina = '" & Me.cboTipoNomina.Text & "') AND dbo.Nomina.Mes =" & Trim(Mid$(Me.cboMes.Text, 1, InStr(1, Me.cboMes.Text, " ", vbTextCompare))) & " AND " & _
                "(dbo.DetalleNomina.MontoINSS <> 0) ORDER BY dbo.Departamento.Departamento ASC"
   
   End If

'Trim (Mid$(Me.cboMes.Text, InStr(1, Me.cboMes.Text, " ", vbTextCompare), Len(Me.cboMes.Text)))


  rptTotalNominasxDepto.DataControl1.ConnectionString = ConexionRep
  rptTotalNominasxDepto.DataControl1.Source = sSQl
  rptTotalNominasxDepto.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & "; Año: " & Me.txtAno.Text & "; Mes: " & Trim(Mid$(Me.cboMes.Text, InStr(1, Me.cboMes.Text, " ", vbTextCompare), Len(Me.cboMes.Text))) & "; Periodo: " & Me.cboPeriodo.Text
  'rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
  rptTotalNominasxDepto.Show 1

End If


End Sub

Private Sub cmdSalir_Click()

Unload Me

End Sub

Private Sub Form_Load()
Dim RutaServer As String
 Dim Server As String
 'Dim Conexion As String


RutaServer = App.Path + "\CntNominas.dll"

  With Me.dtaServidor
     .DatabaseName = RutaServer
     .RecordSource = "Servidor"
     .Refresh
  End With


  If Not IsNull(Me.dtaServidor.Recordset.Servidor) Then
   Server = Me.dtaServidor.Recordset.Servidor
  Else
   MsgBox "No se ha definido el Servidor", vbCritical, "Sistmea de Nominas"
   Exit Sub
  End If



'Borrar despues la siguiente linea

'Server = "Modell"

ConexionRep = "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=" & Server


Me.adoTipoNomina.ConnectionString = ConexionRep
Me.adoTipoNomina.CommandType = adCmdText
Me.adoTipoNomina.RecordSource = "SELECT CodTipoNomina, Nomina, Periodo FROM TipoNomina"
Me.adoTipoNomina.Refresh

Do While Not Me.adoTipoNomina.Recordset.EOF

   Me.cboTipoNomina.AddItem Me.adoTipoNomina.Recordset.Fields("Nomina")
   Me.adoTipoNomina.Recordset.MoveNext

Loop










End Sub


Private Sub txtAno_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = 13 Then
  
  Me.cboMes.SetFocus
  

End If



End Sub
