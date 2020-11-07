VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmExtraTurno 
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   9195
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data dtaServidor 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Borrar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4800
      TabIndex        =   29
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   6960
      TabIndex        =   26
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   16
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   7080
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc adoExtraTurno 
      Height          =   330
      Left            =   9000
      Top             =   7200
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
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
      Caption         =   "Extra Turno"
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
   Begin MSAdodcLib.Adodc adoEmpleado 
      Height          =   330
      Left            =   6000
      Top             =   7560
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
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
      Caption         =   "Empleado"
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
   Begin MSAdodcLib.Adodc adoDepto 
      Height          =   330
      Left            =   7200
      Top             =   7200
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   582
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
      Caption         =   "Depto"
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
   Begin VB.Frame fraElegidos 
      Caption         =   "Seleccion del Personal"
      Height          =   3615
      Left            =   120
      TabIndex        =   24
      Top             =   3120
      Width           =   8775
      Begin VB.ListBox lstEmpleados 
         Height          =   1860
         Left            =   480
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   960
         Width           =   3255
      End
      Begin VB.CommandButton cmdDepto 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "<"
         Height          =   375
         Left            =   3960
         TabIndex        =   14
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdJalar 
         Caption         =   ">"
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   1440
         Width           =   615
      End
      Begin VB.ListBox lstElegidos 
         Height          =   2085
         Left            =   4800
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   960
         Width           =   3735
      End
      Begin VB.ComboBox cboDepto 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblFecha 
         Caption         =   " "
         Height          =   255
         Left            =   5400
         TabIndex        =   28
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   4800
         TabIndex        =   27
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Depto:"
         Height          =   195
         Left            =   360
         TabIndex        =   25
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   "Busqueda por Empleado ó Fecha"
      Height          =   1575
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Reporte"
         Height          =   495
         Left            =   4440
         TabIndex        =   30
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cboCodigo 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Text            =   " "
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   4440
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   23134209
         CurrentDate     =   38570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   3120
         TabIndex        =   22
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblNombre 
         Caption         =   " "
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   1080
         Width           =   5295
      End
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "Datos de Entrada - Salida"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   8775
      Begin VB.TextBox txtTiempoReceso 
         Height          =   285
         Left            =   4800
         MaxLength       =   3
         TabIndex        =   9
         Text            =   " "
         Top             =   960
         Width           =   615
      End
      Begin MSMask.MaskEdBox mskAsistHoraEntrada 
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm AM/PM"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpFechEntrada 
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   23134209
         CurrentDate     =   38570
      End
      Begin MSComCtl2.DTPicker dtpFecSalida 
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   23134209
         CurrentDate     =   38570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tiempo de Receso (min.)"
         Height          =   195
         Left            =   2880
         TabIndex        =   19
         Top             =   960
         Width           =   1770
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Entrada"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hora Entrada"
         Height          =   195
         Left            =   2880
         TabIndex        =   17
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Salida"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmExtraTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Conexion As String

Public Function Chequear_Hora(Hora As String) As Boolean

Dim sMinutos As String
Dim sHora As String
Dim sSegundos As String

sHora = Mid$(Hora, 1, 2)
sMinutos = Mid$(Hora, 4, 2)
sSegundos = Mid$(Hora, 7, 2)


If (IsNumeric(sHora) And CInt(sHora) <= 23) And (IsNumeric(sMinutos) And CInt(sMinutos) <= 59) And (IsNumeric(sSegundos) And CInt(sMinutos) <= 59) Then

   Chequear_Hora = True

Else
   
   Chequear_Hora = False

End If


End Function




Private Sub cboCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
  
  Me.dtpFecha.SetFocus

End If

End Sub



Private Sub cboDepto_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   
   Me.cmdDepto.SetFocus


End If

End Sub

Private Sub cmdAgregar_Click()

Dim sFechaEntrada As String
Dim iCont As Integer

sFechaEntrada = Mid$(Me.dtpFechEntrada.Value, 7, 4) & "-" & Mid$(Me.dtpFechEntrada.Value, 4, 2) & "-" & Mid$(Me.dtpFechEntrada.Value, 1, 2)


If Me.cmdAgregar.Caption = "&Agregar" Then
   Me.fraBusqueda.Enabled = False
   Me.fraGeneral.Enabled = True
  
   Me.dtpFechEntrada.SetFocus
   Me.cmdAgregar.Caption = "&Guardar"
   Me.cmdModificar.Enabled = False
   Me.cmdBorrar.Enabled = False
   
 ElseIf Me.mskAsistHoraEntrada.Text <> "__:__:__" Then

 If Chequear_Hora(Me.mskAsistHoraEntrada.Text) And IsNumeric(Me.txtTiempoReceso.Text) Then

  Me.adoEmpleado.Refresh

  If Me.lstElegidos.ListIndex >= 0 Then

   Do While iCont <= Me.lstElegidos.ListCount
      
     Me.adoExtraTurno.CommandType = adCmdText
     Me.adoExtraTurno.RecordSource = "SELECT * FROM ExtraTurno WHERE CodEmpleado ='" & Mid$(Me.lstElegidos.List(iCont), 1, 6) & "' AND FechaEntrada = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102)"
     Me.adoExtraTurno.Refresh
     
     If Not Me.adoExtraTurno.Recordset.EOF Then
        Me.lstElegidos.RemoveItem (iCont)
        iCont = iCont - 1
     End If
   
     iCont = iCont + 1
     
    Loop
  
    Me.adoExtraTurno.Refresh
    iCont = 0
  
    Do While iCont <= Me.lstElegidos.ListCount - 1
     
      Me.adoExtraTurno.Recordset.AddNew
      Me.adoExtraTurno.Recordset.Fields("CodEmpleado") = Mid$(Me.lstElegidos.List(iCont), 1, 6)
      Me.adoEmpleado.Recordset.Find "[CodEmpleado] ='" & Mid$(Me.lstElegidos.List(iCont), 1, 6) & "'"
      Me.adoExtraTurno.Recordset.Fields("CodTipoNomina") = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
      Me.adoExtraTurno.Recordset.Fields("CodDepartamento") = Me.adoEmpleado.Recordset.Fields("CodDepartamento")
      Me.adoExtraTurno.Recordset.Fields("FechaEntrada") = Me.dtpFechEntrada.Value
      Me.adoExtraTurno.Recordset.Fields("HoraEntrada") = Me.mskAsistHoraEntrada.Text
      Me.adoExtraTurno.Recordset.Fields("TiempoReceso") = CInt(Me.txtTiempoReceso.Text)
      Me.adoExtraTurno.Recordset.Fields("FechaSalida") = Me.dtpFecSalida.Value
      Me.adoExtraTurno.Recordset.Update
      Me.adoExtraTurno.Refresh
      Me.adoEmpleado.Refresh
      iCont = iCont + 1
    
    Loop
  
    Me.fraBusqueda.Enabled = True
    Me.fraGeneral.Enabled = False
  
    Me.cboCodigo.SetFocus
    Me.cmdAgregar.Caption = "&Agregar"
    Me.cmdModificar.Enabled = True
    Me.cmdBorrar.Enabled = True
    
 End If
End If

   




End If

End Sub

Private Sub cmdBorrar_Click()

Dim sFechaEntrada As String
Dim iCont As Integer


If MsgBox("¿Realmente desea borrar a los empleados seleccionados?", vbInformation + vbYesNo, "Borrar") = vbYes And Me.lstElegidos.ListCount > 0 Then
    
   sFechaEntrada = Mid$(Me.lblFecha.Caption, 7, 4) & "-" & Mid$(Me.lblFecha.Caption, 4, 2) & "-" & Mid$(Me.lblFecha.Caption, 1, 2)
     
       Do While iCont < Me.lstElegidos.ListCount
           
         If Me.lstElegidos.Selected(iCont) Then
              Me.adoExtraTurno.CommandType = adCmdText
              Me.adoExtraTurno.RecordSource = "SELECT * FROM ExtraTurno WHERE CodEmpleado ='" & Mid$(Me.lstElegidos.List(iCont), 1, 6) & "' AND FechaEntrada = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102)"
              Me.adoExtraTurno.Refresh
     
             If Not Me.adoExtraTurno.Recordset.EOF Then
                Me.adoExtraTurno.Recordset.Delete
                Me.adoExtraTurno.Recordset.Update
                Me.lstElegidos.RemoveItem (iCont)
                iCont = iCont - 1
                
             End If
             
           End If
           
           iCont = iCont + 1
     
        Loop


End If



End Sub

Private Sub cmdBuscar_Click()


Dim sFechaEntrada As String
Dim bEncontrado As Boolean


sFechaEntrada = Mid$(Me.dtpFecha.Value, 7, 4) & "-" & Mid$(dtpFecha.Value, 4, 2) & "-" & Mid$(dtpFecha.Value, 1, 2)

If Trim(Me.cboCodigo.Text) <> "" Then
   
   Me.adoEmpleado.Recordset.Find "CodEmpleado ='" & Trim(Me.cboCodigo.Text) & "'"
    
   If Not Me.adoEmpleado.Recordset.EOF Then
      
     Me.adoExtraTurno.CommandType = adCmdText
     Me.adoExtraTurno.RecordSource = "SELECT * FROM ExtraTurno WHERE CodEmpleado ='" & Me.adoEmpleado.Recordset.Fields("CodEmpleado") & "' AND FechaEntrada = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102)"
     Me.adoExtraTurno.Refresh
      
      
     If Not Me.adoExtraTurno.Recordset.EOF Then
        Me.lstElegidos.Clear
        Me.cmdModificar.Enabled = True
        Me.lstElegidos.AddItem Me.adoExtraTurno.Recordset.Fields("CodEmpleado") & " " & Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
        
        Me.txtTiempoReceso.Text = Me.adoExtraTurno.Recordset.Fields("TiempoReceso")
        Me.dtpFechEntrada.Value = Me.adoExtraTurno.Recordset.Fields("FechaEntrada")
        Me.dtpFecSalida.Value = Me.adoExtraTurno.Recordset.Fields("FechaSalida")
        Me.mskAsistHoraEntrada.Text = Me.adoExtraTurno.Recordset.Fields("HoraEntrada")
        
        Me.lblFecha.Caption = Me.adoExtraTurno.Recordset.Fields("FechaEntrada")
        Me.cmdModificar.Enabled = True
        Me.cmdBorrar.Enabled = True
        
     Else
        Me.lblFecha.Caption = ""
        Me.cmdModificar.Enabled = False
        
     End If
  End If
      
  
      
   ElseIf Me.dtpFecha.Value <> "__/__/____" Then
     
     Me.adoExtraTurno.CommandType = adCmdText
     Me.adoExtraTurno.RecordSource = "SELECT * FROM ExtraTurno WHERE FechaEntrada = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102)"
     Me.adoExtraTurno.Refresh
        
     Me.adoEmpleado.Refresh
     Me.lstElegidos.Clear
     
     Do While Not Me.adoExtraTurno.Recordset.EOF
      
      Me.adoEmpleado.Recordset.Find "CodEmpleado ='" & Trim(Me.adoExtraTurno.Recordset.Fields("CodEmpleado")) & "'"
      
      Me.adoEmpleado.Recordset.Find "[CodEmpleado] ='" & Me.adoExtraTurno.Recordset.Fields("CodEmpleado") & "'"
      Me.lstElegidos.AddItem Me.adoExtraTurno.Recordset.Fields("CodEmpleado") & " " & Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
      bEncontrado = True
      
      
      Me.txtTiempoReceso.Text = Me.adoExtraTurno.Recordset.Fields("TiempoReceso")
      Me.dtpFechEntrada.Value = Me.adoExtraTurno.Recordset.Fields("FechaEntrada")
      Me.dtpFecSalida.Value = Me.adoExtraTurno.Recordset.Fields("FechaSalida")
      Me.mskAsistHoraEntrada.Text = Me.adoExtraTurno.Recordset.Fields("HoraEntrada")
      
      Me.lblFecha.Caption = Me.adoExtraTurno.Recordset.Fields("FechaEntrada")
      
      Me.adoEmpleado.Refresh
      Me.adoExtraTurno.Recordset.MoveNext
      
     Loop
     
     
     If bEncontrado Then
        Me.cmdModificar.Enabled = True
        Me.cmdBorrar.Enabled = True
        
     Else
        Me.cmdModificar.Enabled = False
        Me.lblFecha.Caption = ""
     End If
     
     
      
   End If
    





End Sub

Private Sub cmdBuscar_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   
   If Me.cmdAgregar.Enabled = True Then
      Me.cmdAgregar.SetFocus
   ElseIf Me.cmdModificar.Enabled Then
     Me.cmdModificar.SetFocus
   End If
   


End If


End Sub

Private Sub cmdDepto_Click()

If Me.cboDepto.Text <> "" Then
  
  Me.adoDepto.CommandType = adCmdText
  Me.adoDepto.RecordSource = "SELECT dbo.Empleado.CodEmpleado, dbo.Empleado.Nombre1, dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2, " & _
                             "dbo.Departamento.Departamento FROM dbo.Empleado INNER JOIN Dbo.Departamento ON dbo.Empleado.CodDepartamento = dbo.Departamento.CodDepartamento " & _
                             "WHERE Departamento.Departamento ='" & Me.cboDepto.Text & "'"
  Me.adoDepto.Refresh
  
  Me.lblFecha.Caption = Me.dtpFechEntrada.Value
  
  If Not Me.adoDepto.Recordset.EOF Then
     Me.lstEmpleados.Clear
     
     Do While Not Me.adoDepto.Recordset.EOF
       
      Me.lstEmpleados.AddItem Me.adoDepto.Recordset.Fields("CodEmpleado") & " " & Me.adoDepto.Recordset.Fields("Nombre1") & " " & Me.adoDepto.Recordset.Fields("Nombre2") & Me.adoDepto.Recordset.Fields("Apellido1") & " " & Me.adoDepto.Recordset.Fields("Apellido2")
      Me.adoDepto.Recordset.MoveNext
      

     Loop
     
  End If

End If


End Sub

Private Sub cmdDepto_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   
   
   Me.lstEmpleados.SetFocus


End If


End Sub

Private Sub cmdJalar_Click()

Dim iCont As Integer
Dim iContElegidos As Integer
Dim bEncontrado As Boolean

If Me.lstEmpleados.ListIndex > 0 Then

  Do While iCont <= Me.lstEmpleados.ListCount - 1
      
     If Me.lstEmpleados.Selected(iCont) Then
        bEncontrado = False
        iContElegidos = 0
        Do While iContElegidos <= Me.lstElegidos.ListCount - 1
            
            If Mid$(Me.lstElegidos.List(iContElegidos), 1, 6) = Mid$(Me.lstEmpleados.List(iCont), 1, 6) Then
               bEncontrado = True
            End If
            
            iContElegidos = iContElegidos + 1
            
        Loop
        
        If Not bEncontrado Then
           Me.lstElegidos.AddItem Me.lstEmpleados.List(iCont)
        End If
        
     End If
     
     iCont = iCont + 1
     


  Loop
  
End If

End Sub

Private Sub cmdModificar_Click()

Dim sFechaEntrada As String
Dim iCont As Integer



If Me.cmdModificar.Caption = "&Modificar" Then
   Me.cmdAgregar.Enabled = False
   Me.cmdModificar.Caption = "&Guardar"
   Me.fraGeneral.Enabled = True
   Me.lstElegidos.SetFocus
   Me.cmdBorrar.Enabled = False

ElseIf Me.cmdModificar.Caption = "&Guardar" And Me.lblFecha.Caption <> "" Then
   
   sFechaEntrada = Mid$(Me.lblFecha.Caption, 7, 4) & "-" & Mid$(Me.lblFecha.Caption, 4, 2) & "-" & Mid$(Me.lblFecha.Caption, 1, 2)

    If Chequear_Hora(Me.mskAsistHoraEntrada.Text) And IsNumeric(Me.txtTiempoReceso.Text) Then

       Do While iCont < Me.lstElegidos.ListCount
           
           If Me.lstElegidos.Selected(iCont) Then
              Me.adoExtraTurno.CommandType = adCmdText
              Me.adoExtraTurno.RecordSource = "SELECT * FROM ExtraTurno WHERE CodEmpleado ='" & Mid$(Me.lstElegidos.List(iCont), 1, 6) & "' AND FechaEntrada = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102)"
              Me.adoExtraTurno.Refresh
     
             If Not Me.adoExtraTurno.Recordset.EOF Then
               Me.adoExtraTurno.Recordset.Fields("FechaEntrada") = Me.dtpFechEntrada.Value
               Me.adoExtraTurno.Recordset.Fields("HoraEntrada") = Me.mskAsistHoraEntrada.Text
               Me.adoExtraTurno.Recordset.Fields("FechaSalida") = Me.dtpFecSalida.Value
               Me.adoExtraTurno.Recordset.Fields("TiempoReceso") = Me.txtTiempoReceso.Text
               Me.adoExtraTurno.Recordset.Update
               
             End If
             
           End If
           
           iCont = iCont + 1
     
        Loop
        
        
        Me.cmdAgregar.Enabled = True
        Me.cmdModificar.Caption = "&Modificar"
        Me.fraGeneral.Enabled = True
        Me.cboCodigo.SetFocus
        Me.cmdBorrar.Enabled = True
        
 Else
     MsgBox "Información incorrecta en la hora de entrada ó tiempo de receso", vbInformation
       
  
End If
  
End If


End Sub

Private Sub cmdQuitar_Click()

Dim iCont As Integer
Dim iContElegidos As Integer
Dim bEncontrado As Boolean

If Me.cmdModificar.Caption = "&Guardar" Then

If Me.lstElegidos.ListIndex >= 0 Then

  Do While iCont <= Me.lstElegidos.ListCount - 1
      
     If Me.lstElegidos.Selected(iCont) Then
        Me.lstElegidos.RemoveItem (iCont)
     End If
     
     iCont = iCont + 1
     
  Loop
  
End If

End If

End Sub

Private Sub cmdReporte_Click()

Dim sSQL As String
Dim sFechaEntrada As String
Dim rptExtraTiempo As New arepExtraTurno


sFechaEntrada = Mid$(Me.dtpFecha.Value, 7, 4) & "-" & Mid$(Me.dtpFecha.Value, 4, 2) & "-" & Mid$(Me.dtpFecha.Value, 1, 2)

sSQL = "SELECT dbo.ExtraTurno.CodEmpleado, dbo.Departamento.Departamento, dbo.ExtraTurno.CodDepartamento, dbo.TipoNomina.Nomina, " & _
       "dbo.ExtraTurno.FechaEntrada, dbo.ExtraTurno.HoraEntrada, dbo.ExtraTurno.FechaSalida, dbo.ExtraTurno.HoraSalida, dbo.ExtraTurno.bActivo, " & _
       "dbo.ExtraTurno.HorasLaboradas , dbo.ExtraTurno.TiempoReceso, dbo.Empleado.Nombre1, dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2 " & _
       "FROM dbo.ExtraTurno INNER JOIN " & _
       "dbo.Empleado ON dbo.ExtraTurno.CodEmpleado = dbo.Empleado.CodEmpleado INNER JOIN " & _
       "dbo.Departamento ON dbo.Empleado.CodDepartamento = dbo.Departamento.CodDepartamento INNER JOIN " & _
       "dbo.TipoNomina ON dbo.ExtraTurno.CodTipoNomina = dbo.TipoNomina.CodTipoNomina " & _
       "WHERE (dbo.ExtraTurno.FechaEntrada = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00', 102))"


rptExtraTiempo.adoControl.ConnectionString = Conexion
rptExtraTiempo.lblMensaje.Caption = "Fecha: " & Me.dtpFecha.Value
rptExtraTiempo.adoControl.Source = sSQL
rptExtraTiempo.Show


End Sub

Private Sub cmdSalir_Click()
  
  Unload Me

End Sub



Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
  
  Me.cmdBuscar.SetFocus

End If


End Sub



Private Sub dtpFechEntrada_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then

  Me.mskAsistHoraEntrada.SetFocus

End If


End Sub


Private Sub dtpFecSalida_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then

  Me.txtTiempoReceso.SetFocus

End If


End Sub

Private Sub Form_Load()


 Dim RutaServer As String
 Dim Server As String
 

 Dim ConexionSTR1 As String
 Dim TxtClaveEntrada As String
'abro el archivo para solo lectura de la cadena de conexion
 Dim NextLine As String
 Dim Autorizado As Boolean
   Autorizado = False

 Open App.Path + "\SysInfo.dll" For Input As #1
  Do Until EOF(1)
   Line Input #1, NextLine
        ConexionSTR1 = Trim(NextLine)
   Loop
 Close #1
  
  
  Conexion = ConexionSTR1

Me.adoEmpleado.ConnectionString = Conexion
Me.adoEmpleado.CommandType = adCmdTable
Me.adoEmpleado.RecordSource = "Empleado"
Me.adoEmpleado.Refresh


Me.adoDepto.ConnectionString = Conexion
Me.adoDepto.CommandType = adCmdTable
Me.adoDepto.RecordSource = "Departamento"
Me.adoDepto.Refresh

Me.adoExtraTurno.ConnectionString = Conexion
Me.adoExtraTurno.CommandType = adCmdTable
Me.adoExtraTurno.RecordSource = "ExtraTurno"
Me.adoExtraTurno.Refresh


Do While Not Me.adoDepto.Recordset.EOF

  Me.cboDepto.AddItem Me.adoDepto.Recordset.Fields("Departamento")
  Me.adoDepto.Recordset.MoveNext


Loop

Me.adoEmpleado.Refresh

Do While Not Me.adoEmpleado.Recordset.EOF

  Me.cboCodigo.AddItem Me.adoEmpleado.Recordset.Fields("CodEmpleado")
  Me.adoEmpleado.Recordset.MoveNext


Loop

Me.adoEmpleado.Refresh




End Sub





Private Sub lstEmpleados_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then

  Me.cmdJalar.SetFocus

End If


End Sub

Private Sub mskAsistHoraEntrada_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   
   Me.dtpFecSalida.SetFocus

End If


End Sub


Private Sub txtTiempoReceso_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = 13 Then
   
   Me.cboDepto.SetFocus

End If

End Sub
