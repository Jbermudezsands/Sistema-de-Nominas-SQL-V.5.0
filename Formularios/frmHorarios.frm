VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmHorarios 
   Caption         =   "Horarios"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   7440
      TabIndex        =   41
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   5040
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc adoHorarioEmpl 
      Height          =   375
      Left            =   5640
      Top             =   6480
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
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
      Caption         =   "                     Horarios Empleados"
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
   Begin MSAdodcLib.Adodc adoTurnos 
      Height          =   375
      Left            =   4440
      Top             =   6480
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
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
      Caption         =   "Turnos"
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
      Left            =   3600
      Top             =   6480
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
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
      Caption         =   "Empleados"
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
   Begin VB.Frame fraEmpleado 
      Caption         =   "Empleado"
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   12135
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCodEmpl 
         Height          =   285
         Left            =   2760
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblEmpleado 
         Height          =   255
         Left            =   6000
         TabIndex        =   40
         Top             =   600
         Width           =   5775
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Codigo o Nombre del Empleado:"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   600
         Width           =   2280
      End
   End
   Begin VB.Frame fraHorarios 
      Caption         =   "Horarios"
      Enabled         =   0   'False
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   12135
      Begin VB.TextBox txtTComida 
         Height          =   285
         Left            =   10320
         TabIndex        =   37
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtSDomingo 
         Height          =   285
         Left            =   8640
         TabIndex        =   35
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtEDomingo 
         Height          =   285
         Left            =   8640
         TabIndex        =   34
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtSSabado 
         Height          =   285
         Left            =   7440
         TabIndex        =   33
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtESabado 
         Height          =   285
         Left            =   7440
         TabIndex        =   32
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtSViernes 
         Height          =   285
         Left            =   6240
         TabIndex        =   31
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtEViernes 
         Height          =   285
         Left            =   6240
         TabIndex        =   30
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtSJueves 
         Height          =   285
         Left            =   4920
         TabIndex        =   29
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtEJueves 
         Height          =   285
         Left            =   4920
         TabIndex        =   28
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtSMiercoles 
         Height          =   285
         Left            =   3600
         TabIndex        =   27
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtEMiercoles 
         Height          =   285
         Left            =   3600
         TabIndex        =   26
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtSMartes 
         Height          =   285
         Left            =   2280
         TabIndex        =   25
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtEMartes 
         Height          =   285
         Left            =   2280
         TabIndex        =   24
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtSLunes 
         Height          =   285
         Left            =   840
         TabIndex        =   23
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtELunes 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ComboBox cboDomingo 
         Height          =   315
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cboSabado 
         Height          =   315
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cboViernes 
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cboJueves 
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cboMiercoles 
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cboMartes 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cboLunes 
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Minutos"
         Height          =   195
         Left            =   11160
         TabIndex        =   38
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label Label11 
         Caption         =   "Tiempo de Comida"
         Height          =   255
         Left            =   10080
         TabIndex        =   36
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Turno"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Domingo"
         Height          =   195
         Left            =   8760
         TabIndex        =   20
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Sabado"
         Height          =   195
         Left            =   7560
         TabIndex        =   19
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Viernes"
         Height          =   195
         Left            =   6360
         TabIndex        =   18
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Jueves"
         Height          =   195
         Left            =   5040
         TabIndex        =   17
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Miercoles"
         Height          =   195
         Left            =   3840
         TabIndex        =   16
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Martes"
         Height          =   195
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Lunes"
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Salida"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Entrada"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmHorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCodEmpl As String

Private Sub cboJueves_Change()

   
   
   If Me.cboJueves.Text <> "" Then
  
   Me.adoTurnos.Recordset.Find "CodTurno LIKE " & Me.cboJueves.Text
   
   If Not Me.adoTurnos.Recordset.EOF Then
      
      Me.txtEJueves.Text = Me.adoTurnos.Recordset.Fields("JEntrada")
      Me.txtSJueves.Text = Me.adoTurnos.Recordset.Fields("JSalida")
   
   
   
   End If
  
   Me.adoTurnos.Refresh
  
  
  
  End If



End Sub

Private Sub cboLunes_Change()

  
  If Me.cboLunes.Text <> "" Then
  
   Me.adoTurnos.Recordset.Find "CodTurno LIKE " & Me.cboLunes.Text
   
   If Not Me.adoTurnos.Recordset.EOF Then
      
      Me.txtELunes.Text = Me.adoTurnos.Recordset.Fields("LEntrada")
      Me.txtSLunes.Text = Me.adoTurnos.Recordset.Fields("LSalida")
   
   
   
   End If
  
   Me.adoTurnos.Refresh
  
  
  
  End If
  
  
  


End Sub


Private Sub cboMartes_Change()

If Me.cboMartes.Text <> "" Then
  
   Me.adoTurnos.Recordset.Find "CodTurno LIKE " & Me.cboMartes.Text
   
   If Not Me.adoTurnos.Recordset.EOF Then
      
      Me.txtEMartes.Text = Me.adoTurnos.Recordset.Fields("MEntrada")
      Me.txtSMartes.Text = Me.adoTurnos.Recordset.Fields("MSalida")
   
   
   
   End If
  
   Me.adoTurnos.Refresh
  
  
  
  End If



End Sub

Private Sub cboMiercoles_Change()


If Me.cboMiercoles.Text <> "" Then
  
   Me.adoTurnos.Recordset.Find "CodTurno LIKE " & Me.cboMiercoles.Text
   
   If Not Me.adoTurnos.Recordset.EOF Then
      
      Me.txtEMiercoles.Text = Me.adoTurnos.Recordset.Fields("MCEntrada")
      Me.txtSMiercoles.Text = Me.adoTurnos.Recordset.Fields("MCSalida")
   
   
   
   End If
  
   Me.adoTurnos.Refresh
  
  
  
  End If



End Sub

Private Sub cboSabado_Change()


If Me.cboSabado.Text <> "" Then
  
   Me.adoTurnos.Recordset.Find "CodTurno LIKE " & Me.cboSabado.Text
   
   If Not Me.adoTurnos.Recordset.EOF Then
      
      Me.txtESabado.Text = Me.adoTurnos.Recordset.Fields("SEntrada")
      Me.txtSSabado.Text = Me.adoTurnos.Recordset.Fields("SSalida")
   
   
   
   End If
  
   Me.adoTurnos.Refresh
  
  
  
  End If


End Sub

Private Sub cboViernes_Change()


If Me.cboViernes.Text <> "" Then
  
   Me.adoTurnos.Recordset.Find "CodTurno LIKE " & Me.cboViernes.Text
   
   If Not Me.adoTurnos.Recordset.EOF Then
      
      Me.txtEViernes.Text = Me.adoTurnos.Recordset.Fields("VEntrada")
      Me.txtSViernes.Text = Me.adoTurnos.Recordset.Fields("VSalida")
   
   
   
   End If
  
   Me.adoTurnos.Refresh
  
  
  
  End If



End Sub

Private Sub cmdBuscar_Click()

If Trim(Me.txtCodEmpl.Text) <> "" Then

    Me.adoEmpleado.Recordset.Find "CodEmpleado LIKE *" & Me.txtCodEmpl.Text & "*" ' OR Nombre1 LIKE " & Me.txtCodEmpl.Text & " OR Nombre2 LIKE " & Me.txtCodEmpl.Text & " OR Apellido1 LIKE " & Me.txtCodEmpl.Text & " OR Apellido2 LIKE " & Me.txtCodEmpl.Text & ""
    
    If Not Me.adoEmpleado.Recordset.EOF Then
     
     Me.lblEmpleado.Caption = Me.adoEmpleado.Recordset.Fields("CodEmpleado") & ", " & Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
     Me.adoHorarioEmpl.Recordset.Find "CodEmpleado LIKE *" & Me.adoEmpleado.Recordset.Fields("CodEmpleado") & "*"
     
     sCodEmpl = Me.adoEmpleado.Recordset.Fields("CodEmpleado")
     
     Me.txtELunes.Text = Me.adoHorarioEmpl.Recordset.Fields("LEntrada")
     Me.txtSLunes.Text = Me.adoHorarioEmpl.Recordset.Fields("LSalida")
     Me.txtEMartes.Text = Me.adoHorarioEmpl.Recordset.Fields("MEntrada")
     Me.txtSMartes.Text = Me.adoHorarioEmpl.Recordset.Fields("MSalida")
     Me.txtEMiercoles.Text = Me.adoHorarioEmpl.Recordset.Fields("MCEntrada")
     Me.txtSMiercoles.Text = Me.adoHorarioEmpl.Recordset.Fields("MCSalida")
     Me.txtEJueves.Text = Me.adoHorarioEmpl.Recordset.Fields("JEntrada")
     Me.txtSJueves.Text = Me.adoHorarioEmpl.Recordset.Fields("JSalida")
     Me.txtEViernes.Text = Me.adoHorarioEmpl.Recordset.Fields("VEntrada")
     Me.txtSViernes.Text = Me.adoHorarioEmpl.Recordset.Fields("VSalida")
     Me.txtESabado.Text = Me.adoHorarioEmpl.Recordset.Fields("SEntrada")
     Me.txtSSabado.Text = Me.adoHorarioEmpl.Recordset.Fields("SSalida")
     Me.txtEDomingo.Text = Me.adoHorarioEmpl.Recordset.Fields("DEntrada")
     Me.txtSDomingo.Text = Me.adoHorarioEmpl.Recordset.Fields("DSalida")
     Me.cboLunes.Text = Me.adoHorarioEmpl.Recordset.Fields("TurnoLunes")
     Me.cboMartes.Text = Me.adoHorarioEmpl.Recordset.Fields("TurnoMartes")
     Me.cboMiercoles.Text = Me.adoHorarioEmpl.Recordset.Fields("TurnoMiercoles")
     Me.cboJueves.Text = Me.adoHorarioEmpl.Recordset.Fields("TurnoJueves")
     Me.cboViernes.Text = Me.adoHorarioEmpl.Recordset.Fields("TurnoViernes")
     Me.cboSabado.Text = Me.adoHorarioEmpl.Recordset.Fields("TurnoSabado")
     Me.cboDomingo.Text = Me.adoHorarioEmpl.Recordset.Fields("TurnoDomingo")
     Me.txtTComida.Text = Me.adoHorarioEmpl.Recordset.Fields("TComida")
     
     
   Else
      
      MsgBox "No se encontro ningun registro con el criterio: " & Me.txtCodEmpl.Text
      sCodEmpl = ""
      Exit Sub
   End If
     
     
End If




End Sub

Private Sub cmdModificar_Click()

   
   If Me.sCodEmpl <> "" Then
   
     If Me.cmdModificar.Caption = "&Modificar" Then
        Me.cmdModificar.Caption = "&Guardar Cambios"
        Me.fraHorarios.Enabled = False
        Me.cboLunes.SetFocus
        
     
     Else
       
      Me.cmdModificar.Caption = "&Modificar"
      Me.txtCodEmpl.SetFocus
     
      Me.fraHorarios.Enabled = True
      
     End If
   
   
   
   
   End If
   


End Sub

Private Sub cmdSalir_Click()

Unload Me

End Sub

Private Sub Form_Load()

 Dim RutaServer As String
 Dim Server As String
 Dim Conexion As String
 
' RutaServer = App.Path + "\CntNominas.dll"
'
'  With Me.DtaServidor
'     .DatabaseName = RutaServer
'     .RecordSource = "Servidor"
'     .Refresh
'  End With
'
'  If Not IsNull(Me.DtaServidor.Recordset.Servidor) Then
'   Server = Me.DtaServidor.Recordset.Servidor
'  Else
'   MsgBox "No se ha definido el Servidor", vbCritical, "Sistmea de Nominas"
'   Exit Sub
'  End If


'Borrar despues la siguiente linea
Server = "Moises\Moises"

Conexion = "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=" & Server

Me.adoEmpleado.ConnectionString = Conexion
Me.adoEmpleado.CommandType = adCmdTable
Me.adoEmpleado.RecordSource = "Empleado"
Me.adoEmpleado.Refresh


Me.adoHorarioEmpl.ConnectionString = Conexion
Me.adoHorarioEmpl.CommandType = adCmdTable
Me.adoHorarioEmpl.RecordSource = "HorarioEmpleado"
Me.adoHorarioEmpl.Refresh

Me.adoTurnos.ConnectionString = Conexion
Me.adoTurnos.CommandType = adCmdTable
Me.adoTurnos.RecordSource = "Turno"
Me.adoTurnos.Refresh

Do While Not Me.adoTurnos.Recordset.EOF

  Me.cboLunes.AddItem Me.adoTurnos.Recordset.Fields("CodTurno")
  Me.cboMartes.AddItem Me.adoTurnos.Recordset.Fields("CodTurno")
  Me.cboMiercoles.AddItem Me.adoTurnos.Recordset.Fields("CodTurno")
  Me.cboJueves.AddItem Me.adoTurnos.Recordset.Fields("CodTurno")
  Me.cboViernes.AddItem Me.adoTurnos.Recordset.Fields("CodTurno")
  Me.cboSabado.AddItem Me.adoTurnos.Recordset.Fields("CodTurno")
  Me.cboDomingo.AddItem Me.adoTurnos.Recordset.Fields("CodTurno")
  
  Me.adoTurnos.Recordset.MoveNext
  
Loop

Me.adoTurnos.Refresh





End Sub


