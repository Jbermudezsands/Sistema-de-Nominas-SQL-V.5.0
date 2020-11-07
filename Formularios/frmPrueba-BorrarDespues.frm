VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPrueba 
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc adoHistorico 
      Height          =   375
      Left            =   1800
      Top             =   4440
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
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=PAYROLL"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=PAYROLL"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Historico"
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton cmdCopiar 
      Caption         =   "Ejecutar"
      Height          =   735
      Left            =   720
      TabIndex        =   4
      Top             =   5280
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc adoTurno 
      Height          =   375
      Left            =   1800
      Top             =   3960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=PAYROLL"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=PAYROLL"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Turno"
      Caption         =   "Adodc1"
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
   Begin MSAdodcLib.Adodc adoHorarioEmpl 
      Height          =   375
      Left            =   1680
      Top             =   2520
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=PAYROLL"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=PAYROLL"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "HorarioEmpleado"
      Caption         =   "Adodc1"
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
   Begin MSAdodcLib.Adodc adoEmpleadoSQL 
      Height          =   375
      Left            =   1440
      Top             =   3360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=PAYROLL"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=PAYROLL"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Empleado"
      Caption         =   "Adodc1"
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
   Begin MSAdodcLib.Adodc adoEmpleadoViejo 
      Height          =   375
      Left            =   2520
      Top             =   2880
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\RespaldoSQL\PlanMetro.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\RespaldoSQL\PlanMetro.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Empleado"
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtResultado 
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtSalida 
      DataField       =   "LSalida"
      DataSource      =   "adoPrueba"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtEntrada 
      DataField       =   "LEntrada"
      DataSource      =   "adoPrueba"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2160
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc adoPrueba 
      Height          =   375
      Left            =   1080
      Top             =   960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=PAYROLL"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=PAYROLL"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "AsistenciaEmpleado"
      Caption         =   "Adodc1"
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
End
Attribute VB_Name = "frmPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub cmdCopiar_Click()

Dim iHistorico As Integer
Dim sCodEmpl As String
Dim bLongitudCod As Byte


'Me.adoEmpleadoViejo.CommandType = adCmdText
'Me.adoEmpleadoViejo.RecordSource = "SELECT * FROM Empleado ORDER BY Cod_Empl ASC"
'Me.adoEmpleadoViejo.Refresh

Me.adoEmpleadoSQL.CommandType = adCmdTable
Me.adoEmpleadoSQL.RecordSource = "Empleado"
Me.adoEmpleadoSQL.Refresh

Me.adoHistorico.CommandType = adCmdTable
Me.adoHistorico.RecordSource = "Historico"
Me.adoHistorico.Refresh

Me.adoHorarioEmpl.CommandType = adCmdTable
Me.adoHorarioEmpl.RecordSource = "HorarioEmpleado"
Me.adoHorarioEmpl.Refresh

Me.adoTurno.CommandType = adCmdTable
Me.adoTurno.RecordSource = "Turno"
Me.adoTurno.Refresh

Me.adoPrueba.CommandType = adCmdTable
Me.adoPrueba.RecordSource = "AsistenciaEmpleado"
Me.adoPrueba.Refresh


iHistorico = 1

Do While Not Me.adoEmpleadoSQL.Recordset.EOF

   sCodEmpl = Me.adoEmpleadoSQL.Recordset.Fields("CodEmpleado")
   
   Me.adoHorarioEmpl.CommandType = adCmdText
   Me.adoHorarioEmpl.RecordSource = "SELECT * FROM HorarioEmpleado WHERE CodEmpleado ='" & Me.adoEmpleadoSQL.Recordset.Fields("CodEmpleado") & "'"
   Me.adoHorarioEmpl.Refresh
   
   Me.adoPrueba.CommandType = adCmdText
   Me.adoPrueba.RecordSource = "SELECT CodEmpleado, FechaEntrada, FechaSalida FROM AsistenciaEmpleado WHERE CodEmpleado ='" & Me.adoEmpleadoSQL.Recordset.Fields("CodEmpleado") & "' AND FechaEntrada = CONVERT(DATETIME, '" & "2005-08-10" & " 00:00:00" & "', 102)"
   Me.adoPrueba.Refresh

   
'   bLongitudCod = 6 - Len(sCodEmpl)
   
'   Select Case bLongitudCod
'
'
'   Case 1:
'      sCodEmpl = "0" & sCodEmpl
'
'   Case 2:
'      sCodEmpl = "00" & sCodEmpl
'
'   Case 3:
'      sCodEmpl = "000" & sCodEmpl
'
'   Case 4:
'      sCodEmpl = "0000" & sCodEmpl
'
'   Case 5:
'      sCodEmpl = "00000" & sCodEmpl
'
' End Select
 
 
' Me.adoPrueba.Recordset.AddNew
' Me.adoPrueba.Recordset.Fields("CodEmpleado") = sCodEmpl
' Me.adoPrueba.Recordset.Fields("FechaEntrada") = "01/08/2005"
' Me.adoPrueba.Recordset.Fields("HoraEntrada") = Me.adoHorarioEmpl.Recordset.Fields("LEntrada")
' Me.adoPrueba.Recordset.Fields("FechaSalida") = "01/08/2005"
' Me.adoPrueba.Recordset.Fields("HoraSalida") = Me.adoHorarioEmpl.Recordset.Fields("LSalida")
' Me.adoPrueba.Recordset.Fields("bActivo") = 0
' Me.adoPrueba.Recordset.Fields("CodTipoNomina") = Me.adoEmpleadoSQL.Recordset.Fields("CodTipoNomina")
' Me.adoPrueba.Recordset.Fields("CodTurno") = "Diurno"
' Me.adoPrueba.Recordset.Update
' Me.adoPrueba.Refresh
 
' Me.adoPrueba.Recordset.AddNew
' Me.adoPrueba.Recordset.Fields("CodEmpleado") = sCodEmpl
' Me.adoPrueba.Recordset.Fields("FechaEntrada") = "10/08/2005"
' Me.adoPrueba.Recordset.Fields("HoraEntrada") = Me.adoHorarioEmpl.Recordset.Fields("MCEntrada")
 Me.adoPrueba.Recordset.Fields("FechaSalida") = "10/08/2005"
' Me.adoPrueba.Recordset.Fields("HoraSalida") = Me.adoHorarioEmpl.Recordset.Fields("MCSalida")
' Me.adoPrueba.Recordset.Fields("bActivo") = 0
' Me.adoPrueba.Recordset.Fields("CodTipoNomina") = Me.adoEmpleadoSQL.Recordset.Fields("CodTipoNomina")
' Me.adoPrueba.Recordset.Fields("CodTurno") = "Diurno"
 Me.adoPrueba.Recordset.Update
 Me.adoPrueba.Refresh
 
 
 
' Me.adoHistorico.Recordset.AddNew
' Me.adoHistorico.Recordset.Fields("Id") = iHistorico
' Me.adoHistorico.Recordset.Fields("CodEmpleado") = sCodEmpl
' Me.adoHistorico.Recordset.Fields("FechaContrato") = Me.adoEmpleadoViejo.Recordset.Fields("Fech_Ing")
' Me.adoHistorico.Recordset.Fields("SueldoActual") = 0
' Me.adoHistorico.Recordset.Fields("SueldoAnterior") = 0
' Me.adoHistorico.Recordset.Fields("SueldoInicial") = 0
' Me.adoHistorico.Recordset.Update
'
' iHistorico = iHistorico + 1
'
' Me.adoHorarioEmpl.Recordset.AddNew
' Me.adoHorarioEmpl.Recordset.Fields("CodEmpleado") = sCodEmpl
' Me.adoHorarioEmpl.Recordset.Fields("LEntrada") = Me.adoTurno.Recordset.Fields("LEntrada")
' Me.adoHorarioEmpl.Recordset.Fields("LSalida") = Me.adoTurno.Recordset.Fields("LSalida")
' Me.adoHorarioEmpl.Recordset.Fields("MEntrada") = Me.adoTurno.Recordset.Fields("MEntrada")
' Me.adoHorarioEmpl.Recordset.Fields("MSalida") = Me.adoTurno.Recordset.Fields("MSalida")
' Me.adoHorarioEmpl.Recordset.Fields("MCEntrada") = Me.adoTurno.Recordset.Fields("MCEntrada")
' Me.adoHorarioEmpl.Recordset.Fields("MCSalida") = Me.adoTurno.Recordset.Fields("MCSalida")
' Me.adoHorarioEmpl.Recordset.Fields("JEntrada") = Me.adoTurno.Recordset.Fields("JEntrada")
' Me.adoHorarioEmpl.Recordset.Fields("JSalida") = Me.adoTurno.Recordset.Fields("JSalida")
' Me.adoHorarioEmpl.Recordset.Fields("VEntrada") = Me.adoTurno.Recordset.Fields("VEntrada")
' Me.adoHorarioEmpl.Recordset.Fields("VSalida") = Me.adoTurno.Recordset.Fields("VSalida")
' Me.adoHorarioEmpl.Recordset.Fields("SEntrada") = Me.adoTurno.Recordset.Fields("SEntrada")
' Me.adoHorarioEmpl.Recordset.Fields("SSalida") = Me.adoTurno.Recordset.Fields("SSalida")
' Me.adoHorarioEmpl.Recordset.Fields("DEntrada") = Me.adoTurno.Recordset.Fields("DEntrada")
' Me.adoHorarioEmpl.Recordset.Fields("DSalida") = Me.adoTurno.Recordset.Fields("DSalida")
' Me.adoHorarioEmpl.Recordset.Fields("TurnoLunes") = Me.adoTurno.Recordset.Fields("CodTurno")
' Me.adoHorarioEmpl.Recordset.Fields("TurnoMartes") = Me.adoTurno.Recordset.Fields("CodTurno")
' Me.adoHorarioEmpl.Recordset.Fields("TurnoMiercoles") = Me.adoTurno.Recordset.Fields("CodTurno")
' Me.adoHorarioEmpl.Recordset.Fields("TurnoJueves") = Me.adoTurno.Recordset.Fields("CodTurno")
' Me.adoHorarioEmpl.Recordset.Fields("TurnoViernes") = Me.adoTurno.Recordset.Fields("CodTurno")
' Me.adoHorarioEmpl.Recordset.Fields("TurnoSabado") = Me.adoTurno.Recordset.Fields("CodTurno")
' Me.adoHorarioEmpl.Recordset.Fields("TurnoDomingo") = Me.adoTurno.Recordset.Fields("CodTurno")
' Me.adoHorarioEmpl.Recordset.Fields("TComida") = Me.adoTurno.Recordset.Fields("TComida")
' Me.adoHorarioEmpl.Recordset.Update
' Me.adoHorarioEmpl.Refresh
 
 Me.adoEmpleadoSQL.Recordset.MoveNext
 
 

Loop




End Sub

Private Sub Command1_Click()

Dim dHEntradaHoy As Date

'Me.txtSalida.Text = Me.adoPrueba.Recordset.Fields("LSalida") - Me.adoPrueba.Recordset.Fields("LEntrada")

'Para deducir el valor del valor de horas trabajadas por dia

 Me.txtSalida.Text = (DateDiff("n", Me.adoPrueba.Recordset.Fields("LEntrada"), Me.adoPrueba.Recordset.Fields("LSalida")) / 60) - (Me.adoPrueba.Recordset.Fields("TComida") / 60)



If Me.adoPrueba.Recordset.Fields("LEntrada") > Time Then
         dHEntradaHoy = Me.adoPrueba.Recordset.Fields("LEntrada")
      Else
         dHEntradaHoy = Time
      End If


End Sub

