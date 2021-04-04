VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AF8CD3F4-666F-11D1-940D-000021A73813}#5.0#0"; "osProgress.ocx"
Begin VB.Form frmDevengadoHora 
   Caption         =   "Devengado x Hora"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.Data dtaServidor 
      Caption         =   "Conexion con el servidor"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin Progress.osProgress ospBarra 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   3120
      Width           =   4575
      _ExtentX        =   6694
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7215
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   3120
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtMes 
         Height          =   285
         Left            =   4200
         TabIndex        =   7
         Text            =   " "
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtAno 
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtPeriodo 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Text            =   " "
         Top             =   810
         Width           =   975
      End
      Begin VB.CommandButton cmdCopiar 
         Caption         =   "Copiar Devengado"
         Height          =   495
         Left            =   840
         TabIndex        =   1
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   3840
         TabIndex        =   8
         Top             =   840
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   2160
         TabIndex        =   5
         Top             =   840
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   540
      End
   End
   Begin MSAdodcLib.Adodc adoEmpleadoSQL 
      Height          =   375
      Left            =   960
      Top             =   4080
      Visible         =   0   'False
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
      Connect         =   $"DevengadoHora.frx":0000
      OLEDBString     =   $"DevengadoHora.frx":0088
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Empleado"
      Caption         =   "Empleado SQL"
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
      Left            =   2400
      Top             =   3960
      Visible         =   0   'False
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Empleado"
      Caption         =   "Empleado Viejo"
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
   Begin MSAdodcLib.Adodc adoIngresos 
      Height          =   495
      Left            =   2400
      Top             =   3600
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ingresos"
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
   Begin MSAdodcLib.Adodc adoFechaPlanilla 
      Height          =   375
      Left            =   480
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Fecha_Planilla"
      Caption         =   "Fecha Planilla"
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
   Begin MSAdodcLib.Adodc adoAsistencia 
      Height          =   330
      Left            =   120
      Top             =   3840
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   582
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
      Connect         =   $"DevengadoHora.frx":0110
      OLEDBString     =   $"DevengadoHora.frx":0198
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "AsistenciaEmpleado"
      Caption         =   "Asistencia Empleado"
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
   Begin MSAdodcLib.Adodc adoPrueba 
      Height          =   375
      Left            =   960
      Top             =   3840
      Visible         =   0   'False
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Devengado_Hora"
      Caption         =   "Devengado por Hora"
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
   Begin MSAdodcLib.Adodc adoAntiguedad 
      Height          =   375
      Left            =   2400
      Top             =   3480
      Visible         =   0   'False
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Reloj\PlanMetro.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Antiguedad"
      Caption         =   "Antiguedad"
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
Attribute VB_Name = "frmDevengadoHora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Function BuscarIncentivo(sTotal As String, Fecha As Date) As Single


Dim fecAct As Date

fecAct = Format(Date, "Short Date")
adoAntiguedad.RecordSource = "Antiguedad"
adoAntiguedad.Refresh

  ' 1 Año
If fecAct - Fecha >= 365 And fecAct - Fecha <= 2 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 1 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
  ' 2 Años
    
ElseIf fecAct - Fecha >= 2 * 365 And fecAct - Fecha <= 3 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 2 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
   
   ' 3 Años
        
ElseIf fecAct - Fecha >= 3 * 365 And fecAct - Fecha <= 4 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 3 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
     
   ' 4 Años
ElseIf fecAct - Fecha >= 4 * 365 And fecAct - Fecha <= 5 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 4 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
  ' 5 Años
ElseIf fecAct - Fecha >= 5 * 365 And fecAct - Fecha <= 6 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 5 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
  ' 6 Años
ElseIf fecAct - Fecha >= 6 * 365 And fecAct - Fecha <= 7 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 6 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
   
   
   ' 7 Años
ElseIf fecAct - Fecha >= 7 * 365 And fecAct - Fecha <= 8 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 7 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
  ' 8 Años
ElseIf fecAct - Fecha >= 8 * 365 And fecAct - Fecha <= 9 * 365 Then
        
   adoAntiguedad.Recordset.Find "[años_acum] like " & 8 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
   
   ' 9 Años
   
ElseIf fecAct - Fecha >= 9 * 365 And fecAct - Fecha <= 10 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 9 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
  ' 10 Años
  
ElseIf fecAct - Fecha >= 10 * 365 And fecAct - Fecha <= 11 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 10 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
    
 ' 11 Años
ElseIf fecAct - Fecha >= 11 * 365 And fecAct - Fecha <= 12 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 11 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
  ' 12 Años
ElseIf fecAct - Fecha >= 12 * 365 And fecAct - Fecha <= 13 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 12 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
   
' 13 Años
  
ElseIf fecAct - Fecha >= 13 * 365 And fecAct - Fecha <= 14 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 13 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
   
' 14 Años
ElseIf fecAct - Fecha >= 14 * 365 And fecAct - Fecha <= 15 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 14 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
' 15 Años
ElseIf fecAct - Fecha >= 15 * 365 And fecAct - Fecha <= 16 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 15 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
 
 ' 16 Años
ElseIf fecAct - Fecha >= 16 * 365 And fecAct - Fecha <= 16 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 16 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
 ' 17 Años
ElseIf fecAct - Fecha >= 17 * 365 And fecAct - Fecha <= 18 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 17 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
' 18 Años
ElseIf fecAct - Fecha >= 18 * 365 And fecAct - Fecha <= 19 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 18 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
 
 ' 19 Años
 ElseIf fecAct - Fecha >= 19 * 365 And fecAct - Fecha <= 20 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 19 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
  
  ' 20 Años
 ElseIf fecAct - Fecha >= 20 * 365 Then
      
   adoAntiguedad.Recordset.Find "[años_acum] like " & 20 & ""
   BuscarIncentivo = CSng(sTotal) * adoAntiguedad.Recordset.Fields(1)
   
 Else
   BuscarIncentivo = 0
  
End If

'adoAntiguedad.RecordSource = "Departamento"
'adoAntiguedad.Refresh




End Function


Private Sub cmdCopiar_Click()

Dim sFecha1 As Variant
Dim sCodViejo As String
Dim iConta As Integer
Dim iAciertos As Integer
Dim sFecha2 As String
Dim iPeriodo As Integer
Dim iAno As Integer
Dim sMes As String
Dim sngHorasLaboradas As Single
Dim sngHorasExtras As Single
Dim sngTotalHoras As Single
Dim sngHoraLunes As Single
Dim sngHoraMartes As Single
Dim sngHoraMiercoles As Single
Dim sngHoraJueves As Single
Dim sngHoraViernes As Single
Dim sngHoraSabado As Single
Dim sngHoraDomingo As Single
Dim iDias As Integer
Dim sngInc As Single
Dim sngSeptimo As Single

iPeriodo = CInt(Me.txtPeriodo.Text)
iAno = CInt(Me.txtAno.Text)
sMes = Trim(Me.txtMes.Text)


Me.adoPrueba.CommandType = adCmdTable
Me.adoPrueba.RecordSource = "Devengado_Hora"
Me.adoPrueba.Refresh

sFecha1 = Mid$("01/08/2005", 7, 4) & "-" & Mid$("01/08/2005", 4, 2) & "-" & Mid$("01/08/2005", 1, 2)
sFecha2 = Mid$("07/08/2005", 7, 4) & "-" & Mid$("07/08/2005", 4, 2) & "-" & Mid$("07/08/2005", 1, 2)

Me.adoEmpleadoSQL.CommandType = adCmdText
Me.adoEmpleadoSQL.RecordSource = "SELECT CodEmpleado, TarifaHoraria, CodTipoNomina FROM Empleado WHERE CodTipoNomina ='02'"
Me.adoEmpleadoSQL.Refresh

Me.adoEmpleadoViejo.CommandType = adCmdText
Me.adoEmpleadoViejo.RecordSource = "SELECT * FROM Empleado"
Me.adoEmpleadoViejo.Refresh

Me.adoIngresos.CommandType = adCmdText
Me.adoIngresos.RecordSource = "SELECT * FROM Ingreso_Empl"
Me.adoIngresos.Refresh

Me.adoFechaPlanilla.CommandType = adCmdText
Me.adoFechaPlanilla.RecordSource = "SELECT * FROM Fecha_Planilla WHERE Actual =True"
Me.adoFechaPlanilla.Refresh

Me.adoAsistencia.CommandType = adCmdText
Me.adoAsistencia.RecordSource = "SELECT * FROM AsistenciaEmpleado WHERE FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND FechaSalida < = CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102)"
Me.adoAsistencia.Refresh


iPeriodo = Me.adoFechaPlanilla.Recordset.Fields("Periodo")
iAno = Me.adoFechaPlanilla.Recordset.Fields("año")
sMes = Me.adoFechaPlanilla.Recordset.Fields("mes")

Me.ospBarra.Visible = True



Do While Not Me.adoEmpleadoSQL.Recordset.EOF
   
    

sCodViejo = Mid$(Me.adoEmpleadoSQL.Recordset.Fields("CodEmpleado"), 2, 5)
iConta = 1
iAciertos = 0

Do While iConta <= 5
    
 If InStr(1, Mid$(sCodViejo, iConta, 1), "0", vbTextCompare) = 1 Then
    iAciertos = iAciertos + 1

 Else
    iConta = 6
 End If

 iConta = iConta + 1

Loop
  
sCodViejo = Mid$(sCodViejo, iAciertos + 1, Len(sCodViejo) - 1)

   Me.adoPrueba.CommandType = adCmdText
   Me.adoPrueba.RecordSource = "SELECT * FROM Devengado_Hora WHERE Cod_Empl =" & sCodViejo & " AND Periodo =" & iPeriodo & " AND mes ='" & sMes & "' AND año =" & iAno
   Me.adoPrueba.Refresh

If Me.adoPrueba.Recordset.EOF Then
   
   Me.adoAsistencia.CommandType = adCmdText
   Me.adoAsistencia.RecordSource = "SELECT * FROM AsistenciaEmpleado WHERE CodEmpleado ='" & Me.adoEmpleadoSQL.Recordset.Fields("CodEmpleado") & "' AND FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND FechaSalida < = CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) ORDER BY FechaEntrada ASC"
   Me.adoAsistencia.Refresh
   
   iDias = 1
   sngHoraLunes = 0
   sngHoraMartes = 0
   sngHoraMiercoles = 0
   sngHoraJueves = 0
   sngHoraViernes = 0
   sngHoraSabado = 0
   sngHoraDomingo = 0
   sngHorasLaboradas = 0
   sngHorasExtras = 0
   sngSeptimo = 0
   
   If Not Me.adoAsistencia.Recordset.EOF Then
   
   Do While iDias <= 7 And Not Me.adoAsistencia.Recordset.EOF
      
      Select Case iDias
      
      Case 1:
           If Me.adoAsistencia.Recordset.Fields("Dia") = "Lun" Then
              sngHoraLunes = Me.adoAsistencia.Recordset.Fields("HLaboradas") + Me.adoAsistencia.Recordset.Fields("HExtras")
              sngHorasLaboradas = Me.adoAsistencia.Recordset.Fields("HLaboradas")
              sngHorasExtras = sngHorasExtras + Me.adoAsistencia.Recordset.Fields("HExtras")
              Me.adoAsistencia.Recordset.MoveNext
           Else
              sngHoraLunes = 0
           End If
           
       Case 2:
           If Me.adoAsistencia.Recordset.Fields("Dia") = "Mart" Then
              sngHoraMartes = Me.adoAsistencia.Recordset.Fields("HLaboradas") + Me.adoAsistencia.Recordset.Fields("HExtras")
              sngHorasLaboradas = sngHorasLaboradas + Me.adoAsistencia.Recordset.Fields("HLaboradas")
              sngHorasExtras = sngHorasExtras + Me.adoAsistencia.Recordset.Fields("HExtras")
              Me.adoAsistencia.Recordset.MoveNext
           Else
              sngHoraMartes = 0
           End If
       
       Case 3:
           If Me.adoAsistencia.Recordset.Fields("Dia") = "Mierc" Then
              sngHoraMiercoles = Me.adoAsistencia.Recordset.Fields("HLaboradas") + Me.adoAsistencia.Recordset.Fields("HExtras")
              sngHorasLaboradas = sngHorasLaboradas + Me.adoAsistencia.Recordset.Fields("HLaboradas")
              sngHorasExtras = sngHorasExtras + Me.adoAsistencia.Recordset.Fields("HExtras")
              Me.adoAsistencia.Recordset.MoveNext
           Else
              sngHoraMiercoles = 0
           End If
           
       Case 4:
           If Me.adoAsistencia.Recordset.Fields("Dia") = "Juev" Then
              sngHoraJueves = Me.adoAsistencia.Recordset.Fields("HLaboradas") + Me.adoAsistencia.Recordset.Fields("HExtras")
              sngHorasLaboradas = sngHorasLaboradas + Me.adoAsistencia.Recordset.Fields("HLaboradas")
              sngHorasExtras = sngHorasExtras + Me.adoAsistencia.Recordset.Fields("HExtras")
              Me.adoAsistencia.Recordset.MoveNext
           Else
              sngHoraJueves = 0
           End If
           
       Case 5:
           If Me.adoAsistencia.Recordset.Fields("Dia") = "Viern" Then
              sngHoraViernes = Me.adoAsistencia.Recordset.Fields("HLaboradas") + Me.adoAsistencia.Recordset.Fields("HExtras")
              sngHorasLaboradas = sngHorasLaboradas + Me.adoAsistencia.Recordset.Fields("HLaboradas")
              sngHorasExtras = sngHorasExtras + Me.adoAsistencia.Recordset.Fields("HExtras")
              Me.adoAsistencia.Recordset.MoveNext
           Else
              sngHoraViernes = 0
           End If
           
       Case 6:
           If Me.adoAsistencia.Recordset.Fields("Dia") = "Sab" Then
              sngHoraSabado = Me.adoAsistencia.Recordset.Fields("HLaboradas") + Me.adoAsistencia.Recordset.Fields("HExtras")
              sngHorasLaboradas = sngHorasLaboradas + Me.adoAsistencia.Recordset.Fields("HLaboradas")
              sngHorasExtras = sngHorasExtras + Me.adoAsistencia.Recordset.Fields("HExtras")
              Me.adoAsistencia.Recordset.MoveNext
           Else
              sngHoraSabado = 0
           End If
           
       Case 7:
           If Me.adoAsistencia.Recordset.Fields("Dia") = "Dom" Then
              sngHoraDomingo = Me.adoAsistencia.Recordset.Fields("HLaboradas") + Me.adoAsistencia.Recordset.Fields("HExtras")
              sngHorasLaboradas = sngHorasLaboradas + Me.adoAsistencia.Recordset.Fields("HLaboradas")
              sngHorasExtras = sngHorasExtras + Me.adoAsistencia.Recordset.Fields("HExtras")
              Me.adoAsistencia.Recordset.MoveNext
           Else
              sngHoraDomingo = 0
           End If
           
           
           
      End Select
      
      iDias = iDias + 1
      
      
   Loop
   
   Me.adoEmpleadoViejo.CommandType = adCmdText
   Me.adoEmpleadoViejo.RecordSource = "SELECT * FROM Empleado WHERE Cod_Empl =" & CInt(sCodViejo)
   Me.adoEmpleadoViejo.Refresh
   
   sngInc = BuscarIncentivo((56 * CSng(Me.adoEmpleadoSQL.Recordset.Fields("TarifaHoraria"))), Me.adoEmpleadoViejo.Recordset.Fields("Fech_Ing"))
   
   
   If sngHorasLaboradas >= 47 Then
      sngSeptimo = Me.adoEmpleadoSQL.Recordset.Fields("TarifaHoraria") * 8
   Else
      sngSeptimo = 0
   End If
   
   
   
    Me.adoPrueba.Recordset.AddNew
      Me.adoPrueba.Recordset.Fields(0) = sCodViejo
      Me.adoPrueba.Recordset.Fields(1) = iPeriodo
      Me.adoPrueba.Recordset.Fields(2) = iAno
      Me.adoPrueba.Recordset.Fields(3) = sMes
      Me.adoPrueba.Recordset.Fields(4) = sngHoraLunes
      Me.adoPrueba.Recordset.Fields(5) = sngHoraMartes
      Me.adoPrueba.Recordset.Fields(6) = sngHoraMiercoles
      Me.adoPrueba.Recordset.Fields(7) = sngHoraJueves
      Me.adoPrueba.Recordset.Fields(8) = sngHoraViernes
      Me.adoPrueba.Recordset.Fields(9) = sngHoraSabado
      Me.adoPrueba.Recordset.Fields(10) = sngHoraDomingo
      Me.adoPrueba.Recordset.Fields(11) = CSng(Me.adoEmpleadoSQL.Recordset.Fields("TarifaHoraria")) * sngHorasLaboradas
      Me.adoPrueba.Recordset.Fields("HExtras") = Format(CSng(sngHorasExtras), "##0.##") * Me.adoEmpleadoSQL.Recordset.Fields("TarifaHoraria") * 2
      Me.adoPrueba.Recordset.Fields("Sept") = Format(CSng(sngSeptimo), "##0.##")
      Me.adoPrueba.Recordset.Fields("Antig") = Format(sngInc, "##0.##")
      Me.adoPrueba.Recordset.Fields("Transp") = 0
      Me.adoPrueba.Recordset.Fields("Alim") = 0
      Me.adoPrueba.Recordset.Fields("INSSLab") = 0
      Me.adoPrueba.Recordset.Fields("Perdida") = 0
      Me.adoPrueba.Recordset.Fields("OtrasDeducc") = 0
      Me.adoPrueba.Recordset.Fields("NHextras") = sngHorasExtras
      Me.adoPrueba.Recordset.Fields("NHOrd") = sngHorasLaboradas
      
      If sngHorasLaboradas >= 48 Then
         Me.adoPrueba.Recordset.Fields("IncPunt") = 40
      Else
         Me.adoPrueba.Recordset.Fields("IncPunt") = 0
      End If
      
      Me.adoPrueba.Recordset.Fields("HJueves") = 0
      
    Me.adoPrueba.Recordset.Update
   
     If sngHorasExtras > 0 Then
        Me.adoIngresos.Recordset.AddNew
        Me.adoIngresos.Recordset.Fields("Cod_Empl") = CInt(sCodViejo)
        Me.adoIngresos.Recordset.Fields("Periodo") = iPeriodo
        Me.adoIngresos.Recordset.Fields("Año") = iAno
        Me.adoIngresos.Recordset.Fields("mes") = sMes
        Me.adoIngresos.Recordset.Fields("Cod_Ing") = "03"
        Me.adoIngresos.Recordset.Fields("Ingreso") = Format(sngHorasExtras * 2 * Me.adoEmpleadoSQL.Recordset.Fields("TarifaHoraria"), "###.##")
        Me.adoIngresos.Recordset.Update
      End If
        
        
      If sngInc > 0 Then
        Me.adoIngresos.Recordset.AddNew
        Me.adoIngresos.Recordset.Fields("Cod_Empl") = CInt(sCodViejo)
        Me.adoIngresos.Recordset.Fields("Periodo") = iPeriodo
        Me.adoIngresos.Recordset.Fields("Año") = iAno
        Me.adoIngresos.Recordset.Fields("mes") = sMes
        Me.adoIngresos.Recordset.Fields("Cod_Ing") = "05"
        Me.adoIngresos.Recordset.Fields("Ingreso") = Format(sngInc, "##.##")
        Me.adoIngresos.Recordset.Update
      End If
        
      If sngHorasLaboradas >= 48 Then
        Me.adoIngresos.Recordset.AddNew
        Me.adoIngresos.Recordset.Fields("Cod_Empl") = CInt(sCodViejo)
        Me.adoIngresos.Recordset.Fields("Periodo") = iPeriodo
        Me.adoIngresos.Recordset.Fields("Año") = iAno
        Me.adoIngresos.Recordset.Fields("mes") = sMes
        Me.adoIngresos.Recordset.Fields("Cod_Ing") = "09"
        Me.adoIngresos.Recordset.Fields("Ingreso") = 40
        Me.adoIngresos.Recordset.Update
      End If
      
      If sngSeptimo > 0 Then
        Me.adoIngresos.Recordset.AddNew
        Me.adoIngresos.Recordset.Fields("Cod_Empl") = CInt(sCodViejo)
        Me.adoIngresos.Recordset.Fields("Periodo") = iPeriodo
        Me.adoIngresos.Recordset.Fields("Año") = iAno
        Me.adoIngresos.Recordset.Fields("mes") = sMes
        Me.adoIngresos.Recordset.Fields("Cod_Ing") = "04"
        Me.adoIngresos.Recordset.Fields("Ingreso") = Format(sngSeptimo, "##.##")
        Me.adoIngresos.Recordset.Update
      End If
      
      
        
  End If
   
 End If
   
  Me.adoEmpleadoSQL.Recordset.MoveNext
   
Loop


End Sub

Private Sub Form_Load()

Dim ruta As String
Dim conexion As Variant
Dim Server As String
Dim RutaServer As Variant

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


'Server = "Moises\Moises"

ruta = App.Path & "\PlanMetro.mdb"
conexion = "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas;Data Source=" & Server


Me.adoEmpleadoViejo.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & ruta
Me.adoIngresos.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & ruta
Me.adoPrueba.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & ruta
Me.adoFechaPlanilla.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & ruta


Me.adoPrueba.CommandType = adCmdTable
Me.adoPrueba.RecordSource = "Devengado_Hora"
Me.adoPrueba.Refresh

'sFecha1 = Mid$("01/08/2005", 7, 4) & "-" & Mid$("01/08/2005", 4, 2) & "-" & Mid$("01/08/2005", 1, 2)
'sFecha2 = Mid$("07/08/2005", 7, 4) & "-" & Mid$("07/08/2005", 4, 2) & "-" & Mid$("07/08/2005", 1, 2)

Me.adoEmpleadoSQL.CommandType = adCmdText
Me.adoEmpleadoSQL.RecordSource = "SELECT CodEmpleado, TarifaHoraria, CodTipoNomina FROM Empleado WHERE CodTipoNomina ='02'"
Me.adoEmpleadoSQL.Refresh

Me.adoEmpleadoViejo.CommandType = adCmdText
Me.adoEmpleadoViejo.RecordSource = "SELECT * FROM Empleado"
Me.adoEmpleadoViejo.Refresh

Me.adoAntiguedad.ConnectionString = conexion
Me.adoAntiguedad.CommandType = adCmdText
Me.adoAntiguedad.RecordSource = "SELECT * FROM Antiguedad"
Me.adoAntiguedad.Refresh

Me.adoIngresos.CommandType = adCmdText
Me.adoIngresos.RecordSource = "SELECT * FROM Ingreso_Empl"
Me.adoIngresos.Refresh

Me.adoFechaPlanilla.CommandType = adCmdText
Me.adoFechaPlanilla.RecordSource = "SELECT * FROM Fecha_Planilla WHERE Actual =True"
Me.adoFechaPlanilla.Refresh

Me.adoAsistencia.CommandType = adCmdText
Me.adoAsistencia.RecordSource = "SELECT CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, FechaSalida, HoraSalida, bActivo, CodTurno, HLaboradas, HExtras, bPermiso, Dia " & _
                                 "FROM AsistenciaEmpleado" ' WHERE FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND FechaSalida < = CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102)"
Me.adoAsistencia.Refresh




End Sub
