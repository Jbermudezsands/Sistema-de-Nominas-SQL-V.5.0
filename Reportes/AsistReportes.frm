VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AF8CD3F4-666F-11D1-940D-000021A73813}#5.0#0"; "osProgress.ocx"
Begin VB.Form frmRepAsistencia 
   Caption         =   "Reportes - Asistencias Empleados"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   10140
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data dtaServidor 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Width           =   4455
   End
   Begin MSAdodcLib.Adodc adoIncentivo 
      Height          =   330
      Left            =   1320
      Top             =   7200
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
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
      Connect         =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemasNominas;Data Source=METRO"
      OLEDBString     =   "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemasNominas;Data Source=METRO"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Incentivos"
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
   Begin MSAdodcLib.Adodc adoHorasExtras 
      Height          =   330
      Left            =   960
      Top             =   6600
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
      Caption         =   "Horas Extras"
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
   Begin MSAdodcLib.Adodc adoConsulta 
      Height          =   330
      Left            =   840
      Top             =   7080
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
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
      Caption         =   "Consulta"
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
   Begin MSAdodcLib.Adodc adoTurno 
      Height          =   375
      Left            =   5280
      Top             =   6720
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "Turno"
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
   Begin MSAdodcLib.Adodc adoTipoNomina 
      Height          =   375
      Left            =   5280
      Top             =   6960
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      Caption         =   "TipoNomina"
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
   Begin MSAdodcLib.Adodc adoPermisos 
      Height          =   330
      Left            =   1680
      Top             =   6720
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      Caption         =   "Permisos"
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
      Left            =   1320
      Top             =   6720
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "Asistencia"
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
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   5520
      TabIndex        =   29
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Reporte"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   28
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Calculo de Horas Labradas y Extras"
      Height          =   1815
      Left            =   240
      TabIndex        =   21
      Top             =   1200
      Width           =   9615
      Begin VB.CommandButton cmdHorasLaboradas 
         Caption         =   "&Calcular"
         Enabled         =   0   'False
         Height          =   495
         Left            =   7440
         TabIndex        =   27
         Top             =   360
         Width           =   1815
      End
      Begin Progress.osProgress ospHoras 
         Height          =   375
         Left            =   480
         TabIndex        =   26
         Top             =   1080
         Width           =   4335
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
      Begin MSComCtl2.DTPicker dtpHHasta 
         Height          =   255
         Left            =   3360
         TabIndex        =   25
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20119553
         CurrentDate     =   38612
      End
      Begin MSComCtl2.DTPicker dtpHDesde 
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20119553
         CurrentDate     =   38612
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   " al"
         Height          =   195
         Left            =   2880
         TabIndex        =   24
         Top             =   480
         Width           =   165
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   480
         TabIndex        =   22
         Top             =   480
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Nomina"
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   9615
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   7680
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cboTipoNomina 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblNoNomina 
         Caption         =   " "
         Height          =   255
         Left            =   6240
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Numero:"
         Height          =   195
         Left            =   5400
         TabIndex        =   13
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Nomina"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones del Reporte"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   9615
      Begin VB.OptionButton optLaboradas 
         Caption         =   "General Laboradas y Extras"
         Height          =   255
         Left            =   6240
         TabIndex        =   14
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton optSalidasNoRegistradas 
         Caption         =   "Salidas No Registradas"
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox cboDepto 
         Height          =   315
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox cboCargo 
         Height          =   315
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optCodigo 
         Caption         =   "Codigo de Empleado"
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   1080
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optCargo 
         Caption         =   "Por Cargo"
         Height          =   255
         Left            =   6240
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optDepto 
         Caption         =   "Departamento"
         Height          =   255
         Left            =   6240
         TabIndex        =   8
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton optFechaIngreso 
         Caption         =   "Por Fecha de Ingreso"
         Height          =   375
         Left            =   3840
         TabIndex        =   7
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton optSexo 
         Caption         =   "Por Sexo"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   255
         Left            =   3360
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20119553
         CurrentDate     =   38563
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20119553
         CurrentDate     =   38566
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Asistencia por:"
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "al"
         Height          =   195
         Left            =   3000
         TabIndex        =   4
         Top             =   480
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   480
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmRepAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ConexionRep As String



Public Function ActAnt(Fecha As Date, snProd As Single) As Single

Dim fecAct As Date
Dim sTotal As String
'Dim wsWS As Workspace
'Dim dbPlanMetro As Database
'Dim rsAnt As Recordset
'Dim rsDevHora As Recordset
Dim sCad As String
Dim BuscarIncentivo As Single
Dim snTemp As Single
Dim snFactorInss As Single
Dim cont  As Integer



Me.adoTipoNomina.CommandType = adCmdText
Me.adoTipoNomina.RecordSource = "SELECT años_acum, porcent FROM Antiguedad"
Me.adoTipoNomina.Refresh
    
'    Cont = 0
'
'    Do While Cont <= rsAnt.RecordCount - 1
'
'       snTemp = rsAnt.Fields("Ingreso")
'       rsAnt.MoveNext
'       Cont = Cont + 1
'
'    Loop
    
    
'    If Not IsNull(rsAnt.Fields("Ings")) Then
'     snTemp = rsAnt.Fields("Ings")
'    End If
'    sTotal = CStr(snProd + snTemp)
    fecAct = Format(Date, "Short Date")
'
'Set rsAnt = dbPlanMetro.OpenRecordset("Antiguedad", dbOpenDynaset)

  ' 1 Año
If fecAct - Fecha >= 365 And fecAct - Fecha <= 2 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 1 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
  ' 2 Años
    
ElseIf fecAct - Fecha >= 2 * 365 And fecAct - Fecha <= 3 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 2 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
   ' 3 Años
        
ElseIf fecAct - Fecha >= 3 * 365 And fecAct - Fecha <= 4 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 3 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
   ' 4 Años
ElseIf fecAct - Fecha >= 4 * 365 And fecAct - Fecha <= 5 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 4 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
  ' 5 Años
ElseIf fecAct - Fecha >= 5 * 365 And fecAct - Fecha <= 6 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 5 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
  ' 6 Años
ElseIf fecAct - Fecha >= 6 * 365 And fecAct - Fecha <= 7 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 6 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
   ' 7 Años
ElseIf fecAct - Fecha >= 7 * 365 And fecAct - Fecha <= 8 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 7 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
  
  ' 8 Años
ElseIf fecAct - Fecha >= 8 * 365 And fecAct - Fecha <= 9 * 365 Then
        
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 8 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
   ' 9 Años
   
ElseIf fecAct - Fecha >= 9 * 365 And fecAct - Fecha <= 10 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 9 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
  
  ' 10 Años
  
ElseIf fecAct - Fecha >= 10 * 365 And fecAct - Fecha <= 11 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 10 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
    
 ' 11 Años
ElseIf fecAct - Fecha >= 11 * 365 And fecAct - Fecha <= 12 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 11 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
  ' 12 Años
ElseIf fecAct - Fecha >= 12 * 365 And fecAct - Fecha <= 13 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 12 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
    
' 13 Años
  
ElseIf fecAct - Fecha >= 13 * 365 And fecAct - Fecha <= 14 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 13 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
    
' 14 Años
ElseIf fecAct - Fecha >= 14 * 365 And fecAct - Fecha <= 15 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 14 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
' 15 Años
ElseIf fecAct - Fecha >= 15 * 365 And fecAct - Fecha <= 16 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 15 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
 ' 16 Años
ElseIf fecAct - Fecha >= 16 * 365 And fecAct - Fecha <= 16 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 16 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
    
 ' 17 Años
ElseIf fecAct - Fecha >= 17 * 365 And fecAct - Fecha <= 18 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 17 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
    
' 18 Años
ElseIf fecAct - Fecha >= 18 * 365 And fecAct - Fecha <= 19 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 18 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
    
 ' 19 Años
 ElseIf fecAct - Fecha >= 19 * 365 And fecAct - Fecha <= 20 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 19 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
  ' 20 Años
 ElseIf fecAct - Fecha >= 20 * 365 Then
      
   Me.adoTipoNomina.Recordset.Find "[años_acum] like " & 20 & ""
   ActAnt = CSng(snProd) * Me.adoTipoNomina.Recordset.Fields(1)
   snFactorInss = Me.adoTipoNomina.Recordset.Fields(1)
   
 Else
   ActAnt = 0
  
End If

'sCad = "SELECT Antig, HExtras, NHextras, [Sal_ordin] from Devengado_hora WHERE Cod_Empl = " & sCodEmpl & " AND Periodo = " & iPer & " AND Año = " & iAnno
'
''Actualizo en la tabla devengado hora
'Set rsAnt = dbPlanMetro.OpenRecordset(sCad, dbOpenDynaset)
'
'If rsAnt.RecordCount > 0 Then
'  rsAnt.Edit
'  BuscarIncentivo = ((rsAnt.Fields("Sal_ordin") / 48) * 56) * snFactorInss
'  rsAnt.Fields("Antig") = Format(BuscarIncentivo, "##00.##")
'  rsAnt.Fields("HExtras") = 0
'  rsAnt.Fields("NHextras") = 0
'  rsAnt.Update
'End If
'
'sCad = "SELECT * FROM Ingreso_Empl WHERE [Cod_Empl] = " & sCodEmpl & " AND [Periodo] = " & iPer & " AND [Año] = " & iAnno & " AND [Cod_Ing] = '05'"
'Set rsAnt = dbPlanMetro.OpenRecordset(sCad, dbOpenDynaset)
'
'If rsAnt.RecordCount > 0 Then
'  rsAnt.Edit
'  rsAnt.Fields("Ingreso") = Format(BuscarIncentivo, "##00.##")
'  rsAnt.Update
'End If

End Function

Private Sub cboTipoNomina_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 And Me.cboTipoNomina.Text <> "" Then
   
Me.adoTipoNomina.CommandType = adCmdText
Me.adoTipoNomina.RecordSource = "SELECT TipoNomina.CodTipoNomina, TipoNomina.Nomina, Nomina.NumNomina, Nomina.FechaNominaINI, Nomina.FechaNomina, " & _
                                "Nomina.Activa FROM Nomina INNER JOIN TipoNomina ON dbo.Nomina.CodTipoNomina = dbo.TipoNomina.CodTipoNomina " & _
                                "WHERE (Nomina.Activa = 1) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "'"
Me.adoTipoNomina.Refresh

Me.dtpHDesde.Value = Me.adoTipoNomina.Recordset.Fields("FechaNominaINI")
Me.dtpHHasta.Value = Me.adoTipoNomina.Recordset.Fields("FechaNomina")
Me.dtpDesde.Value = Me.adoTipoNomina.Recordset.Fields("FechaNominaINI")
Me.dtpHasta.Value = Me.adoTipoNomina.Recordset.Fields("FechaNomina")
Me.lblNoNomina.Caption = Me.adoTipoNomina.Recordset.Fields("NumNomina")
Me.cmdHorasLaboradas.Enabled = True
Me.cmdReporte.Enabled = True


End If


End Sub

Private Sub cmdBuscar_Click()
cboTipoNomina_KeyDown 13, 0


End Sub

Private Sub cmdHorasLaboradas_Click()


Dim sngHorasLaboradas As Single
Dim sngHorasExtras As Single
Dim sngTotalHoras As Single
Dim sDia As String
Dim bUbicacion As Byte
Dim lFecha1 As Long
Dim lFecha2 As Long
Dim dFecha1 As Date
Dim dFecha2 As Date
Dim sFecha1 As String
Dim sFecha2 As String
Dim sngNumLinea As Single
Dim sngID As Single
Dim sCodEmpleado As String
Dim sngTotalHExtras As Single
Dim sngTotalHLaboradas As Single
Dim sngPagoTotal As Single
Dim sngTarifaHoraria As Single
Dim saDias(7) As String
Dim bConta As Byte

saDias(1) = "Lun"
saDias(2) = "Mart"
saDias(3) = "Mierc"
saDias(4) = "Juev"
saDias(5) = "Viern"
saDias(6) = "Sab"
saDias(7) = "Dom"

If Me.cboTipoNomina.Text = "" And Trim(Me.lblNoNomina.Caption) = "" Then
   Exit Sub
End If

dFecha1 = Format(Me.dtpHDesde.Value, "yyyy-mm-dd")
dFecha2 = Me.dtpHHasta.Value
lFecha1 = dFecha1
lFecha2 = dFecha2
sFecha1 = Mid$(Me.dtpHDesde.Value, 7, 4) & "-" & Mid$(Me.dtpHDesde.Value, 4, 2) & "-" & Mid$(Me.dtpHDesde.Value, 1, 2)
sFecha2 = Mid$(Me.dtpHHasta.Value, 7, 4) & "-" & Mid$(Me.dtpHHasta.Value, 4, 2) & "-" & Mid$(Me.dtpHHasta.Value, 1, 2)

Me.adoAsistencia.CommandType = adCmdText
Me.adoAsistencia.RecordSource = "SELECT AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.CodEmpleado1, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, " & _
                                "AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.CodTurno, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.Dia, " & _
                                "AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso, TipoNomina.Nomina FROM AsistenciaEmpleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                                "WHERE     (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND (TipoNomina.Nomina = '" & Me.cboTipoNomina.Text & " ') ORDER BY AsistenciaEmpleado.CodEmpleado ASC"
    '                            "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaSalida <> Null AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' ORDER BY AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC "
Me.adoAsistencia.Refresh

Me.adoTipoNomina.CommandType = adCmdText
Me.adoTipoNomina.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion " & _
                                 "WHERE DetalleHorasProduccion.NumNomina =" & Me.lblNoNomina.Caption
Me.adoTipoNomina.Refresh

Me.adoHorasExtras.CommandType = adCmdText
Me.adoHorasExtras.RecordSource = "SELECT Id, CodEmpleado, NumNomina, CantHoras, Pagada FROM HorasExtras WHERE NumNomina =" & Me.lblNoNomina.Caption
Me.adoHorasExtras.Refresh

Me.adoConsulta.CommandType = adCmdText
Me.adoConsulta.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion"
Me.adoConsulta.Refresh



'Me.adoConsulta.CommandType = adCmdText
'Me.adoConsulta.RecordSource = "SELECT  MAX(DetalleHorasProduccion.NumLinea) AS [MaximaLinea] FROM DetalleHorasProduccion " & _
'                               "WHERE DetalleHorasProduccion.NumNomina =" & Me.lblNoNomina.Caption
'Me.adoConsulta.Refresh

Me.adoTurno.CommandType = adCmdText
Me.adoTurno.RecordSource = "SELECT * FROM Turno"
Me.adoTurno.Refresh

Me.ospHoras.Min = 0
Me.ospHoras.Max = Me.adoAsistencia.Recordset.RecordCount
Me.ospHoras.Value = 0
Me.ospHoras.Visible = True

If Not Me.adoAsistencia.Recordset.EOF Then
   sCodEmpleado = Me.adoAsistencia.Recordset.Fields("CodEmpleado")
End If



Do While Not Me.adoAsistencia.Recordset.EOF
  
  Me.ospHoras.Value = Me.ospHoras.Value + 1
  
  If Me.adoAsistencia.Recordset.Fields("CodEmpleado1") = "002002" Then
     sDia = "Alto"
  End If
     
  sFecha1 = Mid$(Me.adoAsistencia.Recordset.Fields("FechaEntrada"), 7, 4) & "-" & Mid$(Me.adoAsistencia.Recordset.Fields("FechaEntrada"), 4, 2) & "-" & Mid$(Me.adoAsistencia.Recordset.Fields("FechaEntrada"), 1, 2)
  
  Me.adoPermisos.CommandType = adCmdText
  Me.adoPermisos.RecordSource = "SELECT * FROM Permisos WHERE CodEmpleado ='" & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & "' AND Fecha =CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND RegresoPendiente =0 AND Justificado =0"
  Me.adoPermisos.Refresh
  
  
  If Me.adoAsistencia.Recordset.Fields("FechaEntrada") = Me.adoAsistencia.Recordset.Fields("FechaSalida") Then
      Me.adoTurno.CommandType = adCmdText
      Me.adoTurno.RecordSource = "SELECT * FROM Turno WHERE CodTurno ='" & Me.adoAsistencia.Recordset.Fields("CodTurno") & "'"
      Me.adoTurno.Refresh
      
      bUbicacion = InStr(1, Format(Me.adoAsistencia.Recordset.Fields("FechaEntrada"), "Long Date"), " ", vbTextCompare)
      sDia = UCase(Mid$(Format(Me.adoAsistencia.Recordset.Fields("FechaEntrada"), "Long Date"), 1, bUbicacion - 2))
      
            
            
      Select Case sDia
     
      Case "LUNES":
          
          
          sDia = "Lun"
          
          If Me.adoAsistencia.Recordset.Fields("HoraEntrada") <= Me.adoTurno.Recordset.Fields("LEntrada") Then
            If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "12:00:00" Then
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
              sngHorasExtras = 0
            Else
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoTurno.Recordset.Fields("LSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
              sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - sngHorasLaboradas - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            End If
            
           ' sngHorasExtras = Abs((DateDiff("n", Me.adoTurno.Recordset.Fields("LSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60))
           '  If sngHorasExtras <= 0.75 Or sngHorasLaboradas < (DateDiff("n", Me.adoTurno.Recordset.Fields("LEntrada"), Me.adoTurno.Recordset.Fields("LSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60) Then
            '    sngHorasExtras = 0
            ' End If
             
'             Me.adoAsistencia.Recordset.Fields("HLaboradas") = sngHorasLaboradas
'             Me.adoAsistencia.Recordset.Fields("HExtras") = sngHorasExtras
             
          ElseIf (Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "12:00:00" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "06:30:00" And Me.adoTurno.Recordset.Fields("LSalida") >= Me.adoAsistencia.Recordset.Fields("HoraSalida")) Then  ' Or (Me.adoAsistencia.Recordset.Fields("HoraSalida") < "23:59:59" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "17:00:00") Then
             sngHorasLaboradas = (DateDiff("n", Me.adoTurno.Recordset.Fields("LEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
             sngHorasExtras = 0
                          
          Else
             sngHorasLaboradas = (DateDiff("n", Me.adoTurno.Recordset.Fields("LEntrada"), Me.adoTurno.Recordset.Fields("LSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Me.adoTurno.Recordset.Fields("LSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
             
          End If
           
          If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= Me.adoTurno.Recordset.Fields("LSalida") And Me.adoAsistencia.Recordset.Fields("HoraSalida") > "12:00:00" Then
             sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = 0
             
          Else
            
            If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "12:00:00" Then
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
              sngHorasExtras = 0
            Else
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoTurno.Recordset.Fields("LSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
              ' sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - sngHorasLaboradas
            End If
             
          End If
            
'          Me.adoAsistencia.Recordset.Fields("HLaboradas") = sngHorasLaboradas
'          Me.adoAsistencia.Recordset.Fields("HExtras") = sngHorasExtras
          'Me.adoAsistencia.Recordset.Update
         
            
         
            
       Case "MARTES":
                  
          sDia = "Mart"
          If Me.adoAsistencia.Recordset.Fields("HoraEntrada") <= Me.adoTurno.Recordset.Fields("MEntrada") Then
            If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "12:15:00" Then
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
              sngHorasExtras = 0
            Else
             sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoTurno.Recordset.Fields("MSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - sngHorasLaboradas - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            End If
            
           ' sngHorasExtras = Abs((DateDiff("n", Me.adoTurno.Recordset.Fields("MSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60))
           '  If sngHorasExtras <= 0.75 Or sngHorasLaboradas < (DateDiff("n", Me.adoTurno.Recordset.Fields("MEntrada"), Me.adoTurno.Recordset.Fields("MSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60) Then
           '     sngHorasExtras = 0
           '  End If
             
'             Me.adoAsistencia.Recordset.Fields("HLaboradas") = sngHorasLaboradas
'             Me.adoAsistencia.Recordset.Fields("HExtras") = sngHorasExtras
             
           ElseIf (Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "12:15:00" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "06:30:00" And Me.adoTurno.Recordset.Fields("MSalida") >= Me.adoAsistencia.Recordset.Fields("HoraSalida")) Then ' Or (Me.adoAsistencia.Recordset.Fields("HoraSalida") < "23:59:59" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "17:00:00") Then
             sngHorasLaboradas = (DateDiff("n", Me.adoTurno.Recordset.Fields("MEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
             sngHorasExtras = 0
             
           Else
             sngHorasLaboradas = (DateDiff("n", Me.adoTurno.Recordset.Fields("MEntrada"), Me.adoTurno.Recordset.Fields("MSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Me.adoTurno.Recordset.Fields("MSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
           End If
           
          If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= Me.adoTurno.Recordset.Fields("MSalida") And Me.adoAsistencia.Recordset.Fields("HoraSalida") > "12:00:00" Then
             sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = 0
             
          Else
            
            If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "12:15:00" Then
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
              sngHorasExtras = 0
            Else
             sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoTurno.Recordset.Fields("MSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             ' sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - sngHorasLaboradas
            End If
             
          End If
            
            
            
         
            
            
'          Me.adoAsistencia.Recordset.Fields("HLaboradas") = sngHorasLaboradas
'          Me.adoAsistencia.Recordset.Fields("HExtras") = sngHorasExtras
'          Me.adoAsistencia.Recordset.Update
           
        Case "MIÉRCOLES":
                  
          sDia = "Mierc"
          If Me.adoAsistencia.Recordset.Fields("HoraEntrada") <= Me.adoTurno.Recordset.Fields("MCEntrada") Then
            
            If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= Me.adoTurno.Recordset.Fields("MCSalida") Then
             
             If (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) < 5 Then
                sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
                sngHorasExtras = 0
             Else
                sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
                sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - sngHorasLaboradas - (Me.adoTurno.Recordset.Fields("TComida") / 60)

             End If
                
           End If
                
                
            ' sngHorasExtras = Abs((DateDiff("n", Me.adoTurno.Recordset.Fields("MCSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60))
            ' If sngHorasExtras <= 0.75 Or sngHorasLaboradas < (DateDiff("n", Me.adoTurno.Recordset.Fields("MCEntrada"), Me.adoTurno.Recordset.Fields("MCSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60) Then
            '    sngHorasExtras = 0
            ' End If
             
'             Me.adoAsistencia.Recordset.Fields("HLaboradas") = sngHorasLaboradas
'             Me.adoAsistencia.Recordset.Fields("HExtras") = sngHorasExtras
             
          ElseIf (Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "12:15:00" And Me.adoTurno.Recordset.Fields("MCSalida") >= Me.adoAsistencia.Recordset.Fields("HoraSalida")) Then  ' ) Or (Me.adoAsistencia.Recordset.Fields("HoraSalida") < "23:59:59" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "17:00:00") Then
             sngHorasLaboradas = (DateDiff("n", Me.adoTurno.Recordset.Fields("MCEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
             sngHorasExtras = 0
             
          ElseIf (Me.adoAsistencia.Recordset.Fields("HoraSalida") <= Me.adoTurno.Recordset.Fields("MCSalida")) Then
             sngHorasLaboradas = (DateDiff("n", Me.adoTurno.Recordset.Fields("MCEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60 - (Me.adoTurno.Recordset.Fields("TComida") / 60))
             sngHorasExtras = 0
          
          Else
          
             sngHorasLaboradas = (DateDiff("n", Me.adoTurno.Recordset.Fields("MCEntrada"), Me.adoTurno.Recordset.Fields("MCSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Me.adoTurno.Recordset.Fields("MCSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
             
           End If
           
          If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= Me.adoTurno.Recordset.Fields("MCSalida") And Me.adoAsistencia.Recordset.Fields("HoraSalida") > "12:00:00" Then
             sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = 0
             
          Else
          
            If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "12:15:00" Then
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
              'Me.adoTipoNomina.Recordset.Fields("NumLinea") = sngNumLinea
              sngHorasExtras = 0
            Else
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoTurno.Recordset.Fields("MCSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
           '   sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - sngHorasLaboradas
            End If
             
          End If
          
       '  End If
'          Me.adoAsistencia.Recordset.Fields("HLaboradas") = sngHorasLaboradas
'          Me.adoAsistencia.Recordset.Fields("HExtras") = sngHorasExtras
'          Me.adoAsistencia.Recordset.Update
           
       Case "JUEVES":
                  
          sDia = "Juev"
          If Me.adoAsistencia.Recordset.Fields("HoraEntrada") <= Me.adoTurno.Recordset.Fields("JEntrada") Then
            
            If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "12:15:00" Then
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
              sngHorasExtras = 0
            Else
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoTurno.Recordset.Fields("JSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
              sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - sngHorasLaboradas - (Me.adoTurno.Recordset.Fields("TComida") / 60)

            End If
            
            ' sngHorasExtras = Abs((DateDiff("n", Me.adoTurno.Recordset.Fields("JSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60))
            ' If sngHorasExtras <= 0.75 Or sngHorasLaboradas < (DateDiff("n", Me.adoTurno.Recordset.Fields("JEntrada"), Me.adoTurno.Recordset.Fields("JSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60) Then
            '    sngHorasExtras = 0
            ' End If
             
'             Me.adoAsistencia.Recordset.Fields("HLaboradas") = sngHorasLaboradas
'             Me.adoAsistencia.Recordset.Fields("HExtras") = sngHorasExtras
             
          ElseIf (Me.adoAsistencia.Recordset.Fields("HoraSalida") < "12:15:00" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "06:30:00" And Me.adoTurno.Recordset.Fields("JSalida") >= Me.adoAsistencia.Recordset.Fields("HoraSalida")) Then   ') Or (Me.adoAsistencia.Recordset.Fields("HoraSalida") < "23:59:59" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "17:00:00") Then
             sngHorasLaboradas = (DateDiff("n", Me.adoTurno.Recordset.Fields("JEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
             sngHorasExtras = 0
             
          ElseIf Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= Me.adoTurno.Recordset.Fields("JEntrada") And Me.adoTurno.Recordset.Fields("JSalida") <= Me.adoAsistencia.Recordset.Fields("HoraSalida") Then
             sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoTurno.Recordset.Fields("JSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Me.adoTurno.Recordset.Fields("JSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
          
          Else
          
             sngHorasLaboradas = (DateDiff("n", Me.adoTurno.Recordset.Fields("JEntrada"), Me.adoTurno.Recordset.Fields("JSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Me.adoTurno.Recordset.Fields("JSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
             
          End If
           
          If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= Me.adoTurno.Recordset.Fields("JSalida") And Me.adoAsistencia.Recordset.Fields("HoraSalida") > "12:15:00" Then
             sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = 0
             
          Else
            
            If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "12:15:00" Then
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
              sngHorasExtras = 0
            Else
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoTurno.Recordset.Fields("JSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
              
           '   sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) ' - sngHorasLaboradas
            End If
             
          End If
          
          
            
'          Me.adoAsistencia.Recordset.Fields("HLaboradas") = sngHorasLaboradas
'          Me.adoAsistencia.Recordset.Fields("HExtras") = sngHorasExtras
          'Me.adoAsistencia.Recordset.Update
           
        Case "VIERNES":
                  
          
          
          sDia = "Viern"
          
          If Me.adoAsistencia.Recordset.Fields("HoraEntrada") <= Me.adoTurno.Recordset.Fields("VEntrada") Then
            If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "12:15:00" Then
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
              sngHorasExtras = 0
            Else
             sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoTurno.Recordset.Fields("VSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - sngHorasLaboradas - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            End If
            
            'sngHorasExtras = Abs((DateDiff("n", Me.adoTurno.Recordset.Fields("VSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60))
            ' If sngHorasExtras <= 0.75 Or sngHorasLaboradas < (DateDiff("n", Me.adoTurno.Recordset.Fields("VEntrada"), Me.adoTurno.Recordset.Fields("VSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60) Then
            '    sngHorasExtras = 0
            ' End If
             
'             Me.adoAsistencia.Recordset.Fields("HLaboradas") = sngHorasLaboradas
'             Me.adoAsistencia.Recordset.Fields("HExtras") = sngHorasExtras
             
           ElseIf (Me.adoAsistencia.Recordset.Fields("HoraSalida") < "12:15:00" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "06:30:00" And Me.adoTurno.Recordset.Fields("VSalida") >= Me.adoAsistencia.Recordset.Fields("HoraSalida")) Then ' ) Or (Me.adoAsistencia.Recordset.Fields("HoraSalida") < "23:59:59" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "17:00:00") Then
             sngHorasLaboradas = (DateDiff("n", Me.adoTurno.Recordset.Fields("VEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
             sngHorasExtras = 0
             
           Else
             sngHorasLaboradas = (DateDiff("n", Me.adoTurno.Recordset.Fields("VEntrada"), Me.adoTurno.Recordset.Fields("VSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = (DateDiff("n", Me.adoTurno.Recordset.Fields("VSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
           End If
           
          If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= Me.adoTurno.Recordset.Fields("VSalida") And Me.adoAsistencia.Recordset.Fields("HoraSalida") > "12:00:00" Then
             sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
             sngHorasExtras = 0
             
          Else
          
            If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "12:15:00" Then
              sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
              sngHorasExtras = 0
            Else
             sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoTurno.Recordset.Fields("VSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
           '  sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - sngHorasLaboradas
            End If
             
          End If
            
          
            
            
            'If Not Me.adoPermisos.Recordset.EOF Then
'             sngHorasLaboradas = sngHorasLaboradas - (DateDiff("n", Me.adoPermisos.Recordset.Fields("HoraInicio"), Me.adoPermisos.Recordset.Fields("HoraFin")) / 60)
'             Me.adoAsistencia.Recordset.Fields("bPermiso") = 1
'          End If
'
'          Me.adoAsistencia.Recordset.Fields("HLaboradas") = sngHorasLaboradas
'          Me.adoAsistencia.Recordset.Fields("HExtras") = sngHorasExtras
'          Me.adoAsistencia.Recordset.Update
'
           
        Case Else:
                  
             If sDia = "SÁBADO" Then
                sDia = "Sab"
             Else
               sDia = "Dom"
             End If
             
          
           If (Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "12:15:00" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "05:00:00") Or (Me.adoAsistencia.Recordset.Fields("HoraSalida") < "23:59:59" And Me.adoAsistencia.Recordset.Fields("HoraEntrada") >= "17:00:00") Then
              sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
           Else
              sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
           End If
             'Me.adoAsistencia.Recordset.Fields("HLaboradas") = 0
             sngHorasLaboradas = 0
             'Me.adoAsistencia.Recordset.Fields("HExtras") = sngHorasExtras
          
            
          'Me.adoAsistencia.Recordset.Update
           
           
           
            
            
      End Select
     
         
     'Me.lblTotalEmpl.Caption = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida") / 60)) - (Me.adoAsistencia.Recordset.Fields("TComida") / 60)
     
     
     
  Else
  
  
      Me.adoTurno.CommandType = adCmdText
      Me.adoTurno.RecordSource = "SELECT * FROM Turno WHERE CodTurno ='" & Me.adoAsistencia.Recordset.Fields("CodTurno") & "'"
      Me.adoTurno.Refresh
     
      bUbicacion = InStr(1, Format(Me.adoAsistencia.Recordset.Fields("FechaEntrada"), "Long Date"), " ", vbTextCompare)
      sDia = Mid$(Format(Me.adoAsistencia.Recordset.Fields("FechaEntrada"), "Long Date"), 1, bUbicacion - 2)
  
      Select Case sDia
      
      Case "SÁBADO":
         
         sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida") / 60)) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
         sngHorasLaboradas = 0
                
              
  
      Case "DOMINGO":
         
         sngHorasExtras = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida") / 60)) - (Me.adoTurno.Recordset.Fields("TComida") / 60)
         sngHorasLaboradas = 0
               
      Case "LUNES":
         
         If Me.adoTurno.Recordset.Fields("LSalida") <= Me.adoAsistencia.Recordset.Fields("HoraSalida") Then
            sngHorasLaboradas = Format((DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Me.adoTurno.Recordset.Fields("LSalida")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            sngHorasExtras = DateDiff("n", Me.adoTurno.Recordset.Fields("LSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60
         Else
             
            If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "00:00:00" Then
               sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
               sngHorasExtras = 0
            Else
               sngHorasLaboradas = Format((DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
               sngHorasExtras = 0
            End If
         
         End If
         
            
         
         
       Case "MARTES":
         
         If Me.adoTurno.Recordset.Fields("MSalida") <= Me.adoAsistencia.Recordset.Fields("HoraSalida") Then
            sngHorasLaboradas = Format((DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Me.adoTurno.Recordset.Fields("MSalida")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            sngHorasExtras = DateDiff("n", Me.adoTurno.Recordset.Fields("MSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60
         Else
             
            If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "00:00:00" Then
               sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
               sngHorasExtras = 0
            Else
               sngHorasLaboradas = Format((DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
               sngHorasExtras = 0
            End If
         
         End If
         
       
         
         
       Case "MIÉRCOLES":
         
         If Me.adoTurno.Recordset.Fields("MCSalida") <= Me.adoAsistencia.Recordset.Fields("HoraSalida") Then
            sngHorasLaboradas = Format((DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Me.adoTurno.Recordset.Fields("MCSalida")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            sngHorasExtras = DateDiff("n", Me.adoTurno.Recordset.Fields("MCSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60
         Else
             
            If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "00:00:00" Then
               sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
               sngHorasExtras = 0
            Else
               sngHorasLaboradas = Format((DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
               sngHorasExtras = 0
            End If
         
         End If
         
      
         
         
       Case "JUEVES":
         
         If Me.adoTurno.Recordset.Fields("JSalida") <= Me.adoAsistencia.Recordset.Fields("HoraSalida") Then
            sngHorasLaboradas = Format((DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Me.adoTurno.Recordset.Fields("JSalida")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            sngHorasExtras = DateDiff("n", Me.adoTurno.Recordset.Fields("JSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60
         Else
             
            If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "00:00:00" Then
               sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
               sngHorasExtras = 0
            Else
               sngHorasLaboradas = Format((DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
               sngHorasExtras = 0
            End If
         
         End If
         
     
         
         
       Case "VIERNES":
         
         If Me.adoTurno.Recordset.Fields("VSalida") <= Me.adoAsistencia.Recordset.Fields("HoraSalida") Then
            sngHorasLaboradas = Format((DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Me.adoTurno.Recordset.Fields("VSalida")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
            sngHorasExtras = DateDiff("n", Me.adoTurno.Recordset.Fields("VSalida"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60
         Else
             
            If Me.adoAsistencia.Recordset.Fields("HoraSalida") <= "00:00:00" Then
               sngHorasLaboradas = (DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60)
               sngHorasExtras = 0
            Else
               sngHorasLaboradas = Format((DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "23:59:59") / 60) + DateDiff("n", "00:00:00", Me.adoAsistencia.Recordset.Fields("HoraSalida")) / 60) ' - (Me.adoTurno.Recordset.Fields("TComida") / 60)
               sngHorasExtras = 0
            End If
         
         End If
             
                   
      End Select
       
       
      
       
'       lDiaAnterior = DateDiff("n", Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "23:59:59") / 60
'       lDiaSiguiente = (lDiaAnterior + DateDiff("n", "00:00:00", Me.adoAsistencia.Recordset.Fields("HoraSalida") / 60)) - (Me.adoAsistencia.Recordset.Fields("TComida") / 60)
'       Me.lblTotalEmpl.Caption = lDiaSiguiente
     
  End If
  
  
  
  If Not Me.adoPermisos.Recordset.EOF Then
     sngHorasLaboradas = sngHorasLaboradas - (DateDiff("n", Me.adoPermisos.Recordset.Fields("HoraInicio"), Me.adoPermisos.Recordset.Fields("HoraFin")) / 60)
     Me.adoAsistencia.Recordset.Fields("bPermiso") = 1
  End If
         
  Me.adoPermisos.CommandType = adCmdText
  Me.adoPermisos.RecordSource = "SELECT * FROM ExtraTurno WHERE CodEmpleado ='" & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & "' AND FechaEntrada =CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND (bActivo = 0) AND (NOT (HoraSalida IS NULL)) AND (NOT (FechaEntrada IS NULL))"
  Me.adoPermisos.Refresh
          
  If Not Me.adoPermisos.Recordset.EOF And Not IsNull(Me.adoPermisos.Recordset.Fields("HorasLaboradas")) Then
     sngHorasExtras = sngHorasExtras + Me.adoPermisos.Recordset.Fields("HorasLaboradas")
  End If
       
  If sngHorasLaboradas <> 0 Then
     sngHorasLaboradas = Format(sngHorasLaboradas, "##.##")
  Else
     sngHorasLaboradas = 0
  End If
  
  If sngHorasExtras <> 0 Then
     sngHorasExtras = Format(sngHorasExtras, "##.##")
  Else
     sngHorasExtras = 0
  End If
  
  
  Me.adoAsistencia.Recordset.Fields("Dia") = sDia
  Me.adoAsistencia.Recordset.Fields("HLaboradas") = sngHorasLaboradas
  Me.adoAsistencia.Recordset.Fields("HExtras") = sngHorasExtras
  Me.adoAsistencia.Recordset.Update
  
  
  Me.adoAsistencia.Recordset.MoveNext
   
Loop
     
sFecha1 = Mid$(Me.dtpHDesde.Value, 7, 4) & "-" & Mid$(Me.dtpHDesde.Value, 4, 2) & "-" & Mid$(Me.dtpHDesde.Value, 1, 2)
sFecha2 = Mid$(Me.dtpHHasta.Value, 7, 4) & "-" & Mid$(Me.dtpHHasta.Value, 4, 2) & "-" & Mid$(Me.dtpHHasta.Value, 1, 2)

Me.adoAsistencia.CommandType = adCmdText
Me.adoAsistencia.RecordSource = "SELECT Empleado.CodEmpleado, Empleado.TarifaHoraria, Empleado.SalarioMinimo, TipoNomina.Nomina, Empleado.Activo FROM Empleado INNER JOIN TipoNomina ON dbo.Empleado.CodTipoNomina = dbo.TipoNomina.CodTipoNomina " & _
                                "WHERE (TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "') AND (dbo.Empleado.Activo = 1) ORDER BY CodEmpleado ASC"
                                
Me.adoAsistencia.Refresh
   
Me.adoTipoNomina.CommandType = adCmdText
Me.adoTipoNomina.RecordSource = "SELECT AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, " & _
                                "AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.CodTurno, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.Dia, " & _
                                "AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso, TipoNomina.Nomina FROM AsistenciaEmpleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                                "WHERE (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND (TipoNomina.Nomina = '" & Me.cboTipoNomina.Text & " ') ORDER BY AsistenciaEmpleado.CodEmpleado ASC"
Me.adoTipoNomina.Refresh
   
   
Me.ospHoras.Value = 0
Me.ospHoras.Min = 0
Me.ospHoras.Max = Me.adoAsistencia.Recordset.RecordCount
Me.ospHoras.Value = 0


bConta = 1
     
     
     
Do While Not Me.adoAsistencia.Recordset.EOF
  
   Me.ospHoras.Value = Me.ospHoras.Value + 1
  
  Me.adoConsulta.CommandType = adCmdText
  Me.adoConsulta.RecordSource = "SELECT Max(Id) AS [MaximaId] FROM HorasExtras"
  Me.adoConsulta.Refresh
    
  If Not IsNull(Me.adoConsulta.Recordset.Fields("MaximaId")) Then
     sngID = Me.adoConsulta.Recordset.Fields("MaximaId") + 1
  Else
     sngID = 1
  End If
  
  If Me.adoAsistencia.Recordset.Fields("CodEmpleado") = "3627" Then
     sDia = "Alto"
  End If
  
  
  Me.adoTipoNomina.CommandType = adCmdText
  Me.adoTipoNomina.RecordSource = "SELECT Sum(AsistenciaEmpleado.HExtras) AS [SumaExtras] FROM AsistenciaEmpleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                                 "WHERE (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND (TipoNomina.Nomina = '" & Me.cboTipoNomina.Text & "') " & _
                                 "AND CodEmpleado ='" & Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado")) & "'"
  Me.adoTipoNomina.Refresh
  
  
  If Not IsNull(Me.adoTipoNomina.Recordset.Fields("SumaExtras")) Then
    If Me.adoTipoNomina.Recordset.Fields("SumaExtras") <> 0 Then
     sngTotalHExtras = Format(Me.adoTipoNomina.Recordset.Fields("SumaExtras"), "##.##")
    Else
     sngTotalHExtras = 0
    End If
    
  Else
     sngTotalHExtras = 0
  End If
  
 
  Me.adoConsulta.CommandType = adCmdText
  Me.adoConsulta.RecordSource = "SELECT Id, CodEmpleado, NumNomina, CantHoras, Pagada FROM HorasExtras " & _
                               "WHERE NumNomina =" & Me.lblNoNomina.Caption & " AND CodEmpleado ='" & Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado")) & "'"
  Me.adoConsulta.Refresh
            
  If Not Me.adoConsulta.Recordset.EOF Then
    If sngTotalHExtras <> 0 Then
     Me.adoConsulta.Recordset.Fields("CantHoras") = Format(sngTotalHExtras, "##.##")
     Me.adoConsulta.Recordset.Fields("Pagada") = 0
     Me.adoConsulta.Recordset.Update
    Else
     Me.adoConsulta.Recordset.Fields("CantHoras") = 0
     Me.adoConsulta.Recordset.Fields("Pagada") = 0
     Me.adoConsulta.Recordset.Update
    
    End If
    
     Me.adoConsulta.Refresh
     
  ElseIf sngTotalHExtras <> 0 Then
     Me.adoConsulta.Recordset.AddNew
     Me.adoConsulta.Recordset.Fields("Id") = sngID
     Me.adoConsulta.Recordset.Fields("CodEmpleado") = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
     Me.adoConsulta.Recordset.Fields("NumNomina") = Me.lblNoNomina.Caption
     If sngTotalHExtras <> 0 Then
        Me.adoConsulta.Recordset.Fields("CantHoras") = Format(sngTotalHExtras, "##.##")
     Else
        Me.adoConsulta.Recordset.Fields("CantHoras") = 0
     End If
     
     Me.adoConsulta.Recordset.Fields("Pagada") = 0
     Me.adoConsulta.Recordset.Update
     Me.adoConsulta.Refresh
            
  Else
  
     Me.adoConsulta.Recordset.AddNew
     Me.adoConsulta.Recordset.Fields("Id") = sngID
     Me.adoConsulta.Recordset.Fields("CodEmpleado") = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
     Me.adoConsulta.Recordset.Fields("NumNomina") = Me.lblNoNomina.Caption
     
     Me.adoConsulta.Recordset.Fields("CantHoras") = 0
    
     
     Me.adoConsulta.Recordset.Fields("Pagada") = 0
     Me.adoConsulta.Recordset.Update
     Me.adoConsulta.Refresh
     
  
            
  End If
  
  
  
Me.adoAsistencia.Recordset.MoveNext

Loop

Me.adoAsistencia.CommandType = adCmdText
Me.adoAsistencia.RecordSource = "SELECT Empleado.CodEmpleado, Empleado.TarifaHoraria, TipoNomina.Nomina, AsistenciaEmpleado.FechaEntrada, " & _
                                "AsistenciaEmpleado.FechaSalida FROM AsistenciaEmpleado INNER JOIN Empleado ON dbo.AsistenciaEmpleado.CodEmpleado = dbo.Empleado.CodEmpleado INNER JOIN " & _
                                "TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                                 "WHERE (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND (TipoNomina.Nomina = '" & Me.cboTipoNomina.Text & " ') ORDER BY AsistenciaEmpleado.CodEmpleado ASC"
Me.adoAsistencia.Refresh
   
Me.adoTipoNomina.CommandType = adCmdText
Me.adoTipoNomina.RecordSource = "SELECT AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, " & _
                                "AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.CodTurno, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.Dia, " & _
                                "AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso FROM AsistenciaEmpleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                                "WHERE (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) ORDER BY AsistenciaEmpleado.CodEmpleado ASC"
Me.adoTipoNomina.Refresh
   
Me.adoHorasExtras.CommandType = adCmdText
Me.adoHorasExtras.RecordSource = "SELECT AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, " & _
                                "AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.CodTurno, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.Dia, " & _
                                "AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso FROM AsistenciaEmpleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                                "WHERE (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) ORDER BY AsistenciaEmpleado.CodEmpleado ASC"
Me.adoHorasExtras.Refresh
   
   
   
   
Me.ospHoras.Value = 0
Me.ospHoras.Min = 0
Me.ospHoras.Max = Me.adoAsistencia.Recordset.RecordCount + 1
Me.ospHoras.Value = 0

If Me.cboTipoNomina.Text = "Produccion" Then

Do While Not Me.adoAsistencia.Recordset.EOF

bConta = 1

If Me.adoAsistencia.Recordset.Fields("CodEmpleado") = "4296" Then
  bConta = 1
End If





Me.adoHorasExtras.CommandType = adCmdText
Me.adoHorasExtras.RecordSource = "SELECT AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, " & _
                                "AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.CodTurno, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.Dia, " & _
                                "AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso FROM AsistenciaEmpleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                                "WHERE (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) " & _
                                "AND AsistenciaEmpleado.CodEmpleado ='" & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & "' ORDER BY AsistenciaEmpleado.CodEmpleado ASC"
Me.adoHorasExtras.Refresh

sngTotalHLaboradas = 0
sngPagoTotal = 0
Me.ospHoras.Value = Me.ospHoras.Value + 1


Do While bConta <= 7 And Not Me.adoHorasExtras.Recordset.EOF

  
  
  Me.adoConsulta.CommandType = adCmdText
  Me.adoConsulta.RecordSource = "SELECT Max(NumLinea) AS [MaximaLinea] FROM DetalleHorasProduccion"
  Me.adoConsulta.Refresh
    
  If Not IsNull(Me.adoConsulta.Recordset.Fields("MaximaLinea")) Then
     sngNumLinea = Me.adoConsulta.Recordset.Fields("MaximaLinea") + 1
  Else
     sngNumLinea = 1
  End If
  
  
  If Not Me.adoAsistencia.Recordset.EOF Then
  
     Me.adoTipoNomina.CommandType = adCmdText
     Me.adoTipoNomina.RecordSource = "SELECT Sum(AsistenciaEmpleado.Hlaboradas) AS [SumaLaboradas] FROM AsistenciaEmpleado " & _
                                  "WHERE (AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00', 102)) AND (AsistenciaEmpleado.FechaSalida IS NOT NULL) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND CodEmpleado ='" & Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado")) & "' " & _
                                  "AND Dia ='" & Me.adoHorasExtras.Recordset.Fields("Dia") & "'"
     Me.adoTipoNomina.Refresh
  
  
   '  If Me.adoAsistencia.Recordset.Fields("CodEmpleado") = "000761" Then
   '     sDia = "Alto"
   '  End If
   '
  
  
    Select Case Me.adoHorasExtras.Recordset.Fields("Dia")
  
    Case "Lun":
       
       
       If Not IsNull(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas")) Then
         If Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > 0 And Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") <= 9.75 Then
            sngHorasLaboradas = Format(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas"), "##.##")
         
         ElseIf Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > 9.75 Then
            sngHorasLaboradas = 9.75
         
         Else
            sngHorasLaboradas = 0
            
         End If
    
      Else
         sngHorasLaboradas = 0
      End If
       
      Me.adoConsulta.CommandType = adCmdText
      Me.adoConsulta.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion " & _
                                 "WHERE  DetalleHorasProduccion.CodEmpleado ='" & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & "' AND DetalleHorasProduccion.NumNomina =" & Trim(Me.lblNoNomina.Caption)
      Me.adoConsulta.Refresh
            
     If Not Me.adoConsulta.Recordset.EOF Then
       If sngHorasLaboradas <> 0 Then
         Me.adoConsulta.Recordset.Fields("Lunes") = Format(sngHorasLaboradas, "##.##")
       Else
         Me.adoConsulta.Recordset.Fields("Lunes") = 0
       End If
    
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
     Else
       Me.adoConsulta.Recordset.AddNew
       Me.adoConsulta.Recordset.Fields("NumLinea") = sngNumLinea
       Me.adoConsulta.Recordset.Fields("CodEmpleado") = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
       Me.adoConsulta.Recordset.Fields("NumNomina") = Trim(Me.lblNoNomina.Caption)
       If sngHorasLaboradas <> 0 Then
          Me.adoConsulta.Recordset.Fields("Lunes") = Format(sngHorasLaboradas, "##.##")
       Else
          Me.adoConsulta.Recordset.Fields("Lunes") = 0
       End If
     
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
            
            
     End If
       
       
    Case "Mart":
              
       If Not IsNull(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas")) Then
         If Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > 0 And Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") <= 9.75 Then
            sngHorasLaboradas = Format(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas"), "##.##")
            
         ElseIf Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > 9.75 Then
            sngHorasLaboradas = 9.75
            
         Else
            sngHorasLaboradas = 0
         End If
    
      Else
         sngHorasLaboradas = 0
      End If
       
      Me.adoConsulta.CommandType = adCmdText
      Me.adoConsulta.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion " & _
                                 "WHERE  DetalleHorasProduccion.CodEmpleado ='" & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & "' AND DetalleHorasProduccion.NumNomina =" & Me.lblNoNomina.Caption
      Me.adoConsulta.Refresh
            
     If Not Me.adoConsulta.Recordset.EOF Then
       If sngHorasLaboradas <> 0 Then
         Me.adoConsulta.Recordset.Fields("Martes") = Format(sngHorasLaboradas, "##.##")
       Else
         Me.adoConsulta.Recordset.Fields("Martes") = 0
       End If
    
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
     Else
       Me.adoConsulta.Recordset.AddNew
       Me.adoConsulta.Recordset.Fields("NumLinea") = sngNumLinea
       Me.adoConsulta.Recordset.Fields("CodEmpleado") = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
       Me.adoConsulta.Recordset.Fields("NumNomina") = Me.lblNoNomina.Caption
       If sngTotalHExtras <> 0 Then
          Me.adoConsulta.Recordset.Fields("Martes") = Format(sngHorasLaboradas, "##.##")
       Else
          Me.adoConsulta.Recordset.Fields("Martes") = 0
       End If
     
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
            
            
     End If
       
       
       
     Case "Mierc":
       
       If Not IsNull(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas")) Then
         If Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > 0 And Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") <= 9.75 Then
            sngHorasLaboradas = Format(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas"), "##.##")
            
         ElseIf Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > 9.75 Then
            sngHorasLaboradas = 9.75
            
         Else
            sngHorasLaboradas = 0
         End If
    
      Else
         sngHorasLaboradas = 0
      End If
       
      Me.adoConsulta.CommandType = adCmdText
      Me.adoConsulta.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion " & _
                                 "WHERE  DetalleHorasProduccion.CodEmpleado ='" & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & "' AND DetalleHorasProduccion.NumNomina =" & Me.lblNoNomina.Caption
      Me.adoConsulta.Refresh
            
     If Not Me.adoConsulta.Recordset.EOF Then
       If sngHorasLaboradas <> 0 Then
         Me.adoConsulta.Recordset.Fields("Miercoles") = Format(sngHorasLaboradas, "##.##")
       Else
         Me.adoConsulta.Recordset.Fields("Miercoles") = 0
       End If
    
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
     Else
       Me.adoConsulta.Recordset.AddNew
       Me.adoConsulta.Recordset.Fields("NumLinea") = sngNumLinea
       Me.adoConsulta.Recordset.Fields("CodEmpleado") = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
       Me.adoConsulta.Recordset.Fields("NumNomina") = Me.lblNoNomina.Caption
       If sngTotalHExtras <> 0 Then
          Me.adoConsulta.Recordset.Fields("Miercoles") = Format(sngHorasLaboradas, "##.##")
       Else
          Me.adoConsulta.Recordset.Fields("Miercoles") = 0
       End If
     
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
            
            
     End If
       
     Case "Juev":
       
       If Not IsNull(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas")) Then
         If Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > 0 And Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") <= 9.75 Then
            sngHorasLaboradas = Format(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas"), "##.##")
            
         ElseIf Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > 9.75 Then
            sngHorasLaboradas = 9.75
         Else
            sngHorasLaboradas = 0
         End If
    
      Else
         sngHorasLaboradas = 0
      End If
       
      Me.adoConsulta.CommandType = adCmdText
      Me.adoConsulta.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion " & _
                                 "WHERE  DetalleHorasProduccion.CodEmpleado ='" & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & "' AND DetalleHorasProduccion.NumNomina =" & Me.lblNoNomina.Caption
      Me.adoConsulta.Refresh
            
     If Not Me.adoConsulta.Recordset.EOF Then
       If sngHorasLaboradas <> 0 Then
         Me.adoConsulta.Recordset.Fields("Jueves") = Format(sngHorasLaboradas, "##.##")
       Else
         Me.adoConsulta.Recordset.Fields("Jueves") = 0
       End If
    
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
     Else
       Me.adoConsulta.Recordset.AddNew
       Me.adoConsulta.Recordset.Fields("NumLinea") = sngNumLinea
       Me.adoConsulta.Recordset.Fields("CodEmpleado") = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
       Me.adoConsulta.Recordset.Fields("NumNomina") = Me.lblNoNomina.Caption
       If sngTotalHExtras <> 0 Then
         If sngHorasLaboradas = 0 Then
          Me.adoConsulta.Recordset.Fields("Jueves") = 0
         Else
          Me.adoConsulta.Recordset.Fields("Jueves") = Format(sngHorasLaboradas, "##.##")
         End If
       Else
          Me.adoConsulta.Recordset.Fields("Jueves") = 0
       End If
     
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
            
            
     End If
       
       
     Case "Viern":
       
       If Not IsNull(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas")) Then
         If Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > 0 And Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") <= 9 Then
            sngHorasLaboradas = Format(Me.adoTipoNomina.Recordset.Fields("SumaLaboradas"), "##.##")
            
         ElseIf Me.adoTipoNomina.Recordset.Fields("SumaLaboradas") > 9 Then
            sngHorasLaboradas = 9
            
         Else
            sngHorasLaboradas = 0
         End If
    
      Else
         sngHorasLaboradas = 0
      End If
       
      Me.adoConsulta.CommandType = adCmdText
      Me.adoConsulta.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion " & _
                                 "WHERE  DetalleHorasProduccion.CodEmpleado ='" & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & "' AND DetalleHorasProduccion.NumNomina =" & Me.lblNoNomina.Caption
      Me.adoConsulta.Refresh
            
     If Not Me.adoConsulta.Recordset.EOF Then
       If sngHorasLaboradas <> 0 Then
         Me.adoConsulta.Recordset.Fields("Viernes") = Format(sngHorasLaboradas, "##.##")
       Else
         Me.adoConsulta.Recordset.Fields("Viernes") = 0
       End If
    
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
     Else
       Me.adoConsulta.Recordset.AddNew
       Me.adoConsulta.Recordset.Fields("NumLinea") = sngNumLinea
       Me.adoConsulta.Recordset.Fields("CodEmpleado") = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
       Me.adoConsulta.Recordset.Fields("NumNomina") = Me.lblNoNomina.Caption
       If sngTotalHExtras <> 0 Then
         If sngHorasLaboradas = 0 Then
          Me.adoConsulta.Recordset.Fields("Viernes") = 0
         Else
          Me.adoConsulta.Recordset.Fields("Viernes") = Format(sngHorasLaboradas, "##.##")
         End If
         
       Else
          Me.adoConsulta.Recordset.Fields("Viernes") = 0
       End If
     
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
            
            
     End If
       
       
       
     End Select
         
   If Me.adoHorasExtras.Recordset.Fields("Dia") <> "Sab" And Me.adoHorasExtras.Recordset.Fields("Dia") <> "Dom" Then
         
     sCodEmpleado = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
     sngTarifaHoraria = Me.adoAsistencia.Recordset.Fields("TarifaHoraria")
     sngTotalHLaboradas = sngTotalHLaboradas + sngHorasLaboradas
              
     Me.adoConsulta.CommandType = adCmdText
     Me.adoConsulta.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes, " & _
                                 "DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, DetalleHorasProduccion.Jueves, DetalleHorasProduccion.Viernes, " & _
                                 "DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo, DetalleHorasProduccion.TotalHoras, DetalleHorasProduccion.SalarioHora, " & _
                                 "DetalleHorasProduccion.TotalSalarioHora , DetalleHorasProduccion.Pagado FROM DetalleHorasProduccion " & _
                                 "WHERE  DetalleHorasProduccion.CodEmpleado ='" & sCodEmpleado & "' AND DetalleHorasProduccion.NumNomina =" & Trim(Me.lblNoNomina.Caption)
      Me.adoConsulta.Refresh
            
     If Not Me.adoConsulta.Recordset.EOF Then
       If sngTotalHLaboradas <> 0 And Me.adoConsulta.Recordset.Fields("Lunes") + Me.adoConsulta.Recordset.Fields("Martes") + Me.adoConsulta.Recordset.Fields("Miercoles") + Me.adoConsulta.Recordset.Fields("Jueves") + Me.adoConsulta.Recordset.Fields("Viernes") <> 0 Then
         Me.adoConsulta.Recordset.Fields("TotalHoras") = Format(Me.adoConsulta.Recordset.Fields("Lunes") + Me.adoConsulta.Recordset.Fields("Martes") + Me.adoConsulta.Recordset.Fields("Miercoles") + Me.adoConsulta.Recordset.Fields("Jueves") + Me.adoConsulta.Recordset.Fields("Viernes"), "##.##")
       Else
         Me.adoConsulta.Recordset.Fields("TotalHoras") = 0
       End If
       
       Me.adoConsulta.Recordset.Fields("SalarioHora") = sngTarifaHoraria
       
         If Me.adoAsistencia.Recordset.Fields("TarifaHoraria") = 0 Then
            Me.adoConsulta.Recordset.Fields("TotalSalarioHora") = 0
         Else
           If Me.adoAsistencia.Recordset.Fields("TarifaHoraria") * sngTotalHLaboradas = 0 Then
            Me.adoConsulta.Recordset.Fields("TotalSalarioHora") = 0
           ElseIf Me.adoConsulta.Recordset.Fields("Lunes") + Me.adoConsulta.Recordset.Fields("Martes") + Me.adoConsulta.Recordset.Fields("Miercoles") + Me.adoConsulta.Recordset.Fields("Jueves") + Me.adoConsulta.Recordset.Fields("Viernes") <> 0 Then
            sngTotalHLaboradas = Me.adoConsulta.Recordset.Fields("Lunes") + Me.adoConsulta.Recordset.Fields("Martes") + Me.adoConsulta.Recordset.Fields("Miercoles") + Me.adoConsulta.Recordset.Fields("Jueves") + Me.adoConsulta.Recordset.Fields("Viernes")
            Me.adoConsulta.Recordset.Fields("TotalSalarioHora") = Format(sngTarifaHoraria * sngTotalHLaboradas, "##,###.##")
           End If
           
         End If
         
       sngPagoTotal = Me.adoConsulta.Recordset.Fields("TotalSalarioHora")
       Me.adoConsulta.Recordset.Fields("Pagado") = 0
       Me.adoConsulta.Recordset.Update
       Me.adoConsulta.Refresh
       
      End If
              
    End If
              
    Me.adoAsistencia.Recordset.MoveNext
    
    End If
    
     sngHorasLaboradas = 0
     Me.adoHorasExtras.Recordset.MoveNext
     
     bConta = bConta + 1
     
     
    Loop
        
'   Me.adoIncentivo.CommandType = adCmdText
'   Me.adoIncentivo.RecordSource = "SELECT Max(NumIncentivo) AS [MaxNumIncentivo] FROM Incentivo"
'   Me.adoIncentivo.Refresh
'
'   If Not IsNull(Me.adoIncentivo.Recordset.Fields("MaxNumIncentivo")) Then
'      sngNumLinea = Me.adoIncentivo.Recordset.Fields("MAxNumIncentivo") + 1
'   Else
'      sngNumLinea = sngNumLinea + 1
'   End If
'
'   Me.adoIncentivo.CommandType = adCmdText
'   Me.adoIncentivo.RecordSource = "SELECT NumIncentivo, CodEmpleado, CodTipoIncentivo, NumVeces, Pagado FROM Incentivo " & _
'                                  "WHERE CodEmpleado ='" & sCodEmpleado & "' AND CodTipoIncentivo ='01'"
'   Me.adoIncentivo.Refresh
'
'   If Me.adoIncentivo.Recordset.EOF Then
'      Me.adoIncentivo.Recordset.AddNew
'      Me.adoIncentivo.Recordset.Fields("NumIncentivo") = sngNumLinea
'      Me.adoIncentivo.Recordset.Fields("CodEmpleado") = sCodEmpleado
'      Me.adoIncentivo.Recordset.Fields("CodTipoIncentivo") = "01"
'      Me.adoIncentivo.Recordset.Fields("NumVeces") = "n"
'      Me.adoIncentivo.Recordset.Fields("Pagado") = 0
'      Me.adoIncentivo.Recordset.Update
'   Else
'
'      sngNumLinea = Me.adoIncentivo.Recordset.Fields("NumIncentivo")
'
'   End If
'
'   Me.adoIncentivo.CommandType = adCmdText
'   Me.adoIncentivo.RecordSource = "SELECT Id, NumIncentivo, Valor, NumVez, Pagado, NumNomina FROM DetalleIncentivo"
'
'   Me.adoIncentivo.Refresh
'
'   Me.adoIncentivo.CommandType = adCmdText
'   Me.adoIncentivo.RecordSource = "SELECT Max(Id) AS [MaxNumId] FROM DetalleIncentivo"
'
'   Me.adoIncentivo.Refresh
'
'   If Not IsNull(Me.adoIncentivo.Recordset.Fields("MaxNumId")) Then
'      sngID = Me.adoIncentivo.Recordset.Fields("MaxNumId") + 1
'   Else
'      sngID = 1
'   End If
'
'   Me.adoIncentivo.CommandType = adCmdText
'   Me.adoIncentivo.RecordSource = "SELECT Id, NumIncentivo, Valor, NumVez, Pagado, NumNomina FROM DetalleIncentivo " & _
'                                  "WHERE NumIncentivo =" & sngNumLinea & " AND NumNomina =" & Trim(Me.lblNoNomina.Caption)
'   Me.adoIncentivo.Refresh
'
'   Me.adoConsulta.CommandType = adCmdText
'   Me.adoConsulta.RecordSource = "SELECT Id, Codempleado, FechaContrato FROM Historico " & _
'                                 "WHERE  CodEmpleado ='" & sCodEmpleado & "'"
'   Me.adoConsulta.Refresh
'
'
'
'   If Not Me.adoIncentivo.Recordset.EOF And Not Me.adoConsulta.Recordset.EOF Then
'      Me.adoIncentivo.Recordset.Fields("Valor") = ActAnt(Me.adoConsulta.Recordset.Fields("FechaContrato"), 48 * sngTarifaHoraria)
'      Me.adoIncentivo.Recordset.Update
'      Me.adoIncentivo.Refresh
'   Else
'      If Me.adoConsulta.Recordset.EOF Then
'         MsgBox "El empleado " & sCodEmpleado & ", no tiene registrado su fecha de contrato, no se puede obtener su incentivo por antiguedad"
'      ElseIf Me.adoIncentivo.Recordset.EOF Then
'         Me.adoIncentivo.Recordset.AddNew
'         Me.adoIncentivo.Recordset.Fields("Id") = sngID
'         Me.adoIncentivo.Recordset.Fields("NumIncentivo") = sngNumLinea
'         Me.adoIncentivo.Recordset.Fields("Valor") = ActAnt(Me.adoConsulta.Recordset.Fields("FechaContrato"), 48 * sngTarifaHoraria)
'         Me.adoIncentivo.Recordset.Fields("NumVez") = 1
'         Me.adoIncentivo.Recordset.Fields("Pagado") = 0
'         Me.adoIncentivo.Recordset.Fields("NumNomina") = Trim(Me.lblNoNomina.Caption)
'         Me.adoIncentivo.Recordset.Update
'      End If
'
'
'   End If
   
   
   
   
Loop

'ElseIf Me.cboTipoNomina.Text = "Administracion" Then
'
' Me.adoAsistencia.CommandType = adCmdText
' Me.adoAsistencia.RecordSource = "SELECT Empleado.CodEmpleado, TipoNomina.Nomina, Empleado.SueldoPeriodo, Empleado.TarifaHoraria " & _
'                                 "FROM Empleado INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
'                                "WHERE (TipoNomina.Nomina = '" & Me.cboTipoNomina.Text & " ') ORDER BY Empleado.CodEmpleado ASC"
'    '                            "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaSalida <> Null AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' ORDER BY AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC "
' Me.adoAsistencia.Refresh
'
'Me.ospHoras.Value = 0
'Me.ospHoras.Min = 0
'Me.ospHoras.Max = Me.adoAsistencia.Recordset.RecordCount
'Me.ospHoras.Value = 0
'
'
' Do While Not Me.adoAsistencia.Recordset.EOF
'   Me.adoIncentivo.CommandType = adCmdText
'   Me.adoIncentivo.RecordSource = "SELECT Max(NumIncentivo) AS [MaxNumIncentivo] FROM Incentivo"
'   Me.adoIncentivo.Refresh
'
'   Me.ospHoras.Value = Me.ospHoras.Value + 1
'
'
'   If Not IsNull(Me.adoIncentivo.Recordset.Fields("MaxNumIncentivo")) Then
'      sngNumLinea = Me.adoIncentivo.Recordset.Fields("MAxNumIncentivo") + 1
'   Else
'      sngNumLinea = 1
'   End If
'
'   Me.adoIncentivo.CommandType = adCmdText
'   Me.adoIncentivo.RecordSource = "SELECT NumIncentivo, CodEmpleado, CodTipoIncentivo, NumVeces, Pagado FROM Incentivo " & _
'                                  "WHERE CodEmpleado ='" & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & "' AND CodTipoIncentivo ='01'"
'   Me.adoIncentivo.Refresh
'
'   If Me.adoIncentivo.Recordset.EOF Then
'      Me.adoIncentivo.Recordset.AddNew
'      Me.adoIncentivo.Recordset.Fields("NumIncentivo") = sngNumLinea
'      Me.adoIncentivo.Recordset.Fields("CodEmpleado") = Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado"))
'      Me.adoIncentivo.Recordset.Fields("CodTipoIncentivo") = "01"
'      Me.adoIncentivo.Recordset.Fields("NumVeces") = "n"
'      Me.adoIncentivo.Recordset.Fields("Pagado") = 0
'      Me.adoIncentivo.Recordset.Update
'   Else
'
'      sngNumLinea = Me.adoIncentivo.Recordset.Fields("NumIncentivo")
'
'   End If
'
'   Me.adoIncentivo.CommandType = adCmdText
'   Me.adoIncentivo.RecordSource = "SELECT Id, NumIncentivo, Valor, NumVez, Pagado, NumNomina FROM DetalleIncentivo"
'
'   Me.adoIncentivo.Refresh
'
'   Me.adoIncentivo.CommandType = adCmdText
'   Me.adoIncentivo.RecordSource = "SELECT Max(Id) AS [MaxNumId] FROM DetalleIncentivo"
'
'   Me.adoIncentivo.Refresh
'
'   If Not IsNull(Me.adoIncentivo.Recordset.Fields("MaxNumId")) Then
'      sngID = Me.adoIncentivo.Recordset.Fields("MaxNumId") + 1
'   Else
'      sngID = 1
'   End If
'
'   Me.adoIncentivo.CommandType = adCmdText
'   Me.adoIncentivo.RecordSource = "SELECT Id, NumIncentivo, Valor, NumVez, Pagado, NumNomina FROM DetalleIncentivo " & _
'                                  "WHERE NumIncentivo =" & sngNumLinea & " AND NumNomina =" & Trim(Me.lblNoNomina.Caption)
'   Me.adoIncentivo.Refresh
'
'   Me.adoConsulta.CommandType = adCmdText
'   Me.adoConsulta.RecordSource = "SELECT Id, CodEmpleado, FechaContrato FROM Historico " & _
'                                 "WHERE CodEmpleado ='" & Trim(Me.adoAsistencia.Recordset.Fields("CodEmpleado")) & "'"
'   Me.adoConsulta.Refresh
'
'
'
'   If Not Me.adoIncentivo.Recordset.EOF And Not Me.adoConsulta.Recordset.EOF Then
'
'     If ActAnt(Me.adoConsulta.Recordset.Fields("FechaContrato"), Me.adoAsistencia.Recordset.Fields("SueldoPeriodo")) <> 0 Then
'        Me.adoIncentivo.Recordset.Fields("Valor") = Format(ActAnt(Me.adoConsulta.Recordset.Fields("FechaContrato"), Me.adoAsistencia.Recordset.Fields("SueldoPeriodo")), "##.##")
'     Else
'        Me.adoIncentivo.Recordset.Fields("Valor") = 0
'     End If
'
'      Me.adoIncentivo.Recordset.Update
'      Me.adoIncentivo.Refresh
'   Else
'      If Me.adoConsulta.Recordset.EOF Then
'         MsgBox "El empleado " & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & ", no tiene registrado su fecha de contrato, no se puede obtener su incentivo por antiguedad"
'      ElseIf Me.adoIncentivo.Recordset.EOF And Not IsNull(Me.adoConsulta.Recordset.Fields("FechaContrato")) Then
'         Me.adoIncentivo.Recordset.AddNew
'         Me.adoIncentivo.Recordset.Fields("Id") = sngID
'         Me.adoIncentivo.Recordset.Fields("NumIncentivo") = sngNumLinea
'
'         If Me.adoIncentivo.Recordset.Fields("Valor") = ActAnt(Me.adoConsulta.Recordset.Fields("FechaContrato"), Me.adoAsistencia.Recordset.Fields("SueldoPeriodo")) <> 0 Then
'            Me.adoIncentivo.Recordset.Fields("Valor") = Format(Me.adoIncentivo.Recordset.Fields("Valor"), "##.##")
'         Else
'            Me.adoIncentivo.Recordset.Fields("Valor") = 0
'         End If
'
'         Me.adoIncentivo.Recordset.Fields("NumVez") = 1
'         Me.adoIncentivo.Recordset.Fields("Pagado") = 0
'         Me.adoIncentivo.Recordset.Fields("NumNomina") = Trim(Me.lblNoNomina.Caption)
'         Me.adoIncentivo.Recordset.Update
'
'      ElseIf IsNull(Me.adoConsulta.Recordset.Fields("FechaContrato")) Then
'         MsgBox "El empleado " & Me.adoAsistencia.Recordset.Fields("CodEmpleado") & ", no tiene registrado su fecha de contrato, no se puede obtener su incentivo por antiguedad"
'
'
'      End If
'
'
'   End If
'
'
'  Me.adoAsistencia.Recordset.MoveNext
'
' Loop


End If





Me.ospHoras.Visible = False



End Sub

Private Sub cmdReporte_Click()

Dim sSQL As String
Dim rptAsistenciaGen As New arepAsistencia
Dim rptAsistenciaSexo As New arepAsistenciaSexo
Dim rptAsistenciaDepto As New arepAsistenciaDepto
Dim rptAsistenciaCargo As New arepAsistenciaCargo
Dim rptLaboradasExtras As New arepHLaboradas
Dim lFecha1 As Long
Dim lFecha2 As Long
Dim dFecha1 As Date
Dim dFecha2 As Date
Dim sFecha1 As String
Dim sFecha2 As String


dFecha1 = Me.dtpDesde.Value
dFecha2 = Me.dtpHasta.Value

sFecha1 = Mid$(Me.dtpDesde.Value, 7, 4) & "-" & Mid$(Me.dtpDesde.Value, 4, 2) & "-" & Mid$(Me.dtpDesde.Value, 1, 2)
sFecha2 = Mid$(Me.dtpHasta.Value, 7, 4) & "-" & Mid$(Me.dtpHasta.Value, 4, 2) & "-" & Mid$(Me.dtpHasta.Value, 1, 2)

lFecha1 = dFecha1
lFecha2 = dFecha2





If Me.cboTipoNomina.Text <> "" And Me.lblNoNomina.Caption <> "" Then

If Me.optCodigo.Value Then

'sSQl = "SELECT AsistenciaEmpleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, " & _
'       "Empleado.Direccion, Empleado.Nacionalidad, Empleado.Sexo, Empleado.NumCedula, Departamento.Departamento, " & _
'       "Turno.CodTurno, Cargo.Cargo, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, " & _
'       "AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.Dia, " & _
'       "AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso, TipoNomina.Nomina, TipoNomina.Periodo, Turno.TComida " & _
'        "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN " & _
'        "Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
'        "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN TipoNomina ON dbo.AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
'        "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' ORDER BY Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC"


'sSQL = "SELECT dbo.AsistenciaEmpleado.CodEmpleado, dbo.Empleado.Nombre1, dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2, " & _
'       "dbo.Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
'       "dbo.Turno.CodTurno, dbo.Cargo.Cargo, dbo.AsistenciaEmpleado.FechaEntrada, dbo.AsistenciaEmpleado.HoraEntrada, " & _
'       "dbo.AsistenciaEmpleado.FechaSalida, dbo.AsistenciaEmpleado.HoraSalida, dbo.AsistenciaEmpleado.bActivo, dbo.AsistenciaEmpleado.HLaboradas, " & _
'       "dbo.AsistenciaEmpleado.Dia, dbo.AsistenciaEmpleado.HExtras, dbo.AsistenciaEmpleado.bPermiso, dbo.TipoNomina.Nomina, " & _
'       "dbo.TipoNomina.Periodo , dbo.Turno.TComida, dbo.HorasExtras.CantHoras, dbo.HorasExtras.NumNomina " & _
'       "FROM dbo.AsistenciaEmpleado INNER JOIN " & _
'       "dbo.Empleado ON dbo.AsistenciaEmpleado.CodEmpleado = dbo.Empleado.CodEmpleado INNER JOIN " & _
'       "dbo.Turno ON dbo.AsistenciaEmpleado.CodTurno = dbo.Turno.CodTurno INNER JOIN " & _
'       "dbo.Departamento ON dbo.Empleado.CodDepartamento = dbo.Departamento.CodDepartamento INNER JOIN " & _
'       "dbo.Cargo ON dbo.Empleado.CodCargo = dbo.Cargo.CodCargo INNER JOIN " & _
'       "dbo.TipoNomina ON dbo.AsistenciaEmpleado.CodTipoNomina = dbo.TipoNomina.CodTipoNomina INNER JOIN " & _
'       "dbo.HorasExtras ON dbo.Empleado.CodEmpleado = dbo.HorasExtras.CodEmpleado " & _
'       "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' AND dbo.HorasExtras.NumNomina =" & Me.lblNoNomina.Caption & " ORDER BY Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC"
        
sSQL = "SELECT dbo.AsistenciaEmpleado.CodEmpleado, dbo.AsistenciaEmpleado.CodEmpleado1, dbo.Empleado.Nombre1, dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2, " & _
                      "dbo.Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
                      "dbo.Turno.CodTurno, dbo.Cargo.Cargo, dbo.AsistenciaEmpleado.FechaEntrada, dbo.AsistenciaEmpleado.HoraEntrada, " & _
                      "dbo.AsistenciaEmpleado.FechaSalida, dbo.AsistenciaEmpleado.HoraSalida, dbo.AsistenciaEmpleado.bActivo, dbo.AsistenciaEmpleado.HLaboradas, " & _
                      "dbo.AsistenciaEmpleado.Dia, dbo.AsistenciaEmpleado.HExtras, dbo.AsistenciaEmpleado.bPermiso, dbo.TipoNomina.Nomina, " & _
                      "dbo.TipoNomina.Periodo , dbo.Turno.TComida, dbo.HorasExtras.CantHoras, dbo.HorasExtras.NumNomina " & _
       "FROM dbo.AsistenciaEmpleado INNER JOIN dbo.Empleado ON dbo.AsistenciaEmpleado.CodEmpleado = dbo.Empleado.CodEmpleado INNER JOIN " & _
                      "dbo.Turno ON dbo.AsistenciaEmpleado.CodTurno = dbo.Turno.CodTurno INNER JOIN " & _
                      "dbo.Departamento ON dbo.Empleado.CodDepartamento = dbo.Departamento.CodDepartamento INNER JOIN " & _
                      "dbo.Cargo ON dbo.Empleado.CodCargo = dbo.Cargo.CodCargo INNER JOIN " & _
                      "dbo.TipoNomina ON dbo.AsistenciaEmpleado.CodTipoNomina = dbo.TipoNomina.CodTipoNomina INNER JOIN " & _
                      "dbo.HorasExtras ON dbo.Empleado.CodEmpleado = dbo.HorasExtras.CodEmpleado " & _
                      "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' AND dbo.HorasExtras.NumNomina =" & Me.lblNoNomina.Caption & _
                      " GROUP BY dbo.AsistenciaEmpleado.CodEmpleado1, dbo.AsistenciaEmpleado.CodEmpleado, dbo.Empleado.Nombre1, dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2, " & _
                      "dbo.Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
                      "dbo.Turno.CodTurno, dbo.Cargo.Cargo, dbo.AsistenciaEmpleado.FechaEntrada, dbo.AsistenciaEmpleado.HoraEntrada, " & _
                      "dbo.AsistenciaEmpleado.FechaSalida, dbo.AsistenciaEmpleado.HoraSalida, dbo.AsistenciaEmpleado.bActivo, dbo.AsistenciaEmpleado.HLaboradas, " & _
                      "dbo.AsistenciaEmpleado.Dia, dbo.AsistenciaEmpleado.HExtras, dbo.AsistenciaEmpleado.bPermiso, dbo.TipoNomina.Nomina, " & _
                      "dbo.TipoNomina.Periodo , dbo.Turno.TComida, dbo.HorasExtras.CantHoras, dbo.HorasExtras.NumNomina"

rptAsistenciaGen.DataControl1.ConnectionString = ConexionRep
rptAsistenciaGen.DataControl1.Source = sSQL
rptAsistenciaGen.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value
'rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
rptAsistenciaGen.Show 1

ElseIf Me.optLaboradas.Value Then

   
   
   sSQL = "SELECT TOP 100 PERCENT dbo.Empleado.CodEmpleado, dbo.Empleado.Nombre1, dbo.Empleado.Nombre2, dbo.Empleado.Apellido1, dbo.Empleado.Apellido2, " & _
          "dbo.Empleado.TarifaHoraria, dbo.TipoNomina.Nomina, dbo.DetalleHorasProduccion.NumNomina, dbo.DetalleHorasProduccion.Lunes, " & _
          "dbo.DetalleHorasProduccion.Martes, dbo.DetalleHorasProduccion.Miercoles, dbo.DetalleHorasProduccion.Jueves, " & _
          "dbo.DetalleHorasProduccion.Viernes, dbo.DetalleHorasProduccion.Sabado, dbo.DetalleHorasProduccion.Domingo, " & _
          "dbo.DetalleHorasProduccion.TotalHoras, dbo.DetalleHorasProduccion.SalarioHora, dbo.DetalleHorasProduccion.TotalSalarioHora, " & _
          "dbo.HorasExtras.CantHoras , dbo.Departamento.Departamento, dbo.HorasExtras.NumNomina " & _
          "FROM dbo.Departamento INNER JOIN " & _
          "dbo.Empleado ON dbo.Departamento.CodDepartamento = dbo.Empleado.CodDepartamento INNER JOIN dbo.DetalleHorasProduccion ON dbo.Empleado.CodEmpleado = dbo.DetalleHorasProduccion.CodEmpleado INNER JOIN " & _
          "dbo.HorasExtras ON dbo.Empleado.CodEmpleado = dbo.HorasExtras.CodEmpleado INNER JOIN " & _
          "dbo.TipoNomina ON dbo.Empleado.CodTipoNomina = dbo.TipoNomina.CodTipoNomina " & _
          "WHERE TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' AND DetalleHorasProduccion.NumNomina =" & Me.lblNoNomina.Caption & " AND HorasExtras.NumNomina =" & Me.lblNoNomina.Caption & " ORDER BY Empleado.CodEmpleado ASC"
          
      rptLaboradasExtras.DataControl1.ConnectionString = ConexionRep
      rptLaboradasExtras.DataControl1.Source = sSQL
      rptLaboradasExtras.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value
      'rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
      rptLaboradasExtras.Show 1

      
      

ElseIf Me.optSexo.Value Then

sSQL = "SELECT AsistenciaEmpleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
       "Turno.CodTurno, Cargo.Cargo, Historico.FechaNacimiento, Historico.FechaContrato, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.HExtras, AsistenciaEmpleado.bPermiso, TipoNomina.Nomina , TipoNomina.Periodo, Turno.TComida " & _
       "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
       "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
       "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' ORDER BY Empleado.Sexo, Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC "

rptAsistenciaSexo.DataControl1.ConnectionString = ConexionRep
rptAsistenciaSexo.DataControl1.Source = sSQL
rptAsistenciaSexo.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value & ", Por Sexo"
'rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
rptAsistenciaSexo.Show 1

ElseIf Me.optDepto.Value And Me.cboDepto.Text <> "" Then

   If Me.cboDepto.Text = "Todos" Then
   
    sSQL = "SELECT AsistenciaEmpleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
           "Turno.CodTurno, Cargo.Cargo, Historico.FechaNacimiento, Historico.FechaContrato, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.HExtras, AsistenciaEmpleado.bPermiso, TipoNomina.Nomina , TipoNomina.Periodo, Turno.TComida " & _
           "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
           "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
           "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' ORDER BY Departamento.Departamento, Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC "
   Else
   
    sSQL = "SELECT AsistenciaEmpleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
           "Turno.CodTurno, Cargo.Cargo, Historico.FechaNacimiento, Historico.FechaContrato, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.HExtras, AsistenciaEmpleado.bPermiso, TipoNomina.Nomina , TipoNomina.Periodo, Turno.TComida " & _
           "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
           "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
           "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' AND Departamento.Departamento ='" & Me.cboDepto.Text & "' ORDER BY Departamento.Departamento, Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC "
   End If
    
    
    
    
   rptAsistenciaDepto.DataControl1.ConnectionString = ConexionRep
   rptAsistenciaDepto.DataControl1.Source = sSQL
   rptAsistenciaDepto.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value & ", Por Depto"
   'rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
   rptAsistenciaDepto.Show 1

ElseIf Me.optCargo.Value And Me.cboCargo.Text <> "" Then

  If Me.cboCargo.Text = "Todos" Then
    sSQL = "SELECT AsistenciaEmpleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
           "Turno.CodTurno, Cargo.Cargo, Historico.FechaNacimiento, Historico.FechaContrato, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.HExtras, AsistenciaEmpleado.bPermiso, TipoNomina.Nomina , TipoNomina.Periodo, Turno.TComida " & _
           "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
           "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
           "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' ORDER BY Cargo.Cargo, Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC "
  Else
   
    sSQL = "SELECT AsistenciaEmpleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.Direccion, dbo.Empleado.Nacionalidad, dbo.Empleado.Sexo, dbo.Empleado.NumCedula, dbo.Departamento.Departamento, " & _
           "Turno.CodTurno, Cargo.Cargo, Historico.FechaNacimiento, Historico.FechaContrato, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.HExtras, AsistenciaEmpleado.bPermiso, TipoNomina.Nomina , TipoNomina.Periodo, Turno.TComida " & _
           "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
           "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN TipoNomina ON AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
           "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND AsistenciaEmpleado.FechaEntrada <= CONVERT(DATETIME, '" & sFecha2 & " 00:00:00" & "', 102) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' AND Cargo.Cargo ='" & Me.cboCargo.Text & "' ORDER BY Cargo.Cargo, Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC "
   End If
    
    
    
    
   rptAsistenciaCargo.DataControl1.ConnectionString = ConexionRep
   rptAsistenciaCargo.DataControl1.Source = sSQL
   rptAsistenciaCargo.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value & ", Por Cargo"
   'rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
   rptAsistenciaCargo.Show 1


ElseIf Me.optSalidasNoRegistradas.Value Then



sSQL = "SELECT AsistenciaEmpleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, " & _
       "Empleado.Direccion, Empleado.Nacionalidad, Empleado.Sexo, Empleado.NumCedula, Departamento.Departamento, " & _
       "Turno.CodTurno, Cargo.Cargo, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.HoraEntrada, " & _
       "AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraSalida, AsistenciaEmpleado.bActivo, AsistenciaEmpleado.HLaboradas, AsistenciaEmpleado.Dia, " & _
       "AsistenciaEmpleado.HExtras , AsistenciaEmpleado.bPermiso, TipoNomina.Nomina, TipoNomina.Periodo, Turno.TComida " & _
        "FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN " & _
        "Turno ON AsistenciaEmpleado.CodTurno = Turno.CodTurno INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN " & _
        "Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN TipoNomina ON dbo.AsistenciaEmpleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
        "WHERE AsistenciaEmpleado.FechaEntrada >= CONVERT(DATETIME, '" & sFecha1 & " 00:00:00" & "', 102) AND (AsistenciaEmpleado.FechaSalida IS NULL) AND TipoNomina.Nomina ='" & Me.cboTipoNomina.Text & "' ORDER BY Empleado.CodEmpleado, AsistenciaEmpleado.FechaEntrada ASC"

rptAsistenciaGen.lblTitulo.Caption = Me.optSalidasNoRegistradas.Caption
rptAsistenciaGen.DataControl1.ConnectionString = ConexionRep
rptAsistenciaGen.DataControl1.Source = sSQL
rptAsistenciaGen.lblMensaje.Caption = "Nomina: " & Me.cboTipoNomina.Text & ", del " & Me.dtpDesde.Value & " al " & Me.dtpHasta.Value
'rptAsistenciaGen.lblMensaje.Caption = sMensajeReporte
rptAsistenciaGen.Show 1




End If



End If




End Sub

Private Sub cmdSalir_Click()

Unload Me

End Sub

Private Sub Form_Activate()

Me.cboTipoNomina.SetFocus

End Sub

Private Sub Form_Load()

 Dim RutaServer As String
 Dim Server As String
 Dim Conexion As String
 Dim Clave As String
 Dim User As String
 
 

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
  ConexionRep = ConexionSTR1

ConexionRep = Conexion

Me.adoAsistencia.ConnectionString = Conexion
Me.adoPermisos.ConnectionString = Conexion
Me.adoTipoNomina.ConnectionString = Conexion
Me.adoTurno.ConnectionString = Conexion

Me.adoConsulta.ConnectionString = Conexion
Me.adoIncentivo.ConnectionString = Conexion
Me.adoHorasExtras.ConnectionString = Conexion



Me.adoAsistencia.CommandType = adCmdText
Me.adoAsistencia.RecordSource = "SELECT * FROM AsistenciaEmpleado"
Me.adoAsistencia.Refresh

Me.adoIncentivo.CommandType = adCmdText
Me.adoIncentivo.RecordSource = "SELECT NumIncentivo, CodEmpleado, CodTipoIncentivo, NumVeces, Pagado FROM Incentivo"
Me.adoIncentivo.Refresh

Me.adoPermisos.CommandType = adCmdText
Me.adoPermisos.RecordSource = "SELECT * FROM Permisos"
Me.adoPermisos.Refresh

Me.adoTurno.CommandType = adCmdText
Me.adoTurno.RecordSource = "SELECT * FROM Turno"
Me.adoTurno.Refresh

Me.adoTipoNomina.CommandType = adCmdTable
Me.adoTipoNomina.RecordSource = "Departamento"
Me.adoTipoNomina.Refresh

Me.cboDepto.AddItem "Todos"

Do While Not Me.adoTipoNomina.Recordset.EOF

   Me.cboDepto.AddItem Me.adoTipoNomina.Recordset.Fields("Departamento")
   Me.adoTipoNomina.Recordset.MoveNext

Loop

Me.adoTipoNomina.CommandType = adCmdTable
Me.adoTipoNomina.RecordSource = "Cargo"
Me.adoTipoNomina.Refresh

Me.cboCargo.AddItem "Todos"

Do While Not Me.adoTipoNomina.Recordset.EOF

   Me.cboCargo.AddItem Me.adoTipoNomina.Recordset.Fields("Cargo")
   Me.adoTipoNomina.Recordset.MoveNext

Loop

Me.adoTipoNomina.CommandType = adCmdText
Me.adoTipoNomina.RecordSource = "SELECT TipoNomina.CodTipoNomina, TipoNomina.Nomina, Nomina.NumNomina, Nomina.FechaNominaINI, Nomina.FechaNomina, " & _
                                "Nomina.Activa FROM Nomina INNER JOIN TipoNomina ON dbo.Nomina.CodTipoNomina = dbo.TipoNomina.CodTipoNomina " & _
                                "WHERE (Nomina.Activa = 1)"
Me.adoTipoNomina.Refresh

Do While Not Me.adoTipoNomina.Recordset.EOF

   Me.cboTipoNomina.AddItem Me.adoTipoNomina.Recordset.Fields("Nomina")
   Me.adoTipoNomina.Recordset.MoveNext

Loop







End Sub

