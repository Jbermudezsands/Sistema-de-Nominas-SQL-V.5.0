VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmBajas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Despidos y Renuncias"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   FillStyle       =   2  'Horizontal Line
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adoAuxiliar 
      Height          =   375
      Left            =   4800
      Top             =   7560
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.CommandButton CmdRenovar 
      Caption         =   "Renovacion"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7680
      TabIndex        =   88
      Top             =   6120
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc AdoEmpresa 
      Height          =   375
      Left            =   600
      Top             =   8640
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
      Caption         =   "AdoEmpresa"
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
   Begin MSAdodcLib.Adodc AdoElimina 
      Height          =   375
      Left            =   600
      Top             =   7920
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
      Caption         =   "AdoElimina"
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
   Begin MSAdodcLib.Adodc AdoSalarios 
      Height          =   375
      Left            =   120
      Top             =   7920
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
      Caption         =   "AdoSalarios"
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
   Begin MSAdodcLib.Adodc AdoInicioAño 
      Height          =   375
      Left            =   6600
      Top             =   9000
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "AdoInicioAño"
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
   Begin MSAdodcLib.Adodc AdoAntiguedad 
      Height          =   375
      Left            =   3120
      Top             =   9000
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "AdoAntiguedad"
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
   Begin VB.CommandButton CmdCalculos 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   120
      TabIndex        =   44
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   9240
      TabIndex        =   40
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton CmdEfectuar 
      Caption         =   "Procesar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   39
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton CmdCalcular 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   4440
      TabIndex        =   38
      Top             =   7200
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc DtaNominas 
      Height          =   375
      Left            =   3120
      Top             =   8280
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "DtaNominas"
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
   Begin MSAdodcLib.Adodc DtaPrestamo 
      Height          =   375
      Left            =   3120
      Top             =   7920
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "DtaPrestamo"
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
   Begin MSAdodcLib.Adodc DtaDeducciones 
      Height          =   375
      Left            =   2280
      Top             =   8400
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "DtaDeducciones"
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
   Begin MSAdodcLib.Adodc DtaHrsExtras 
      Height          =   375
      Left            =   2760
      Top             =   8400
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "DtaHrsExtras"
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
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   360
      Top             =   8400
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "DtaConsulta"
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
   Begin MSAdodcLib.Adodc DtaTipoNomina 
      Height          =   375
      Left            =   360
      Top             =   8520
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "DtaTipoNomina"
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
   Begin MSAdodcLib.Adodc DtaInss 
      Height          =   375
      Left            =   3480
      Top             =   7920
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "DtaInss"
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
   Begin MSAdodcLib.Adodc DtaHistorico 
      Height          =   375
      Left            =   360
      Top             =   9000
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "DtaHistorico"
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
   Begin MSAdodcLib.Adodc DtaDeduccion 
      Height          =   375
      Left            =   5760
      Top             =   8400
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "DtaDeduccion"
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
      Left            =   360
      Top             =   8280
      Width           =   2775
      _ExtentX        =   4895
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
   Begin MSAdodcLib.Adodc DtaControles 
      Height          =   375
      Left            =   360
      Top             =   7920
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "DtaControles"
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
   Begin MSAdodcLib.Adodc DtaAdelanto 
      Height          =   375
      Left            =   600
      Top             =   8040
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "DtaAdelanto"
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
   Begin MSAdodcLib.Adodc DtaBajas 
      Height          =   375
      Left            =   360
      Top             =   8280
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "DtaBajas"
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
   Begin MSAdodcLib.Adodc DtaIR 
      Height          =   375
      Left            =   600
      Top             =   8040
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "DtaIR"
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
   Begin VB.PictureBox Picture1 
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5355
      ScaleWidth      =   10635
      TabIndex        =   0
      Top             =   600
      Width           =   10695
      Begin TabDlg.SSTab SSTab1 
         Height          =   5175
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   9128
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Datos Generales"
         TabPicture(0)   =   "FrmBajas.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label6"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label10"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label12"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label54"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label55"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label56"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label5"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label7"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label8"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label9"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label11"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label13"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Label14"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label4"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Label16"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Line1"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Line2"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Line3"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Line4"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "Line6"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "Label15"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "ChkSueldoActual"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "TxtFechaHistorial"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "TxtNombre1"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "TxtDireccion"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "TxtApellido2"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "TxtApellido1"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "TxtNombre2"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "TxtDepartamento"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "TxtCargo"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "TxtSexo"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "TxtCodEmpleado"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "TxtFechaContrato"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "TxtAnnos"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "TxtMeses"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "TxtDias"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "CmdBuscarEmpleado"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "TxtMotivo"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "Frame1"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "TxtUltFechaNomina"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "TxtDiasTrabajados"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "TxtCodTipoNomina"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "ChkAntiguedad"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "TxtCodEmpleado1"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).ControlCount=   47
         TabCaption(1)   =   "Historial Salarial"
         TabPicture(1)   =   "FrmBajas.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).Control(1)=   "CmdDetalle"
         Tab(1).Control(2)=   "TDBGridSalarios"
         Tab(1).Control(3)=   "TDBGridBonos"
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Ingresos / Egresos"
         TabPicture(2)   =   "FrmBajas.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame4"
         Tab(2).Control(1)=   "Frame2"
         Tab(2).ControlCount=   2
         Begin TrueOleDBGrid70.TDBGrid TDBGridBonos 
            Bindings        =   "FrmBajas.frx":0054
            Height          =   2175
            Left            =   -74880
            TabIndex        =   86
            Top             =   600
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   3836
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Basico"
            Columns(0).DataField=   "SalarioBasico"
            Columns(0).NumberFormat=   "###,###0.00"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Produccion"
            Columns(1).DataField=   "Destajo"
            Columns(1).NumberFormat=   "###,###0.00"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Bono"
            Columns(2).DataField=   "BonoProduccion"
            Columns(2).NumberFormat=   "###,###0.00"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Horas Extra"
            Columns(3).DataField=   "HorasExtras"
            Columns(3).NumberFormat=   "###,###0.00"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Septimo"
            Columns(4).DataField=   "Septimo"
            Columns(4).NumberFormat=   "###,###0.00"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Incentivo"
            Columns(5).DataField=   "Incentivos"
            Columns(5).NumberFormat=   "###,###0.00"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Otros Ingresos"
            Columns(6).DataField=   "Otros"
            Columns(6).NumberFormat=   "###,###0.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Total Ingresos"
            Columns(7).DataField=   "TotalIngresos"
            Columns(7).NumberFormat=   "###,###0.00"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "MES"
            Columns(8).DataField=   "MES"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "AÑO"
            Columns(9).DataField=   "AÑO"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   10
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=10"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1693"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=2"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=1773"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1693"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=1"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=1773"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1693"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=1"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=1773"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1693"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=1"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=1667"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1588"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=2"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=1667"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1588"
            Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=1"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=1931"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1852"
            Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=2"
            Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(36)=   "Column(7).Width=2117"
            Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2037"
            Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=2"
            Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(41)=   "Column(8).Width=1138"
            Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1058"
            Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=1"
            Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(46)=   "Column(9).Width=1138"
            Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=1058"
            Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=1"
            Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   2
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "SALARIOS DE LOS ULTIMOS 6 MESES"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   14215660
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=50,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=51,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=52,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=53,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=55,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=56,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=57,.parent=17"
            _StyleDefs(76)  =   "Named:id=33:Normal"
            _StyleDefs(77)  =   ":id=33,.parent=0"
            _StyleDefs(78)  =   "Named:id=34:Heading"
            _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(80)  =   ":id=34,.wraptext=-1"
            _StyleDefs(81)  =   "Named:id=35:Footing"
            _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(83)  =   "Named:id=36:Selected"
            _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=37:Caption"
            _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(87)  =   "Named:id=38:HighlightRow"
            _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=39:EvenRow"
            _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(91)  =   "Named:id=40:OddRow"
            _StyleDefs(92)  =   ":id=40,.parent=33"
            _StyleDefs(93)  =   "Named:id=41:RecordSelector"
            _StyleDefs(94)  =   ":id=41,.parent=34"
            _StyleDefs(95)  =   "Named:id=42:FilterBar"
            _StyleDefs(96)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGridSalarios 
            Bindings        =   "FrmBajas.frx":006E
            Height          =   2175
            Left            =   -74880
            TabIndex        =   65
            Top             =   600
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   3836
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Salario Basico"
            Columns(0).DataField=   "SalarioBasico"
            Columns(0).NumberFormat=   "###,###0.00"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Reembolso"
            Columns(1).DataField=   "Reembolso"
            Columns(1).NumberFormat=   "###,###0.00"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Produccion"
            Columns(2).DataField=   "Destajo"
            Columns(2).NumberFormat=   "###,###0.00"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Septimo Dias"
            Columns(3).DataField=   "Septimo"
            Columns(3).NumberFormat=   "###,###0.00"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Incentivos"
            Columns(4).DataField=   "Incentivos"
            Columns(4).NumberFormat=   "##,##0.00"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Otros Ingresos"
            Columns(5).DataField=   "Otros"
            Columns(5).NumberFormat=   "###,###0.00"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Total Ingresos"
            Columns(6).DataField=   "TotalIngresos"
            Columns(6).NumberFormat=   "###,###0.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "MES"
            Columns(7).DataField=   "MES"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "AÑO"
            Columns(8).DataField=   "AÑO"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2037"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=2"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=1931"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1852"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=2"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=1773"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1693"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=2"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=1931"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1852"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=2"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=1773"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1693"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=2"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=1931"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1852"
            Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=2"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=2461"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2381"
            Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=2"
            Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(36)=   "Column(7).Width=1402"
            Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1323"
            Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=1"
            Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(41)=   "Column(8).Width=1402"
            Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1323"
            Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=1"
            Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   2
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "SALARIOS DE LOS ULTIMOS 6 MESES"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   14215660
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=70,.parent=13,.alignment=1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
            _StyleDefs(72)  =   "Named:id=33:Normal"
            _StyleDefs(73)  =   ":id=33,.parent=0"
            _StyleDefs(74)  =   "Named:id=34:Heading"
            _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(76)  =   ":id=34,.wraptext=-1"
            _StyleDefs(77)  =   "Named:id=35:Footing"
            _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(79)  =   "Named:id=36:Selected"
            _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=37:Caption"
            _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(83)  =   "Named:id=38:HighlightRow"
            _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=39:EvenRow"
            _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(87)  =   "Named:id=40:OddRow"
            _StyleDefs(88)  =   ":id=40,.parent=33"
            _StyleDefs(89)  =   "Named:id=41:RecordSelector"
            _StyleDefs(90)  =   ":id=41,.parent=34"
            _StyleDefs(91)  =   "Named:id=42:FilterBar"
            _StyleDefs(92)  =   ":id=42,.parent=33"
         End
         Begin VB.Frame Frame2 
            Caption         =   "Prestaciones"
            Height          =   3375
            Left            =   -74880
            TabIndex        =   52
            Top             =   480
            Width           =   4455
            Begin VB.CheckBox ChkIncentivos 
               Caption         =   "Incentivos"
               Height          =   255
               Left            =   1800
               TabIndex        =   92
               Top             =   1320
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.CheckBox ChkOtroPlanilla 
               Caption         =   "Otros Ingresos Planilla"
               Height          =   255
               Left            =   1800
               TabIndex        =   87
               Top             =   1080
               Width           =   1935
            End
            Begin VB.CheckBox Chk13mes 
               Caption         =   "13vo Mes"
               Height          =   255
               Left            =   240
               TabIndex        =   64
               Top             =   360
               Value           =   1  'Checked
               Width           =   1455
            End
            Begin VB.CheckBox ChkVaca 
               Caption         =   "Vacaciones"
               Height          =   255
               Left            =   1800
               TabIndex        =   63
               Top             =   360
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.CheckBox ChkAntigue 
               Caption         =   "Antiguedad"
               Height          =   195
               Left            =   240
               TabIndex        =   62
               Top             =   600
               Value           =   1  'Checked
               Width           =   1455
            End
            Begin VB.CheckBox ChkExtra 
               Caption         =   "Horas Extra"
               Height          =   255
               Left            =   1800
               TabIndex        =   61
               Top             =   600
               Width           =   1575
            End
            Begin VB.CheckBox ChkCargo 
               Caption         =   "Viaticos"
               Height          =   195
               Left            =   240
               TabIndex        =   60
               Top             =   840
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.CheckBox ChkOtro 
               Caption         =   "Otros Ingresos"
               Height          =   255
               Left            =   1800
               TabIndex        =   59
               Top             =   840
               Width           =   1335
            End
            Begin VB.TextBox TxtDescuentoDias 
               Height          =   285
               Left            =   2400
               TabIndex        =   58
               Top             =   1680
               Width           =   855
            End
            Begin VB.TextBox TxtOtrPrestacion 
               Height          =   285
               Left            =   960
               TabIndex        =   57
               Top             =   2040
               Visible         =   0   'False
               Width           =   2535
            End
            Begin VB.TextBox TxtMontoOtrPrestacion 
               Height          =   285
               Left            =   1920
               TabIndex        =   56
               Top             =   2400
               Visible         =   0   'False
               Width           =   1575
            End
            Begin ACTIVESKINLibCtl.SkinLabel LblMonto 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmBajas.frx":0088
               TabIndex        =   53
               Top             =   2400
               Visible         =   0   'False
               Width           =   1455
            End
            Begin ACTIVESKINLibCtl.SkinLabel LblPrestacion 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmBajas.frx":0100
               TabIndex        =   54
               Top             =   2040
               Visible         =   0   'False
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel TxtDiasDescuento 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmBajas.frx":0172
               TabIndex        =   55
               Top             =   1680
               Width           =   2175
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Deducciones"
            Height          =   3375
            Left            =   -70440
            TabIndex        =   49
            Top             =   480
            Width           =   4335
            Begin VB.CheckBox ChkIR 
               Caption         =   "IR"
               Height          =   255
               Left            =   360
               TabIndex        =   93
               Top             =   1200
               Width           =   2055
            End
            Begin VB.CheckBox ChkPrestamo 
               Caption         =   "Prestamos"
               Height          =   255
               Left            =   360
               TabIndex        =   51
               Top             =   480
               Width           =   2295
            End
            Begin VB.CheckBox ChkDeducciones 
               Caption         =   "Deducciones"
               Height          =   255
               Left            =   360
               TabIndex        =   50
               Top             =   840
               Width           =   2055
            End
         End
         Begin VB.TextBox TxtCodEmpleado1 
            Height          =   285
            Left            =   1560
            TabIndex        =   47
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox ChkAntiguedad 
            Caption         =   "Antiguedad Base para Calculo"
            Height          =   255
            Left            =   5040
            TabIndex        =   46
            Top             =   3480
            Width           =   3255
         End
         Begin VB.TextBox TxtCodTipoNomina 
            Height          =   285
            Left            =   3360
            TabIndex        =   45
            Top             =   3360
            Width           =   615
         End
         Begin VB.TextBox TxtDiasTrabajados 
            Height          =   285
            Left            =   7320
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   1320
            Width           =   975
         End
         Begin MSComCtl2.DTPicker TxtUltFechaNomina 
            Height          =   300
            Left            =   6120
            TabIndex        =   41
            Top             =   2640
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            Format          =   82182145
            CurrentDate     =   38802
         End
         Begin VB.Frame Frame1 
            Caption         =   "Tipo de Baja"
            Height          =   615
            Left            =   4680
            TabIndex        =   34
            Top             =   600
            Width           =   3735
            Begin VB.OptionButton OptDespido 
               Caption         =   "Despido"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton OptRenuncia 
               Caption         =   "Renuncia"
               Height          =   255
               Left            =   1080
               TabIndex        =   36
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton OptFinContrato 
               Caption         =   "Final. Contrato"
               Height          =   255
               Left            =   2160
               TabIndex        =   35
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.TextBox TxtMotivo 
            Height          =   615
            Left            =   5040
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Top             =   4320
            Width           =   3495
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
            Left            =   3000
            Picture         =   "FrmBajas.frx":0202
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox TxtDias 
            Height          =   285
            Left            =   1560
            TabIndex        =   14
            Top             =   4320
            Width           =   495
         End
         Begin VB.TextBox TxtMeses 
            Height          =   285
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox TxtAnnos 
            Height          =   285
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox TxtFechaContrato 
            Height          =   300
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   2160
            Width           =   1815
         End
         Begin VB.TextBox TxtCodEmpleado 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3360
            TabIndex        =   10
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TxtSexo 
            Height          =   285
            Left            =   1560
            TabIndex        =   9
            Top             =   3360
            Width           =   1455
         End
         Begin VB.TextBox TxtCargo 
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   3000
            Width           =   2415
         End
         Begin VB.TextBox TxtDepartamento 
            Height          =   285
            Left            =   1560
            TabIndex        =   7
            Top             =   2640
            Width           =   2415
         End
         Begin VB.TextBox TxtNombre2 
            Height          =   285
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   6
            Top             =   1200
            Width           =   2415
         End
         Begin VB.TextBox TxtApellido1 
            Height          =   285
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   5
            Top             =   1560
            Width           =   2415
         End
         Begin VB.TextBox TxtApellido2 
            Height          =   285
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   4
            Top             =   1920
            Width           =   2415
         End
         Begin VB.TextBox TxtDireccion 
            Height          =   285
            Left            =   1560
            MaxLength       =   200
            TabIndex        =   3
            Top             =   2280
            Width           =   2415
         End
         Begin VB.TextBox TxtNombre1 
            Height          =   285
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   2
            Top             =   840
            Width           =   2415
         End
         Begin SmartButtonProject.SmartButton CmdDetalle 
            Height          =   855
            Left            =   -67080
            TabIndex        =   83
            Top             =   3600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1508
            Caption         =   "Imp Detalle"
            Picture         =   "FrmBajas.frx":0350
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
         Begin VB.Frame Frame3 
            Caption         =   "Calculos Basicos del Salario"
            Height          =   2295
            Left            =   -74880
            TabIndex        =   66
            Top             =   2760
            Width           =   9015
            Begin VB.CheckBox Check1 
               Caption         =   "Salario Prom Diferencia Fecha"
               Enabled         =   0   'False
               Height          =   255
               Left            =   3240
               TabIndex        =   90
               Top             =   1680
               Width           =   3135
            End
            Begin SmartButtonProject.SmartButton CmdImprimirHistorial 
               Height          =   855
               Left            =   6600
               TabIndex        =   82
               Top             =   840
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   1508
               Caption         =   "Imp Historial"
               Picture         =   "FrmBajas.frx":066A
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
            Begin VB.TextBox TxtTarifa 
               Height          =   285
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   81
               Text            =   "0.00"
               Top             =   1800
               Width           =   1335
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmBajas.frx":0984
               TabIndex        =   80
               Top             =   1800
               Width           =   1215
            End
            Begin VB.TextBox TxtSalarioBasico 
               Height          =   285
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   79
               Text            =   "0.00"
               Top             =   1080
               Width           =   1335
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
               Height          =   375
               Left            =   240
               OleObjectBlob   =   "FrmBajas.frx":09FE
               TabIndex        =   78
               Top             =   1080
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker DTPFechaIniAgui 
               Height          =   285
               Left            =   5040
               TabIndex        =   77
               Top             =   1200
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               Format          =   82182145
               CurrentDate     =   38821
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
               Height          =   255
               Left            =   3240
               OleObjectBlob   =   "FrmBajas.frx":0A78
               TabIndex        =   76
               Top             =   1200
               Width           =   1815
            End
            Begin MSComCtl2.DTPicker DTPFechaIniVaca 
               Height          =   285
               Left            =   5040
               TabIndex        =   75
               Top             =   840
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               Format          =   82182145
               CurrentDate     =   38821
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Left            =   3240
               OleObjectBlob   =   "FrmBajas.frx":0B02
               TabIndex        =   74
               Top             =   840
               Width           =   1935
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmBajas.frx":0B8E
               TabIndex        =   73
               Top             =   1440
               Width           =   1215
            End
            Begin VB.TextBox TxtAntiguedad 
               Height          =   285
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   72
               Text            =   "0.00"
               Top             =   1440
               Width           =   1335
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmBajas.frx":0C00
               TabIndex        =   70
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox TxtSalarioAlto 
               Height          =   285
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   69
               Text            =   "0.00"
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox TxtSalarioPromedio 
               Height          =   285
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   68
               Text            =   "0.00"
               Top             =   360
               Width           =   1335
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmBajas.frx":0C7E
               TabIndex        =   67
               Top             =   360
               Width           =   1455
            End
            Begin Threed.SSCommand TxtSalarios 
               Height          =   525
               Left            =   3240
               TabIndex        =   71
               Top             =   240
               Width           =   5415
               _ExtentX        =   9551
               _ExtentY        =   926
               _Version        =   196610
               Font3D          =   2
               MarqueeStyle    =   4
               ForeColor       =   192
               MarqueeDelay    =   5
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "No se ha Definido el Empleado"
               ButtonStyle     =   4
               AutoRepeat      =   -1  'True
            End
         End
         Begin MSComCtl2.DTPicker TxtFechaHistorial 
            Height          =   300
            Left            =   6120
            TabIndex        =   84
            Top             =   3120
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            Format          =   82182145
            CurrentDate     =   38802
         End
         Begin XtremeSuiteControls.CheckBox ChkSueldoActual 
            Height          =   255
            Left            =   600
            TabIndex        =   91
            Top             =   3720
            Width           =   3375
            _Version        =   786432
            _ExtentX        =   5953
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Sueldo Actual --> Basico en Liquidacion"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label15 
            Caption         =   "Fecha Historial"
            Height          =   375
            Left            =   4680
            TabIndex        =   85
            Top             =   3120
            Width           =   1455
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000014&
            BorderStyle     =   6  'Inside Solid
            BorderWidth     =   3
            X1              =   240
            X2              =   3600
            Y1              =   4080
            Y2              =   4080
         End
         Begin VB.Line Line4 
            BorderColor     =   &H80000004&
            BorderWidth     =   3
            X1              =   3600
            X2              =   240
            Y1              =   4080
            Y2              =   4080
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   3
            X1              =   4440
            X2              =   3600
            Y1              =   3840
            Y2              =   4080
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000014&
            BorderStyle     =   6  'Inside Solid
            BorderWidth     =   3
            X1              =   4440
            X2              =   9000
            Y1              =   3840
            Y2              =   3840
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000004&
            BorderWidth     =   3
            X1              =   4440
            X2              =   4440
            Y1              =   600
            Y2              =   3840
         End
         Begin VB.Label Label16 
            Caption         =   "Dias"
            Height          =   255
            Left            =   6960
            TabIndex        =   42
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Motivo"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5040
            TabIndex        =   33
            Top             =   3960
            Width           =   3375
         End
         Begin VB.Label Label14 
            Caption         =   "Desde Ultima Nómina"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   2160
            TabIndex        =   31
            Top             =   4320
            Width           =   2055
         End
         Begin VB.Label Label13 
            Caption         =   "Totales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7680
            TabIndex        =   30
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Días Trabajados"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   4320
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Fecha Egreso"
            Height          =   375
            Left            =   4680
            TabIndex        =   28
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Meses Trabajados"
            Height          =   255
            Left            =   4680
            TabIndex        =   27
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Años Trabajados"
            Height          =   255
            Left            =   4800
            TabIndex        =   26
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha de Ingreso"
            Height          =   255
            Left            =   4680
            TabIndex        =   25
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label56 
            Caption         =   "Segundo Apellido:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label55 
            Caption         =   "Primer Apellido:"
            Height          =   375
            Left            =   360
            TabIndex        =   23
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label54 
            Caption         =   "Segundo Nombre:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label12 
            Caption         =   "Cargo:"
            Height          =   255
            Left            =   720
            TabIndex        =   21
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Depto:"
            Height          =   255
            Left            =   720
            TabIndex        =   20
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Sexo:"
            Height          =   255
            Left            =   960
            TabIndex        =   19
            Top             =   3360
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Direccion:"
            Height          =   255
            Left            =   720
            TabIndex        =   18
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Primer Nombre:"
            Height          =   255
            Left            =   360
            TabIndex        =   17
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "CodEmpleado:"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   480
            Width           =   1695
         End
      End
   End
   Begin Threed.SSCommand CmdAcercade 
      Height          =   525
      Left            =   120
      TabIndex        =   48
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   926
      _Version        =   196610
      Font3D          =   2
      MarqueeStyle    =   4
      ForeColor       =   8388608
      MarqueeDelay    =   5
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Despidos y Renuncias..."
      ButtonStyle     =   4
      AutoRepeat      =   -1  'True
   End
   Begin MSAdodcLib.Adodc AdoEmpleado 
      Height          =   375
      Left            =   960
      Top             =   8400
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "AdoEmpleado"
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
   Begin MSAdodcLib.Adodc DtaHorarioEmpleado 
      Height          =   375
      Left            =   6720
      Top             =   7920
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "DtaHorarioEmpleado"
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
   Begin MSAdodcLib.Adodc DtaTurnos 
      Height          =   375
      Left            =   6840
      Top             =   7440
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "DtaTurnos"
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
   Begin MSAdodcLib.Adodc AdoHistoricos 
      Height          =   375
      Left            =   6960
      Top             =   7080
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "AdoHistoricos"
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
   Begin MSAdodcLib.Adodc AdoTraslado 
      Height          =   375
      Left            =   480
      Top             =   7440
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "AdoTraslado"
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
   Begin XtremeSuiteControls.ProgressBar osProgress 
      Height          =   375
      Left            =   1560
      TabIndex        =   89
      Top             =   6120
      Width           =   4455
      _Version        =   786432
      _ExtentX        =   7858
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin VB.Line Line5 
      X1              =   3120
      X2              =   3120
      Y1              =   6840
      Y2              =   6720
   End
End
Attribute VB_Name = "FrmBajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public FechaBusqueda1 As Date, FechaHistorico1 As Date, SSueldoActual As Double, Referencia As String
Public Sub LlenarLiquidacion()
Dim FechaContrato As Date, FechaInicio As Date, FechaEgreso As Date
Dim FechaHoy As Date, SueldoPeriodo As Double, FechaHistorico As Date
Dim FechaUltNomina As Date, i As Integer, NumeroEmpleado As Double
Dim annos As Date
Dim SQlEmpleado As String
Dim SQlPrestamo As String
Dim SQlDeducciones As String
Dim SqlNominas As String, DiasMes As Double, DiasReales As Double, MesReal As Double
Dim FechaBusqueda As Date, Año As Integer, Mes As Integer
Dim Contador As Integer, TotalSalario As Double, Salario As Double, SalarioAlto As Double, SalarioPromedio As Double
Dim SueldoActual As Double

'SQlEmpleado = "SELECT  Empleado.SalarioFijo, Empleado.SueldoPeriodo, Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion, Empleado.Direccion AS Expr1, Empleado.Sexo, Empleado.Activo, Empleado.TarifaHoraria, Empleado.SueldoActualBasico FROM Departamento INNER JOIN Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento WHERE  (Empleado.CodEmpleado = " & TxtCodEmpleado.Text & ")"
SQlEmpleado = "SELECT Empleado.SalarioFijo, Empleado.SueldoPeriodo, Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion, Empleado.Direccion AS Expr1, Empleado.Sexo, Empleado.Activo, Empleado.TarifaHoraria, Empleado.SueldoActualBasico, Historico.SueldoActual FROM Departamento INNER JOIN Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado Where (Empleado.CodEmpleado = " & TxtCodEmpleado.Text & ")"
DtaEmpleado.RecordSource = SQlEmpleado
DtaEmpleado.Refresh

SQlPrestamo = "SELECT Prestamo.NumPrestamo, Prestamo.CodEmpleado, Prestamo.Monto, Prestamo.CantCuotas, Prestamo.Interes, Prestamo.Saldo, Prestamo.FechaInicial, Prestamo.Cancelado From Prestamo WHERE Prestamo.Cancelado=0 AND Prestamo.CodEmpleado='" & TxtCodEmpleado.Text & "'"
Dtaprestamo.RecordSource = SQlPrestamo
Dtaprestamo.Refresh

SQlDeducciones = "SELECT Deduccion.NumDeduccion, Deduccion.CodEmpleado, Deduccion.CodTipoDeduccion, DetalleDeduccion.NumDeduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.pagado  FROM Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion WHERE DetalleDeduccion.pagado=0 AND Deduccion.CodEmpleado='" & TxtCodEmpleado.Text & "'"
DtaDeducciones.RecordSource = SQlDeducciones
DtaDeducciones.Refresh

DoEvents

If Not DtaEmpleado.Recordset.EOF Then
       If DtaEmpleado.Recordset("activo") = False Then
           MsgBox "Este empleado ya fue dado de Baja"
           Exit Sub
        End If
    Me.txtCodEmpleado1.Text = DtaEmpleado.Recordset("CodEmpleado1")
    CodEmpleado = DtaEmpleado.Recordset("CodEmpleado")
    TxtNombre1 = DtaEmpleado.Recordset("Nombre1")
    TxtNombre2 = DtaEmpleado.Recordset("Nombre2")
    TxtApellido1 = DtaEmpleado.Recordset("Apellido1")
    TxtApellido2 = DtaEmpleado.Recordset("Apellido2")
    TxtDireccion = DtaEmpleado.Recordset("Direccion")
    txtCargo = DtaEmpleado.Recordset("Cargo")
    txtDepartamento = DtaEmpleado.Recordset("departamento")
    txtSexo = DtaEmpleado.Recordset("sexo")
    SalarioBasico = DtaEmpleado.Recordset("SueldoPeriodo")
    If DtaEmpleado.Recordset("SalarioFijo") = "S" Then
     SueldoFijo = True
    Else
     SueldoFijo = False
    End If
    
        If Not IsNull(DtaEmpleado.Recordset("SueldoActualBasico")) = True Then
         If DtaEmpleado.Recordset("SueldoActualBasico") = True Then
          Me.ChkSueldoActual.Value = 1
          If Not IsNull(DtaEmpleado.Recordset("SueldoActual")) Then
             SueldoActual = DtaEmpleado.Recordset("SueldoActual")
          End If
        Else
          Me.ChkSueldoActual.Value = 0
          SueldoActual = 0
         End If
        End If
Else
    TxtNombre1 = ""
    TxtNombre2 = ""
    TxtApellido1 = ""
    TxtApellido2 = ""
    TxtDireccion = ""
    txtCargo = ""
    txtDepartamento = ""
    txtSexo = ""
    TxtAnnos.Text = ""
    TxtMeses.Text = ""
    TxtFechaContrato.Text = ""
    TxtUltFechaNomina.Value = Now
    TxtMotivo.Text = ""
   Exit Sub
End If

DiasMes = 0

DtaHistorico.RecordSource = "SELECT Id, Codempleado, FechaBaja, MotivoBaja, FechaAumento, MotivoAumento, FechaInicialSusp, FechaFinalSusp, MotivoSuspencion, FechaNacimiento,FechaContrato , CargoInicial, CargoActual, CargoAnterior, SueldoInicial, SueldoAnterior, SueldoActual, CuentaDebito, CuentaCredito From Historico"
DtaHistorico.Refresh
Do While Not DtaHistorico.Recordset.EOF
    If TxtCodEmpleado.Text = DtaHistorico.Recordset("CodEmpleado") Then
        If Not IsNull(DtaHistorico.Recordset("FechaContrato")) Then
           TxtFechaContrato.Text = DtaHistorico.Recordset("FechaContrato")
           FechaContrato = DtaHistorico.Recordset("FechaContrato")
        Else
        MsgBox "No ha sido grabada la fecha de contrato no puede realizarse la baja"
        CmdEfectuar.Enabled = False
        Exit Sub
        End If
           Exit Do
        
    End If
DtaHistorico.Recordset.MoveNext
Loop
FechaHoy = Format(Now, "dd/mm/yyyy")

SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[HorasExtras]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos] AS TotalBruto, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[HorasExtras]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]-[Deducciones]-[Prestamo]-[MontoINSS]-[MontoIR] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
DtaNominas.RecordSource = SqlNominas
DtaNominas.Refresh

If Not DtaNominas.Recordset.EOF Then
   DtaNominas.Recordset.MoveLast
   FechaUltNomina = DtaNominas.Recordset("FechaNomina")
   TxtUltFechaNomina.Value = FechaUltNomina
   NumFecha1 = FechaUltNomina
Else
   TxtUltFechaNomina = TxtFechaContrato
   'MsgBox "No ha sido Grabada Ninguna nómina a este empleado, Se le realizará la baja desde su contrato"
End If







'///////////Busco la Fecha para la Busqueda////////////////////////////

NumeroEmpleado = Me.TxtCodEmpleado.Text

If Me.ChkSueldoActual.Value = xtpUnchecked Then
SqlSalarios = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo,SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO,SUM(DetalleNomina.Comisiones) As Comisiones " & _
              "FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
              "Having (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.Comisiones) <> 0) And (DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") " & _
              "ORDER BY Nomina.Ano, Nomina.Mes "
Else
SqlSalarios = "SELECT DISTINCT TOP (100) PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SeptimoDia) AS Septimos, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,  SUM(DetalleNomina.Destajo + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones + DetalleNomina.Incentivos) + AVG(Historico.SueldoActual) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Comisiones) AS Comisiones, AVG(Historico.SueldoActual) AS SalarioBasico FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
              "Having (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.Comisiones) <> 0) And (DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") ORDER BY AÑO, Nomina.Mes"
End If



Me.DtaConsulta.RecordSource = SqlSalarios
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
 Me.DtaConsulta.Recordset.MoveLast
Else
 FechaHistorico = Format(Now, "dd/mm/yyyy")
 FechaBusqueda = Format(Now, "dd/mm/yyyy")
End If
i = 0
Do While Not Me.DtaConsulta.Recordset.BOF
  If i = 1 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")

  ElseIf i = 5 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    Exit Do
  ElseIf i = 0 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    FechaHistorico = Me.DtaConsulta.Recordset("FechaFin")
  Else
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
  End If
  i = i + 1

  Me.DtaConsulta.Recordset.MovePrevious
Loop


FechaEgreso = Me.TxtUltFechaNomina.Value

FechaContrato = Me.TxtFechaContrato.Text

'//////////SUMO 1 PARA AJUSTAR QUE SIEMPRE DA 1 DIA MENOS//////
'annos = CDbl(FechaEgreso) - CDbl(FechaContrato) + 1
MDIPrimero.DtaControles.Refresh
DiasMes = MDIPrimero.DtaControles.Recordset("DiasMes")

'annos = CalcularDiasVaca(FechaContrato, FechaEgreso)
Dias = CalcularDiasAntiguedad(FechaContrato, FechaEgreso) / 0.083333
annos = Dias
TxtAnnos.Text = Format(annos / 365, "###,##0.00")
TxtMeses.Text = Format(annos / DiasMes, "###,##0.00")
Dias = Format(annos * 365, "###,##0.00")
    Me.CmdEfectuar.Enabled = False
    Me.CmdRenovar.Enabled = False
    
DiasReales = CalcularDiasVaca(FechaBusqueda, FechaHistorico)
MesReal = DiasReales / DiasMes
'If MesReal > 6 Then
' MesReal = 6
'End If
    
'Me.TxtDiasTrabajados.Text = Format(annos, "###,##0")
'Me.TxtDiasTrabajados.Text = CalcularDiasVaca(FechaContrato, FechaEgreso)

Me.TxtDiasTrabajados.Text = CalcularDiasAntiguedad(FechaContrato, FechaEgreso) / 0.083333


Año = Year(FechaBusqueda)
Mes = Month(FechaBusqueda)



If Me.AdoEmpresa.Recordset("FormatoNomina") = "Nomina Bono Produccion" Then

   
    If Me.ChkExtra.Value = xtpUnchecked Then
    
       SqlSalarios = "SELECT DISTINCT " & _
                      "TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Antiguedad) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, " & _
                      "SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, " & _
                      "SUM(DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion) AS Incentivos, " & _
                      "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo  + DetalleNomina.BonoProduccion + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.HorasExtras + DetalleNomina.IncetivoProduccion + DetalleNomina.Antiguedad) AS TotalIngresos, " & _
                      "MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, " & _
                      "SUM(DetalleNomina.BonoProduccion) AS BonoProduccion, SUM(DetalleNomina.HorasExtras) AS HorasExtras " & _
                      "FROM         DetalleNomina INNER JOIN " & _
                      "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina " & _
                      "GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                      "HAVING      (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) " & _
                      "Between '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') " & _
                      "ORDER BY Nomina.Ano, Nomina.Mes "
               
               Me.TDBGridBonos.Visible = True
               Me.TDBGridSalarios.Visible = False
'               Me.TDBGridBonos.Columns(3).Visible = True
      Else
       SqlSalarios = "SELECT DISTINCT " & _
                      "TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Antiguedad) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, " & _
                      "SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, " & _
                      "SUM(DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion) AS Incentivos, " & _
                      "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo  + DetalleNomina.BonoProduccion + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.IncetivoProduccion + DetalleNomina.Antiguedad) AS TotalIngresos, " & _
                      "MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, " & _
                      "SUM(DetalleNomina.BonoProduccion) AS BonoProduccion " & _
                      "FROM         DetalleNomina INNER JOIN " & _
                      "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina " & _
                      "GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                      "HAVING      (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) " & _
                      "Between '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') " & _
                      "ORDER BY Nomina.Ano, Nomina.Mes "
               
               Me.TDBGridBonos.Visible = True
               Me.TDBGridSalarios.Visible = False
'               Me.TDBGridBonos.Columns(3).Visible = False
      
      
      End If
Else
  
  If Me.ChkSueldoActual.Value = xtpUnchecked Then
       ' SqlSalarios = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Antiguedad) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo,SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.Antiguedad)AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO,SUM(DetalleNomina.Comisiones) As Comisiones " &
                      '"FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                     ' "HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) Between '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') " & _
                       '"ORDER BY Nomina.Ano, Nomina.Mes "
      If Me.ChkIncentivos.Value = xtpUnchecked Then
            SqlSalarios = "SELECT DISTINCT"
            SqlSalarios = SqlSalarios + "  TOP (100) PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Antiguedad) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo,"
            SqlSalarios = SqlSalarios + "    SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(0) AS Incentivos,"
            SqlSalarios = SqlSalarios + "  SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso +  DetalleNomina.Antiguedad)"
            SqlSalarios = SqlSalarios + "   AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Reembolso) AS Reembolso,"
            SqlSalarios = SqlSalarios + "    Empleado.SueldoPeriodo"
            SqlSalarios = SqlSalarios + " FROM         DetalleNomina INNER JOIN"
            SqlSalarios = SqlSalarios + "   Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN"
            SqlSalarios = SqlSalarios + "  Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado"
            SqlSalarios = SqlSalarios + " GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano, Empleado.SueldoPeriodo"
            SqlSalarios = SqlSalarios + " HAVING      (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "')"
            SqlSalarios = SqlSalarios + " ORDER BY AÑO, Nomina.Mes"
      
      
      
      
      Else
            SqlSalarios = "SELECT DISTINCT"
            SqlSalarios = SqlSalarios + "  TOP (100) PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Antiguedad) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo,"
            SqlSalarios = SqlSalarios + "    SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,"
            SqlSalarios = SqlSalarios + "  SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso + DetalleNomina.Incentivos + DetalleNomina.Antiguedad)"
            SqlSalarios = SqlSalarios + "   AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Reembolso) AS Reembolso,"
            SqlSalarios = SqlSalarios + "    Empleado.SueldoPeriodo"
            SqlSalarios = SqlSalarios + " FROM         DetalleNomina INNER JOIN"
            SqlSalarios = SqlSalarios + "   Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN"
            SqlSalarios = SqlSalarios + "  Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado"
            SqlSalarios = SqlSalarios + " GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano, Empleado.SueldoPeriodo"
            SqlSalarios = SqlSalarios + " HAVING      (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "')"
            SqlSalarios = SqlSalarios + " ORDER BY AÑO, Nomina.Mes"
      End If
        
   Else
   
      If Me.ChkIncentivos.Value = xtpUnchecked Then
   
        SqlSalarios = "SELECT DISTINCT DetalleNomina.CodEmpleado, AVG(Historico.SueldoActual) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SeptimoDia * 0) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(0) AS Incentivos, SUM(DetalleNomina.Destajo + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso) + AVG(Historico.SueldoActual) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Reembolso) AS Reembolso FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano  " & _
                      "HAVING  (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') ORDER BY AÑO, Nomina.Mes"
      Else
         SqlSalarios = "SELECT DISTINCT TOP (100) PERCENT DetalleNomina.CodEmpleado, AVG(Historico.SueldoActual) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SeptimoDia * 0) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,  SUM(DetalleNomina.Destajo + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso + DetalleNomina.Incentivos) + AVG(Historico.SueldoActual) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Reembolso) As Reembolso FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                      "HAVING (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') ORDER BY AÑO, Nomina.Mes"
      
      End If
   End If
    
       Me.TDBGridBonos.Visible = False
       Me.TDBGridSalarios.Visible = True
       



End If


FechaBusqueda1 = FechaBusqueda
FechaHistorico1 = FechaHistorico


Me.AdoSalarios.RecordSource = SqlSalarios
Me.AdoSalarios.Refresh


If SueldoFijo = True Then

If Me.AdoSalarios.Recordset.EOF Then
  SueldoPeriodo = 0
  
  
  
Else
  Me.AdoSalarios.Recordset.MoveLast
  'SueldoPeriodo = Me.AdoSalarios.Recordset("TotalIngresos")
  
  If DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
    If IsNull(Me.AdoSalarios.Recordset("SueldoPeriodo")) Then
        SueldoPeriodo = 0
    Else
     SueldoPeriodo = Me.AdoSalarios.Recordset("SueldoPeriodo") * 2
    End If
  
  
   
  ElseIf DtaTipoNomina.Recordset("Periodo") = "Mensual" Then
    SueldoPeriodo = Me.AdoSalarios.Recordset("SueldoPeriodo")
  ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
    If Me.ChkSueldoActual.Value = xtpUnchecked Then
      SueldoPeriodo = (Me.AdoSalarios.Recordset("SueldoPeriodo") / 14) * DiasMes
    Else
       SueldoPeriodo = Me.AdoSalarios.Recordset("SalarioBasico")
    End If
  End If
  
  'SueldoPeriodo = 0
End If



 Me.TxtSalarios.Caption = "Empleado con Salario Fijo"
   
    '/////////////VERIFICO SI SE UTILIZA LA ANTIGUEDAD COMO BASE//////////////////////
    '/////////////PARA EL CALCULO DE LA LIQUIDACION///////////////////////////////////
       If ChkAntiguedad.Value = 1 Then
        Años = Int(Me.TxtAnnos.Text)
        Me.AdoAntiguedad.RecordSource = "SELECT años_acum, porcent From Antiguedad Where (años_acum = " & Años & ")"
        Me.AdoAntiguedad.Refresh
        If Not Me.AdoAntiguedad.Recordset.EOF Then
         PAntiguedad = 1 + Me.AdoAntiguedad.Recordset("porcent")
         Me.txtAntiguedad.Text = Me.AdoAntiguedad.Recordset("porcent")
        Else
         Me.txtAntiguedad.Text = 0
        End If
        SalarioPromedio = SueldoPeriodo * PAntiguedad
        SalarioAlto = SueldoPeriodo * PAntiguedad
         
       Else
        SalarioPromedio = SueldoPeriodo
        SalarioAlto = SueldoPeriodo
         Me.txtAntiguedad.Text = 0
       End If
 
  
Else
    Me.txtAntiguedad.Text = 0
    Me.TxtSalarios.Caption = "Empleado con Salario Variable"
    Contador = 0
    TotalSalario = 0
    Salario = 0
    SalarioAlto = 0
    Do While Not Me.AdoSalarios.Recordset.EOF
    
      If Not IsNull(Me.AdoSalarios.Recordset("TotalIngresos")) Then
        TotalSalario = TotalSalario + Me.AdoSalarios.Recordset("TotalIngresos")
        Salario = Me.AdoSalarios.Recordset("TotalIngresos")
      Else
        Salario = 0
      End If
 
        If Salario > SalarioAlto Then
            SalarioAlto = Salario
        End If
 
        Contador = Contador + 1
        Me.AdoSalarios.Recordset.MoveNext
    Loop
   
   If Not Contador = 0 Then
    If Me.Check1.Value = 0 Then
    SalarioPromedio = TotalSalario / Contador '//////esto divide por los meses enteres //////
    
    Else
    
        If Contador < 6 Then
'            MDIPrimero.DtaEmpresa.Refresh
'            If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
'             If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("SalarioMinimo")) Then
'               SalarioMinimo = MDIPrimero.DtaEmpresa.Recordset("SalarioMinimo")
'             End If
'             TarifaHoraria = (SalarioMinimo / 30) / 8
'            End If
        
        
        
    '     SalarioPromedio = TotalSalario / Contador         'MesReal  '/////ESTO LO DIVIDO POR EL TIEMPO REAL TRABAJADO /////
           SalarioPromedio = SueldoActual
          
         Else
           SalarioPromedio = TotalSalario / Contador
         
         End If
      
    End If
   End If

 End If
 
    Me.TxtSalarioPromedio.Text = Format(SalarioPromedio, "##,##0.00")
    Me.TxtSalarioAlto.Text = Format(SalarioAlto, "##,##0.00")
    
    
     If Me.ChkSueldoActual.Value = xtpChecked Then
      Me.AdoSalarios.Refresh
      If Not Me.AdoSalarios.Recordset.EOF Then
'         SalarioPromedio = Me.AdoSalarios.Recordset("TotalIngresos")
'         Me.TxtSalarioPromedio.Text = Format(SalarioPromedio, "##,##0.00")
'         Me.TxtSalarioAlto.Text = Format(SalarioPromedio, "##,##0.00")
         
      End If
     End If
    
    Dim AñoActual As Integer ', CodTipoNomina As String
    
    CodigoEmpleado = Me.TxtCodEmpleado.Text


'/////////CONSULTA EL SALARIO Y TIPO DE NOMINA DEL EMPLEADO//////////////////////////

 sql = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento," & vbLf
 sql = sql & "Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.OtrosIngresos, Empleado.DescripOtrIngre," & vbLf
 sql = sql & "Empleado.ExentoIr, Empleado.PagoInssPatronal, Empleado.Activo, Empleado.Liquidado, Empleado.Ausente, Empleado.SalarioFijo," & vbLf
 sql = sql & "Empleado.SumarSubsidio , Empleado.PorcientoIncentivo, Empleado.Gravidez, TipoNomina.Periodo" & vbLf
 sql = sql & "FROM Empleado INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina" & vbLf
 sql = sql & "WHERE     (Empleado.CodEmpleado = '" & CodigoEmpleado & "') AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
 Me.DtaConsulta.RecordSource = sql
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  TipoNomina = Me.DtaConsulta.Recordset("Periodo")
  CodTipoNomina = Me.DtaConsulta.Recordset("CodTipoNomina")
 Else
  MsgBox "Este Empleado no Existe", vbCritical, "Sistema de Nominas"
  Exit Sub
 End If
    
    
    
    
       '/////////////BUSCO EL INICIO DEL PERIODO/////////////////////
       AñoActual = Year(Me.TxtUltFechaNomina.Value)
       Me.AdoInicioAño.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (año = " & AñoActual & ") AND (Periodo = 1)"
       Me.AdoInicioAño.Refresh
       If Not Me.AdoInicioAño.Recordset.EOF Then
        FechaInicio = Me.AdoInicioAño.Recordset("Inicio")
       End If
    
    
       '//////////////////////////////////////////////////////////////////
       '//////////CALCULO CUANTOS DIAS TIENE TRABAJADOS////////////////////
       '////////////////////////////////////////////////////////////////////
    Dim Fecha1 As Date, Fecha2 As Date
       
       Fecha1 = DateSerial(Year(FechaEgreso), 5, 31)
       FechaEgreso = Me.TxtUltFechaNomina.Value
       If FechaContrato < FechaInicio Then
        FechaInicioAgui = DateSerial(Year(FechaEgreso) - 1, 12, 1)
        Me.DTPFechaIniAgui.Value = FechaInicioAgui
        If FechaEgreso > Fecha1 Then
          Me.DTPFechaIniVaca.Value = DateSerial(Year(FechaEgreso), 6, 1)
        Else
          Me.DTPFechaIniVaca.Value = DateSerial(Year(FechaEgreso) - 1, 12, 1)
        End If
       Else
       FechaInicioAgui = FechaContrato
       Me.DTPFechaIniAgui.Value = FechaInicioAgui
       Me.DTPFechaIniVaca.Value = FechaContrato
       End If

End Sub



Private Sub ChkAntiguedad_Click()
Dim FechaEgreso As Date, FechaContrato As Date, Año As Integer, Mes As Integer, i As Integer
Dim FechaBusqueda As Date, TotalSalario As Double, SalarioPromedio As Double, Contador As Integer
Dim SqlSalarios As String, SalarioAlto As Double, Salario As Double, FechaHistorico As Date, NumeroEmpleado As Double
FechaEgreso = Me.TxtUltFechaNomina.Value
FechaContrato = Me.TxtFechaContrato.Text
'//////////SUMO 1 PARA AJUSTAR QUE SIEMPRE DA 1 DIA MENOS//////
'annos = CDbl(FechaEgreso) - CDbl(FechaContrato) + 1
Dias = CalcularDiasAntiguedad(FechaContrato, FechaEgreso) / 0.083333
annos = Dias
TxtAnnos.Text = Format(annos / 365, "###,##0.00")
TxtMeses.Text = Format(annos / 30.41, "###,##0.00")
Me.TxtDiasTrabajados.Text = Format(annos, "###,##0")
Dias = annos
Me.CmdEfectuar.Enabled = False

'///////////Busco la Fecha para la Busqueda////////////////////////////

NumeroEmpleado = Me.TxtCodEmpleado.Text

SqlSalarios = "SELECT DISTINCT" & vbLf
SqlSalarios = SqlSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
SqlSalarios = SqlSalarios & "AS Destajo, SUM(dbo.DetalleNomina.Incentivos) AS Incentivos, SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo)" & vbLf
SqlSalarios = SqlSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes AS MES," & vbLf
SqlSalarios = SqlSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
SqlSalarios = SqlSalarios & "FROM   dbo.DetalleNomina INNER JOIN" & vbLf
SqlSalarios = SqlSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
SqlSalarios = SqlSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
SqlSalarios = SqlSalarios & "Having (dbo.DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") And (Sum(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0)"
Me.DtaConsulta.RecordSource = SqlSalarios
Me.DtaConsulta.Refresh
Me.DtaConsulta.Recordset.MoveLast
i = 0
Do While Not Me.DtaConsulta.Recordset.BOF
  If i = 1 Then
    FechaHistorico = Me.DtaConsulta.Recordset("FechaFin")

  ElseIf i = 6 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    Exit Do
  ElseIf i = 0 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    FechaHistorico = Me.DtaConsulta.Recordset("FechaFin")
  Else
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
  End If
  i = i + 1

  Me.DtaConsulta.Recordset.MovePrevious
Loop


FechaEgreso = Me.TxtUltFechaNomina.Value
'FechaHistorico = DateSerial(Year(FechaEgreso), Month(FechaEgreso), 1 - 1)
FechaContrato = Me.TxtFechaContrato.Text
'FechaBusqueda = DateSerial(Year(FechaEgreso), Month(FechaEgreso) - 6, 1)
Año = Year(FechaBusqueda)
Mes = Month(FechaBusqueda)

    SqlSalarios = "SELECT DISTINCT" & vbLf
    SqlSalarios = SqlSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
    SqlSalarios = SqlSalarios & "AS Destajo, SUM(dbo.DetalleNomina.SeptimoDia) AS Septimo, SUM(dbo.DetalleNomina.OtrosIngresos) AS Otros, SUM(dbo.DetalleNomina.Incentivos)" & vbLf
    SqlSalarios = SqlSalarios & "AS Incentivos," & vbLf
    SqlSalarios = SqlSalarios & "SUM (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.OtrosIngresos)" & vbLf
    SqlSalarios = SqlSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes," & vbLf
    SqlSalarios = SqlSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
    SqlSalarios = SqlSalarios & "FROM    dbo.DetalleNomina INNER JOIN" & vbLf
    SqlSalarios = SqlSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
    SqlSalarios = SqlSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
    SqlSalarios = SqlSalarios & "HAVING(SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0) And (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND" & vbLf
    SqlSalarios = SqlSalarios & "'" & Format(FechaHistorico, "yyyymmdd") & "')" & vbLf
    SqlSalarios = SqlSalarios & "ORDER BY dbo.Nomina.Ano, dbo.Nomina.Mes"

Me.AdoSalarios.RecordSource = SqlSalarios
Me.AdoSalarios.Refresh


If SueldoFijo = True Then
 
 If Me.AdoSalarios.Recordset.EOF Then
  SueldoPeriodo = 0
 Else
  Me.AdoSalarios.Recordset.MoveLast
  SueldoPeriodo = Me.AdoSalarios.Recordset("TotalIngresos")
 End If
 Me.TxtSalarios.Caption = "Empleado con Salario Fijo"
   
    '/////////////VERIFICO SI SE UTILIZA LA ANTIGUEDAD COMO BASE//////////////////////
    '/////////////PARA EL CALCULO DE LA LIQUIDACION///////////////////////////////////
       If ChkAntiguedad.Value = 1 Then
        Años = Int(Me.TxtAnnos.Text)
        Me.AdoAntiguedad.RecordSource = "SELECT años_acum, porcent From Antiguedad Where (años_acum = " & Años & ")"
        Me.AdoAntiguedad.Refresh
        If Not Me.AdoAntiguedad.Recordset.EOF Then
         PAntiguedad = 1 + Me.AdoAntiguedad.Recordset("porcent")
         Me.txtAntiguedad.Text = Me.AdoAntiguedad.Recordset("porcent")
        Else
         Me.txtAntiguedad.Text = 0
        End If
        SalarioPromedio = SueldoPeriodo * PAntiguedad
        SalarioAlto = SueldoPeriodo * PAntiguedad
         
       Else
        SalarioPromedio = SueldoPeriodo
        SalarioAlto = SueldoPeriodo
         Me.txtAntiguedad.Text = 0
       End If
 
  
Else
    Me.txtAntiguedad.Text = 0
    Me.TxtSalarios.Caption = "Empleado con Salario Variable"
    Contador = 0
    TotalSalario = 0
    Salario = 0
    SalarioAlto = 0
    Do While Not Me.AdoSalarios.Recordset.EOF
        TotalSalario = TotalSalario + Me.AdoSalarios.Recordset("TotalIngresos")
        Salario = Me.AdoSalarios.Recordset("TotalIngresos")
 
        If Salario > SalarioAlto Then
            SalarioAlto = Salario
        End If
 
        Contador = Contador + 1
        Me.AdoSalarios.Recordset.MoveNext
    Loop

    SalarioPromedio = TotalSalario / Contador

 End If
    Me.TxtSalarioPromedio.Text = Format(SalarioPromedio, "##,##0.00")
    Me.TxtSalarioAlto.Text = Format(SalarioAlto, "##,##0.00")

    Dim AñoActual As Integer, CodTipoNomina As String
    
    CodigoEmpleado = Me.TxtCodEmpleado.Text


'/////////CONSULTA EL SALARIO Y TIPO DE NOMINA DEL EMPLEADO//////////////////////////

 sql = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento," & vbLf
 sql = sql & "Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.OtrosIngresos, Empleado.DescripOtrIngre," & vbLf
 sql = sql & "Empleado.ExentoIr, Empleado.PagoInssPatronal, Empleado.Activo, Empleado.Liquidado, Empleado.Ausente, Empleado.SalarioFijo," & vbLf
 sql = sql & "Empleado.SumarSubsidio , Empleado.PorcientoIncentivo, Empleado.Gravidez, TipoNomina.Periodo" & vbLf
 sql = sql & "FROM Empleado INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina" & vbLf
 sql = sql & "WHERE     (Empleado.CodEmpleado = '" & CodigoEmpleado & "') AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
 Me.DtaConsulta.RecordSource = sql
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  TipoNomina = Me.DtaConsulta.Recordset("Periodo")
  CodTipoNomina = Me.DtaConsulta.Recordset("CodTipoNomina")
 Else
  MsgBox "Este Empleado no Existe", vbCritical, "Sistema de Nominas"
  Exit Sub
 End If
    
    
    
    
       '/////////////BUSCO EL INICIO DEL PERIODO/////////////////////
       AñoActual = Year(Me.TxtUltFechaNomina.Value)
       Me.AdoInicioAño.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (año = " & AñoActual & ") AND (Periodo = 1)"
       Me.AdoInicioAño.Refresh
       If Not Me.AdoInicioAño.Recordset.EOF Then
        FechaInicio = Me.AdoInicioAño.Recordset("Inicio")
       End If
    
    
       '//////////////////////////////////////////////////////////////////
       '//////////CALCULO CUANTOS DIAS TIENE TRABAJADOS////////////////////
       '////////////////////////////////////////////////////////////////////
    Dim Fecha1 As Date, Fecha2 As Date
       
       Fecha1 = DateSerial(Year(FechaEgreso), 5, 31)
       FechaEgreso = Me.TxtUltFechaNomina.Value
       If FechaContrato < FechaInicio Then
        FechaInicioAgui = DateSerial(Year(FechaEgreso) - 1, 12, 1)
        Me.DTPFechaIniAgui.Value = FechaInicioAgui
        If FechaEgreso > Fecha1 Then
          Me.DTPFechaIniVaca.Value = DateSerial(Year(FechaEgreso), 6, 1)
        Else
          Me.DTPFechaIniVaca.Value = DateSerial(Year(FechaEgreso) - 1, 12, 1)
        End If
       Else
       FechaInicioAgui = FechaContrato
       Me.DTPFechaIniAgui.Value = FechaInicioAgui
       Me.DTPFechaIniVaca.Value = FechaContrato
       End If


End Sub

Private Sub ChkExtra_Click()
Dim CodigoEmpleado As String
CodigoEmpleado = TxtCodEmpleado.Text

Me.TxtCodEmpleado.Text = 0
Me.TxtCodEmpleado.Text = CodigoEmpleado



End Sub



Private Sub ChkIncentivos_Click()
LlenarLiquidacion
End Sub

Private Sub CmdCalcular_Click()
Dim Sueldo As Double
Dim Fecha As String
Dim i As Integer, Espacio As String
Dim SalMayor As Double, Año As Integer
Dim SalTemp As Double, Meses As Integer
Dim SalBrutoTemp As Double, j As Integer
Dim SalBrutoMayor As Double, H As Integer
Dim Mes As Byte, SqlHrsExtras As String
Dim DiaMes As Double, HE As Integer
Dim DiaSemana As Double, DiasAntiguedad As Integer
Dim Mes13 As Double
Dim VACACIONES As Double, SalarioMensual As Double
Dim MontoAntiguedad As Double, MontoInss As Double, MontoInssPatronal As Double
Dim MontoHRSExtras As Double, MontoHora As Double
Dim TextOtro As String, TotalIngresos As Double
Dim Otro As Double, TotalEgresos As Double, NetoPagar As Double
Dim FechaAntiguedad As Date
Dim DiferenciaAnnos As Double
Dim Prestamo As Double
Dim Deducciones As Double
Dim MontoNomPropor As Double 'salario de nomina proporcional
Dim Fechass As Date

Espacio = " "
CodEmpleado = TxtCodEmpleado.Text
'extraigo el salario mayor de los ultimos seis meses
SalMayor = 0
SalBrutoMayor = 0
CantRegistros = 0
Anno = Year(Now)
MesActual = Month(Now)
CodEmpleado = TxtCodEmpleado.Text

Fechass = (Format(Now, "dd/mm/yyyy"))
        
'SQLS
        

        
 


'//////Hago un Ciclo para los ultimos 5 años////////////////
For i = 0 To 5
  
If TxtMotivo.Text = "" Then
   MsgBox "El motivo de la baja no puede quedar vacío"
   TxtMotivo.SetFocus
   Exit Sub
End If
  MesActual = Month(Now)
'////////Verifico si el sueldo fijo es Variable////////
If SueldoFijo = False Then
   If MesActual > 6 Then
            Meses = MesActual - 6
   Else
           Meses = 1
   End If
        For Mes = Meses To MesActual
        SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where ((Month([Nomina].[FechaNomina])) =  '" & Mes & "' ) And ((Year([Nomina].[FechaNomina])) ='" & Anno & "') and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
            DtaNominas.RecordSource = SqlNominas
            DtaNominas.Refresh
               SalTemp = 0
              Do While Not DtaNominas.Recordset.EOF
               SalTemp = DtaNominas.Recordset("Total")
               CantRegistros = CantRegistros + 1
               DtaNominas.Recordset.MoveNext
              Loop
        If SalMayor < SalTemp Then
           SalMayor = SalTemp
        End If
  
       Next
        
             
  Else
'///Asigno el salario basico debengado///////////
  SalMayor = SalarioBasico
  Exit For
  
End If
            
  Next i




DtaControles.Refresh
DiasMes = DtaControles.Recordset("DiasMes")
DiasSemana = DtaControles.Recordset("DiasSemana")

 '///////////////////NOMINAS//////////////////////////
 SqlNominas = "SELECT Nomina.NumNomina, Nomina.CodTipoNomina, DetalleNomina.CodEmpleado FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
 DtaNominas.RecordSource = SqlNominas
 DtaNominas.Refresh

 CodTipoNomina = DtaNominas.Recordset("CodTipoNomina")

 DtaTipoNomina.Refresh
 Do While Not DtaTipoNomina.Recordset.EOF
   If DtaTipoNomina.Recordset("CodTipoNomina") = CodTipoNomina Then
   Exit Do
   End If
 DtaTipoNomina.Recordset.MoveNext
 Loop

 

'dependiendo del tipo de pago se hace el calculo del salario básico
     If DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
           SalMayor = SalMayor / 3
           SalBrutoMayor = SalBrutoMayor / 3
     ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
           SalMayor = SalMayor / 6
           SalBrutoMayor = SalBrutoMayor / 6

     ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
        SalMayor = SalMayor * 2
        SalBrutoMayor = SalarioBasico * 2
     ElseIf DtaTipoNomina.Recordset("Periodo") = "Semanal" Then
        SalMayor = SalMayor * 4
        SalBrutoMayor = SalarioBasico * 4
     End If






'MsgBox ("El Salario Mayor Bruto Mensual de los ultimos seis meses es: " & Str(SalBrutoMayor))
'MsgBox ("El Salario Mayor Mensual  es: " & Str(Format(SalMayor, "##,##0.00")))

'//////////Calculo el Factor///////////////////
CantMes = Month(NumFecha1)
Factor = (Month(NumFecha1) + 1) / 12

'*******************************************************
'******************************************************
'////////creo el pago de las prestaciones////////////
'*******************************************************
'*******************************************************



'///////////////////////////////////////////////////////////
'////////////VERIFIO EL CALCULO DE LA ANTIGUEDAD/////////////////////
'///////////////////////////////////////////////////////////////////

If Me.ChkAntigue.Value = 1 Then
 Mes = Month(Now)
  Año = Year(Now)
  
    

  DiasAntiguedad = 0
  DtaHistorico.RecordSource = "SELECT Historico.Codempleado, Historico.FechaBaja, Historico.FechaContrato From Historico Where (((Historico.CodEmpleado) = '" & CodEmpleado & "'))"
  DtaHistorico.Refresh
  
   
'////////Verifico cuantos dias tiene de trabajar///////////////////
  If Not DtaHistorico.Recordset.EOF Then
   If Not IsNull(DtaHistorico.Recordset("FechaContrato")) Then
     FechaContrato = DtaHistorico.Recordset("FechaContrato")
     NumFecha2 = FechaContrato
     NumFecha1 = Now
     
     annos = (CDbl(NumFecha1) - CDbl(NumFecha2)) / 365
     CantMeses = annos * 12
 '////////////Si es tiene menos de un años de trajabar///////////////////////
 
     If annos >= 1 And annos <= 3 Then
         DiasAntiguedad = CantMeses * 2.5
     ElseIf annos > 3 And annos < 6 Then
     'Resto los 3 años en meses
     CantMeses = CantMeses - 36
'Esto es equivalente para los 3 primeros años 30/12 y para los ultimos 3 20/12
       DiasAntiguedad = (36 * 2.5) + (CantMeses * 1.6666667)
     ElseIf annos >= 6 Then
 'Esto es lo maximo a recibir, por trajador
        DiasAntiguedad = 150
     
     End If
    End If
   End If

'///////Calculo el salario proporcional del 13vo mes////////////
  MontoAntiguedad = (SalMayor * DiasAntiguedad) / DiasMes
   
Else
    DiasAntiguedad = 0
    MontoAntiguedad = 0
End If







'//////////////////////////////////////////////////////////////////////
'/////////VERIFICO SI SE CALCULA EL 13VO MES////////////////
'///////////////////////////////////////////////////////////////////////
If Chk13mes.Value = 1 Then


  Mes = Month(Now)
  Año = Year(Now)
  
  FechaIniAgui = CDate("01/12/" & Str(Year(Now) - 1))
  FechaFinAgui = Format(Now, "dd/mm/yyyy")
  
  

  Dias = 0
  DtaHistorico.RecordSource = "SELECT Historico.Codempleado, Historico.FechaBaja, Historico.FechaContrato From Historico Where (((Historico.CodEmpleado) = '" & CodEmpleado & "'))"
  DtaHistorico.Refresh
  
'////////////////////////////////////////////////////////////////////
'/////////////Busco el Adelanto de 13vo mes Registrados//////////////
'////////////////////////////////////////////////////////////////////////

NumFecha1 = FechaIniAgui
NumFecha2 = FechaFinAgui
        Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo WHERE (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Adelanto13vo.FechaAdelanto) Between  " & NumFecha1 & " And " & NumFecha2 & ") AND ((Adelanto13vo.TipoAdelanto)='13vo Mes'))"
        Me.DtaAdelanto.Refresh
        Adelanto13vo = 0
        Do While Not DtaAdelanto.Recordset.EOF
         Adelanto13vo = Adelanto13vo + DtaAdelanto.Recordset("MontoAdelanto")
         DtaAdelanto.Recordset.MoveNext
        Loop
   
'////////Verifico cuantos dias tiene de trabajar///////////////////
  If Not DtaHistorico.Recordset.EOF Then
   If Not IsNull(DtaHistorico.Recordset("FechaContrato")) Then
     FechaContrato = DtaHistorico.Recordset("FechaContrato")
     NumFecha2 = FechaContrato
     NumFecha1 = Now
     
     annos = (CDbl(NumFecha1) - CDbl(NumFecha2)) / 365
     CantMeses = annos * 12
 '////////////Si es tiene menos de un años de trajabar///////////////////////
 
     If CantMeses <= 12 Then
      Dias = CantMeses * 2.5
     Else

      NumFecha2 = CDate("01/12/" & Anno - 1)
      annos = (CDbl(NumFecha1) - CDbl(NumFecha2)) / 365
      CantMeses = Format(annos * 12, "##,##0.0000")
      If CantMeses > 12 Then
       Dias = 12 * 2.5
      Else
       Dias = CantMeses * 2.5
      End If
     End If
    End If
   End If

'///////Calculo el salario proporcional del 13vo mes////////////
   Mes13 = (SalMayor * Dias) / DiasMes
   
Else
    Dias = 0
    Mes13 = 0
End If



'//////////////////////////////////////////////////////////////////////
'/////////INICIO EL CALCULO DE LAS VACACIONES CORRESPONDIENTES/////////
'/////////////Calculo las Vacaciones////////////////////////////////////
If ChkVaca.Value = 1 Then
 SalMayor = 0
 CantMeses = 0
 CantRegistros = 0
 
 
 
 
 
 If Me.DtaEmpleado.Recordset("SalarioFijo") = "N" Then
 '///////////Si el Salario es Variable Busco el Salario Mayor/////////
    If Month(Now) <= 6 Then
        'Enero - Junio
        SalMayor = 0
        For Mes = 1 To 6
        SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where ((Month([Nomina].[FechaNomina])) =  '" & Mes & "' ) And ((Year([Nomina].[FechaNomina])) ='" & Anno & "') and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
            DtaNominas.RecordSource = SqlNominas
            DtaNominas.Refresh
            Fecha = "01/01/" & Anno
        FechaIniVaca = Fecha
        NumFecha1 = Fecha
        Fecha = "30/06/" & Anno
        FechaFinVaca = Fecha
        NumFecha2 = Fecha
'//////////////////////////////////////////////////////////////////////
'/////////////Busco el Adelanto de Vacaciones Registrados//////////////
'////////////////////////////////////////////////////////////////////////
        
         
        
        Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo WHERE (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Adelanto13vo.FechaAdelanto) Between  " & NumFecha1 & " And " & NumFecha2 & ") AND ((Adelanto13vo.TipoAdelanto)='Vacaciones'))"
        Me.DtaAdelanto.Refresh
        AdelantoVaca = 0
        
        Do While Not DtaAdelanto.Recordset.EOF
         AdelantoVaca = AdelantoVaca + DtaAdelanto.Recordset("MontoAdelanto")
         DtaAdelanto.Recordset.MoveNext
        Loop
        
        
        
               SalTemp = 0
               If Not DtaNominas.Recordset.EOF Then
                CantMeses = CantMeses + 1
               End If
              Do While Not DtaNominas.Recordset.EOF
               SalTemp = SalTemp + DtaNominas.Recordset("Total")
               CantRegistros = CantRegistros + 1
               DtaNominas.Recordset.MoveNext
              Loop
        
         SalMayor = SalTemp + SalMayor
  
       Next
    Else
    'Julio- Diciembre
    For Mes = 7 To 12
   
        SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where ((Month([Nomina].[FechaNomina])) =  '" & Mes & "' ) And ((Year([Nomina].[FechaNomina])) ='" & Anno & "') and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
            DtaNominas.RecordSource = SqlNominas
            DtaNominas.Refresh
               SalTemp = 0
               
        Fecha = "01/07/" & Anno
        FechaIniVaca = Fecha
        NumFecha1 = CDate(Fecha)
        Fecha = "31/12/" & Anno
        FechaFinVaca = Fecha
        NumFecha2 = CDate(Fecha)
        
'///////Busco el Adelanto de las Vacaciones/////////////////////
        
        'Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo WHERE (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Adelanto13vo.FechaAdelanto) Between " & NumFecha1 & " And " & NumFecha2 & "))"
        Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Last(Adelanto13vo.FechaAdelanto) AS ÚltimoDeFechaAdelanto, Sum(Adelanto13vo.MontoAdelanto) AS TotalMontoAdelanto, Last(Adelanto13vo.[Ref/Cheque]) AS [ÚltimoDeRef/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo GROUP BY Adelanto13vo.CodEmpleado, Adelanto13vo.TipoAdelanto HAVING (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Last(Adelanto13vo.FechaAdelanto)) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((Adelanto13vo.TipoAdelanto)='Vacaciones'))"
        Me.DtaAdelanto.Refresh
        AdelantoVaca = 0
        
        Do While Not DtaAdelanto.Recordset.EOF
         AdelantoVaca = AdelantoVaca + DtaAdelanto.Recordset("TotalMontoAdelanto")
         DtaAdelanto.Recordset.MoveNext
        Loop
        
               If Not DtaNominas.Recordset.EOF Then
                CantMeses = CantMeses + 1
               End If
              Do While Not DtaNominas.Recordset.EOF
               SalTemp = SalTemp + DtaNominas.Recordset("Total")
               CantRegistros = CantRegistros + 1
               DtaNominas.Recordset.MoveNext
              Loop
        
         SalMayor = SalTemp + SalMayor
        'If SalMayor < SalTemp Then
         '  SalMayor = SalTemp
        'End If
        
       Next
     
    End If
 Else '///Si es salario fijo le calculo el ultimo salario ////////
 
'/////////////////////////////////////////////////////////////////////////
'/////////////Busco el Adelanto de Vacaciones Registrados//////////////
'////////////////////////////////////////////////////////////////////////
 Mes = Month(Now)
  Año = Year(Now)
   If Mes <= 6 Then
   FechaIniVaca = CDate("01/01/" & Str(Año))
   FechaFinVaca = CDate("30/06/" & Str(Año))
  Else
   FechaIniVaca = CDate("01/07/" & Str(Año))
   FechaFinVaca = CDate("31/12/" & Str(Año))
  End If

NumFecha1 = FechaIniVaca
NumFecha2 = FechaFinVaca

        Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo WHERE (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Adelanto13vo.FechaAdelanto) Between  " & NumFecha1 & " And " & NumFecha2 & ") AND ((Adelanto13vo.TipoAdelanto)='Vacaciones'))"
        'Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo WHERE (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Adelanto13vo.FechaAdelanto) Between " & NumFecha1 & " And " & NumFecha2 & "))"
        Me.DtaAdelanto.Refresh
        AdelantoVaca = 0
        
        Do While Not DtaAdelanto.Recordset.EOF
         AdelantoVaca = AdelantoVaca + DtaAdelanto.Recordset("MontoAdelanto")
         DtaAdelanto.Recordset.MoveNext
        Loop
 
  
 
 
    SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where (((DetalleNomina.CodEmpleado) = '" & CodEmpleado & "'))"
    DtaNominas.RecordSource = SqlNominas
    DtaNominas.Refresh
    If DtaNominas.Recordset.EOF Then
     Edicion = False
     'DtaNominas.Recordset.MoveLast
    End If
     
  '///////Selecciono el Salario Mayor de la Tabla Empleados/////////////////
     SalMayor = DtaEmpleado.Recordset("SueldoPeriodo")
    
    'dependiendo del tipo de pago se hace el calculo del salario básico
    
    If DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
         SalMayor = SalMayor
    ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
          SalMayor = SalMayor
    ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
         SalMayor = SalMayor
    End If
    
     
    
    
   If Month(Now) <= 6 Then
      SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina WHERE (((DetalleNomina.CodEmpleado)='" & CodEmpleado & "') AND ((Month([Nomina].[FechaNomina])) Between 1 And 6) AND ((Year([Nomina].[FechaNomina]))= " & Anno & " ))"
      DtaNominas.RecordSource = SqlNominas
      DtaNominas.Refresh
      If Not DtaNominas.Recordset.EOF Then
       DtaNominas.Recordset.MoveLast
      End If
      CantRegistros = 0
      
   Else
   
     SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina WHERE (((DetalleNomina.CodEmpleado)='" & CodEmpleado & "') AND ((Month([Nomina].[FechaNomina])) Between 7 And 12) AND ((Year([Nomina].[FechaNomina]))= " & Anno & " ))"
      DtaNominas.RecordSource = SqlNominas
      DtaNominas.Refresh
     If Not DtaNominas.Recordset.EOF Then
      DtaNominas.Recordset.MoveLast
     End If
      CantRegistros = 0
   End If

   
 End If  '/////////Fin del If salario fijo/////////////////
        
'///////dependiendo del tipo de pago se hace el calculo del salario básico
     If DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
        Sueldo = DtaEmpleado.Recordset("SueldoPeriodo") / 3
         If CantRegistros > 0 Then
           SalMayor = SalMayor / CantMeses
         Else
           SalMayor = SalMayor / 3
           CantRegistros = DtaNominas.Recordset.RecordCount
         End If
     ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
       Sueldo = DtaEmpleado.Recordset("SueldoPeriodo") / 6
          If CantRegistros > 0 Then
           SalMayor = SalMayor / CantMeses
          Else
           SalMayor = SalMayor / 6
           CantRegistros = DtaNominas.Recordset.RecordCount
          End If
     ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
     Sueldo = DtaEmpleado.Recordset("SueldoPeriodo") * 2
         If CantRegistros > 0 Then
           SalMayor = (SalMayor / CantMeses)
         Else
           SalMayor = SalMayor * 2
           CantRegistros = DtaNominas.Recordset.RecordCount
          End If
     End If
   DiasDescuento = val(Me.TxtDescuentoDias.Text)
   VACACIONES = SalMayor * (((CantRegistros * 1.25) - DiasDescuento) / DiasMes)
Else
   'Dias = 0
   VACACIONES = 0
End If


'////////////////////////////////////////////////////////////////////
'////////Calculo el Salario para el Horas Extras//////////////////
'////////////////////////////////////////////////////////////////////
If Me.ChkExtra.Value = 1 Then
 '/////////////////////////////// HORAS EXTRAS /////////////////////////////////////////
        'calculo el monto por hora dependiendo del tipo de nomina
        
        SqlHrsExtras = "SELECT HorasExtras.CodEmpleado, HorasExtras.NumNomina, HorasExtras.CantHoras, HorasExtras.Pagada From HorasExtras WHERE  HorasExtras.Pagada=False AND HorasExtras.CodEmpleado='" & CodEmpleado & "'"
        DtaHrsExtras.RecordSource = SqlHrsExtras
        DtaHrsExtras.Refresh
        
        
        
        Select Case DtaTipoNomina.Recordset("Periodo")
        
        Case "Semanal Viernes"
            MontoHora = Format(SalarioBasico / (DiasSemana * 8), "###,##0.00")
        Case "Semanal Sabado"
            MontoHora = Format(SalarioBasico / (DiasSemana * 8), "###,##0.00")
        Case "Catorcenal los Viernes"
            MontoHora = Format(SalarioBasico / 112, "###,##0.00")
        Case "Catorcenal los Sabados"
            MontoHora = Format(SalarioBasico / 112, "###,##0.00")
        Case "Quincenal"
            MontoHora = Format(SalarioBasico / ((DiasMes * 8) / 2), "###,##0.00")
        Case "Mensual"
            MontoHora = Format(SalarioBasico / (DiasMes * 8), "###,##0.00")
        Case "Trimestral"
            MontoHora = Format(SalarioBasico / (DiasMes * 8 * 3), "###,##0.00")
        Case "Semestral"
            MontoHora = Format(SalarioBasico / (DiasMes * 8 * 6), "###,##0.00")
        End Select
        
        'agregar horas extras
        If Not DtaHrsExtras.Recordset.EOF Then
            HE = DtaHrsExtras.Recordset("canthoras")
            MontoHRSExtras = HE * MontoHora * 2

        Else
            MontoHRSExtras = 0
            HE = 0
        End If
Else
     MontoHRSExtras = 0
End If


'///////////////////////////////////////////////////////////////////
'/////////////Calculo otros Ingresos////////////////////////////////
'///////////////////////////////////////////////////////////////////

If ChkOtro.Value = 1 Then
    If TxtMontoOtrPrestacion < 0 Or TxtOtrPrestacion = "" Then
       MsgBox "Hay un error en otra prestación"
       TxtMontoOtrPrestacion.SetFocus
       Exit Sub
    Else
       TextOtro = TxtOtrPrestacion.Text
       Otro = val(TxtMontoOtrPrestacion.Text)
    End If
Else
    TextOtro = "Ninguna"
    Otro = 0
End If

'////////////////////////////////////////////////////////////////////////////
'////////Hago el Calculo Proporcional de los dias trabajados////////////////
'/////////////////////////////////////////////////////////////////////////////

If val(TxtDias.Text) > 0 Then
MontoNomPropor = (SalarioBasico / DiasMes) * val(TxtDias.Text)
Else
MontoNomPropor = 0
End If




'*************************************************************************
'*************************************************************************
'//////////////////DEDUCCIONES DEL EMPLEADO///////////////////////////////
'***************************************************************************
'***************************************************************************

'/////////////////////////////////////////////////////////////////////////
'//////////////Busco si el empleado tiene Prestamo////////////////////////
'/////////////////////////////////////////////////////////////////////////
        
If Me.ChkPrestamo.Value = 1 Then
        '///////////////Prestamos//////////////////////////
        SQlPrestamo = "SELECT MovPrestamo.NumPrestamo, MovPrestamo.NumCuota, MovPrestamo.CuotaIgual, MovPrestamo.Cancelado, MovPrestamo.NumNomina, Prestamo.CodEmpleado FROM Prestamo INNER JOIN MovPrestamo ON Prestamo.NumPrestamo = MovPrestamo.NumPrestamo WHERE MovPrestamo.Cancelado=0 AND Prestamo.CodEmpleado='" & CodEmpleado & "'"
        Dtaprestamo.RecordSource = SQlPrestamo
        Dtaprestamo.Refresh
Prestamo = 0
Do While Not Dtaprestamo.Recordset.EOF
   Prestamo = Dtaprestamo.Recordset("CuotaIgual") + Prestamo
 Me.Dtaprestamo.Recordset.MoveNext
Loop

End If
'//////////////////////////////////////////////////////////////////////////
'//////////Busco si tiene deducciones/////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////

'/////////////////////////////////////////////////////////////
'////////////Cancelo deducciones///////////////////////////////
'///////////////////////////////////////////////////////////////
If Me.Chk13mes.Value = 1 Then
 '/////////////// Deducciones //////////////////////////
        SQlDeducciones = "SELECT Deduccion.CodTipoDeduccion, DetalleDeduccion.NumDeduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.Pagado, DetalleDeduccion.NumNomina, Deduccion.CodEmpleado FROM Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion WHERE DetalleDeduccion.Pagado=0 AND Deduccion.CodEmpleado= '" & CodEmpleado & "'"
        DtaDeducciones.RecordSource = SQlDeducciones
        DtaDeducciones.Refresh

Do While Not DtaDeducciones.Recordset.EOF
    If DtaDeducciones.Recordset("CodEmpleado") = TxtCodEmpleado Then
         Deducciones = Deducciones + val(Me.DtaDeducciones.Recordset("valor"))
     
    Else
     Deducciones = 0
    End If
DtaDeducciones.Recordset.MoveNext
Loop
Else
 Deducciones = 0
End If
'/////////////////////////////////////////////////////////////////////////////////
'////////////////////CALCULO DEL INSS/////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////

MontoInss = 0
MontoInssPatronal = 0
SalarioMensual = MontoNomPropor + VACACIONES + MontoHRSExtras + Otro

         
         DtaInss.Refresh
         Do While Not DtaInss.Recordset.EOF
                If DtaTipoNomina.Recordset("Periodo") = "Semanal Viernes" Then
                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      MontoInss = DtaInss.Recordset("montolaboral1")
                      MontoInssPatronal = DtaInss.Recordset("montopatronal1")
                      Exit Do
                   End If
                   
                ElseIf DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      MontoInss = DtaInss.Recordset("montolaboral1")
                      MontoInssPatronal = DtaInss.Recordset("montopatronal1")
                      Exit Do
                   End If
                ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Viernes" Then
                
                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      If CantSabados = 4 Then
                        If DiaFin < 28 Then
                        MontoInss = (DtaInss.Recordset("montolaboral4") / 2)
                        MontoInssPatronal = (DtaInss.Recordset("montopatronal4") / 2)
                        Exit Do
                       Else
                        MontoInssMensual = DtaInss.Recordset("montolaboral4")
                        MontoInssPatronalMensual = DtaInss.Recordset("montopatronal4")
                        MontoInss = MontoInssMensual - MontoInssAnterior
                        MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
                       End If
                      Else
              '/////////Calcula para Cinco Semanas////////
                     If DiaFin < 28 Then
                        MontoInss = (DtaInss.Recordset("montolaboral5") / 2)
                        MontoInssPatronal = (DtaInss.Recordset("montopatronal5") / 2)
                        Exit Do
                     Else
                        MontoInssMensual = DtaInss.Recordset("montolaboral5")
                        MontoInssPatronalMensual = DtaInss.Recordset("montopatronal5")
                        MontoInss = MontoInssMensual - MontoInssAnterior
                        MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
                     End If
                      
                      End If
                   End If
                ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
                
                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      If CantSabados = 4 Then
               '/////////////Calcula para cuatro semanas////
                       If DiaFin < 28 Then
                        MontoInss = (DtaInss.Recordset("montolaboral4") / 2)
                        MontoInssPatronal = (DtaInss.Recordset("montopatronal4") / 2)
                        Exit Do
                       Else
                        MontoInssMensual = DtaInss.Recordset("montolaboral4")
                        MontoInssPatronalMensual = DtaInss.Recordset("montopatronal4")
                        MontoInss = MontoInssMensual - MontoInssAnterior
                        MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
                       End If
                      Else
              '/////////Calcula para Cinco Semanas////////
      
                        MontoInssMensual = DtaInss.Recordset("montolaboral5")
                        MontoInssPatronalMensual = DtaInss.Recordset("montopatronal5")
        
                  
                      End If
                   End If
                ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
                               
                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      If CantSabados = 4 Then
                       '///////Calculo para 4 Semanas///////////
   
                         MontoInssMensual = DtaInss.Recordset("montolaboral4")
                         MontoInssPatronalMensual = DtaInss.Recordset("montopatronal4")
 
                         Exit Do
             
                      Else
                      '///Calculo para 5 Semansas//////////
                       If DiaFin < 28 Then
                        MontoInss = (DtaInss.Recordset("montolaboral5") / 2)
                        MontoInssPatronal = (DtaInss.Recordset("montopatronal5") / 2)
                        
                        Exit Do
                       Else
                        MontoInssMensual = DtaInss.Recordset("montolaboral5")
                        MontoInssPatronalMensual = DtaInss.Recordset("montopatronal5")
                        MontoInss = MontoInssMensual - MontoInssAnterior
                        MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
                       End If
                      End If
                   End If
                
                ElseIf DtaTipoNomina.Recordset("Periodo") = "Mensual" Then

                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      If CantSabados = 4 Then
                        MontoInss = DtaInss.Recordset("montolaboral4")
                        MontoInssPatronal = DtaInss.Recordset("montopatronal4")
                        Exit Do
                      Else
                        MontoInss = DtaInss.Recordset("montolaboral5")
                        MontoInssPatronal = DtaInss.Recordset("montopatronal5")
                        Exit Do
                      
                      End If
                   End If
                
                ElseIf DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
                
                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      If CantSabados = 4 Then
                        MontoInss = DtaInss.Recordset("montolaboral4") * 3
                        MontoInssPatronal = DtaInss.Recordset("montopatronal4") * 3
                        Exit Do
                      Else
                        MontoInss = DtaInss.Recordset("montolaboral5") * 3
                        MontoInssPatronal = DtaInss.Recordset("montopatronal5") * 3
                        Exit Do
                      
                      End If
                   End If
                
                
                
                ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
                
                
                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
                      If CantSabados = 4 Then
                        MontoInss = DtaInss.Recordset("montolaboral4") * 6
                        MontoInssPatronal = DtaInss.Recordset("montopatronal4") * 6
                        Exit Do
                      Else
                        MontoInss = DtaInss.Recordset("montolaboral5") * 6
                        MontoInssPatronal = DtaInss.Recordset("montopatronal5") * 6
                        Exit Do
                      
                      End If
                   End If
                
                End If
                
         DtaInss.Recordset.MoveNext
         Loop
 'del if que pregunta si el empleado ers excento de INSS
 
'////////////////////////////////////////////////////////////////////////////////////
'////////////////////////CALCULO DEL IR//////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////

MontoIr = 0
MontoIRPatronal = 0

'Hago el Calcul del nuevo Techo para el Ir
MontoBrutoMensual = SalarioMensual - MontoInss

        'agregar IR laboral y patronal
        MontoIr = 0
        MontoIRPatronal = 0
        
        DtaIR.Refresh
        DtaIR.Recordset.MoveNext
        MinIR = DtaIR.Recordset("desde")
        MinIR = MinIR - 1
        MinIR = (MinIR / 12)
     '   MsgBox MinIR
        Do While Not DtaIR.Recordset.EOF
        
           'ubicar la linea
         If DtaTipoNomina.Recordset("Periodo") = "Semanal Viernes" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIr = Format(MontoIr / CantSabados / 12, "###,##0.00")
               MontoIRPatronal = MontoIr
               Exit Do
            End If
            End If
            
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIr = Format(MontoIr / CantSabados / 12, "###,##0.00")
               MontoIRPatronal = MontoIr
               Exit Do
                       
            End If
            End If
            
        ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Viernes" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
  '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
 
                MontoIrMensual = Format(MontoIr / 1 / 12, "###,##0.00")
                MontoIr = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior

            End If
            Else
               MontoIrMensual = 0
               MontoIr = MontoIrMensual - MontoIrAnterior
               MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
            End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
  '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
 
                MontoIrMensual = Format(MontoIr / 1 / 12, "###,##0.00")
                MontoIr = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior

            End If
            Else
               MontoIrMensual = 0
                MontoIr = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
            End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'///////Verfico si el la Ultima Quincena para hacer ajustes////////////

                MontoIrMensual = Format(MontoIr / 1 / 12, "###,##0.00")
                MontoIr = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior

            End If
            Else
               MontoIrMensual = 0
                MontoIr = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
            End If
         
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Mensual" Then
           If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIr = Format(MontoIr / 12, "###,##0.00")
               MontoIRPatronal = MontoIr
               Exit Do
            End If
         End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
           If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIr = Format(MontoIr / 4, "###,##0.00")
               MontoIRPatronal = MontoIr
               Exit Do
            End If
           End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
             If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIr = Format(MontoIr / 2, "###,##0.00")
               MontoIRPatronal = MontoIr
               Exit Do
            End If
            End If
         End If
        DtaIR.Recordset.MoveNext
        Loop
'fin del calculo del ir







'////////////Registro los Cambios en la tabla de Bajas/////////////////
Criterio = "CodEmpleado='" & CodEmpleado & "'"

'////////Busco si existe registro en las bajas///////////////
'e.DtaBajas.Recordset.Find Criterio
Me.DtaBajas.RecordSource = "SELECT Id  From Bajas "
Me.DtaBajas.Refresh
If DtaBajas.Recordset.EOF Then
 Id = 0
Else
 Me.DtaBajas.Recordset.MoveLast
 Id = Me.DtaBajas.Recordset("id") + 1
End If

Me.DtaBajas.RecordSource = "SELECT Id,  AnnosTrabajados,CodEmpleado, FechaBaja, MesesTrabajados, DiasTrabajados, MontoNomPropor, MontoVaca, Monto13Mes, MontoAnosTrab, MontoCargoConfianza,MontoAntiguedad, TipoBaja, MotivoBaja, Otro, MontoOtro, Prestamo, Deducciones, SalarioMensual, MontoINSS, MontoIR, FechaIniAgui, FechaFinAgui,DiasAguinaldo , DiasVacaciones, DiasMenosVaca, FechaIniVaca, FechaFinVaca, HorasExtra, Viaticos, MontoHorasExtra From Bajas WHERE     (CodEmpleado = '" & CodEmpleado & "')"
Me.DtaBajas.Refresh

'////////////////////////Sumo el total de los ingresos y Egresoso//////////////////////////
TotalIngresos = MontoNomPropor + VACACIONES + Mes13 + MontoAntiguedad + MontoHRSExtras + Otro
TotalEgresos = Deducciones + Prestamo + MontoInss + MontoIr
If Me.DtaBajas.Recordset.EOF Then
'////////Agrego un nuevo Registro////////////////
 DtaBajas.Recordset.AddNew
    DtaBajas.Recordset("id") = Id
    DtaBajas.Recordset("CodEmpleado") = CodEmpleado
    DtaBajas.Recordset("fechabaja") = Format(Now, "dd/mm/yyyy")
    DtaBajas.Recordset("DiasTrabajados") = val(TxtDias.Text)
    DtaBajas.Recordset("MontoNomPropor") = MontoNomPropor
    DtaBajas.Recordset("annostrabajados") = val(Me.TxtAnnos.Text)
    DtaBajas.Recordset("mesestrabajados") = val(Me.TxtMeses.Text)
    DtaBajas.Recordset("DiasTrabajados") = val(TxtDias.Text)
    DtaBajas.Recordset("montovaca") = Format(VACACIONES, "##,##0.00")
    DtaBajas.Recordset("monto13mes") = Format(Mes13, "##,##0.00")
    DtaBajas.Recordset("MontoAntiguedad") = MontoAntiguedad
    'DtaBajas.Recordset.montoCargoConfianza = CargoConfianza
    DtaBajas.Recordset("HorasExtra") = HE
    DtaBajas.Recordset("MontoHorasExtra") = MontoHRSExtras
    DtaBajas.Recordset("Otro") = TextOtro
    DtaBajas.Recordset("montootro") = Otro
    If Not FechaIniVaca = "0:00:00" Then
     DtaBajas.Recordset("FechaIniVaca") = CDate(FechaIniVaca)
    End If
    If Not FechaFinVaca = "0:00:00" Then
    DtaBajas.Recordset("FechaFinVaca") = CDate(FechaFinVaca)
    End If
    If Not FechaIniAgui = "0:00:00" Then
    DtaBajas.Recordset("FechaIniAgui") = FechaIniAgui
    End If
    If Not FechaFinAgui = "0:00:00" Then
     DtaBajas.Recordset("FechaFinAgui") = FechaFinAgui
    End If
    If Not Dias = 0 Then
    DtaBajas.Recordset("DiasAguinaldo") = Format(Dias, "##,##0.00")
    End If
    If Not CantRegistros = 0 Then
    DtaBajas.Recordset("DiasVacaciones") = CantRegistros * 1.25
    End If
    If Not DiasDescuento = 0 Then
    DtaBajas.Recordset("DiasMenosVaca") = DiasDescuento
    End If
    DtaBajas.Recordset("SalarioMensual") = SalMayor
    
    
    
    If OptFinContrato Then
      DtaBajas.Recordset("tipobaja") = "Fin de Contrato"
    ElseIf OptDespido Then
      DtaBajas.Recordset("tipobaja") = "Despido"
    Else
      DtaBajas.Recordset("tipobaja") = "Renuncia"
    End If
    DtaBajas.Recordset("MOTIVOBAJA") = TxtMotivo.Text
    DtaBajas.Recordset("Prestamo") = Prestamo
    DtaBajas.Recordset("Deducciones") = Deducciones
    DtaBajas.Recordset("MontoInss") = MontoInss
    DtaBajas.Recordset("MontoIR") = MontoIr
DtaBajas.Recordset.Update

Else
'/////////////Edito el registro existente.///////////////
 'Me.DtaBajas.Recordset.Edit
    DtaBajas.Recordset("CodEmpleado") = CodEmpleado
    DtaBajas.Recordset("fechabaja") = Format(Now, "dd/mm/yyyy")
    DtaBajas.Recordset("DiasTrabajados") = val(TxtDias.Text)
    DtaBajas.Recordset("MontoNomPropor") = MontoNomPropor
    DtaBajas.Recordset("annostrabajados") = val(Me.TxtAnnos.Text)
    DtaBajas.Recordset("mesestrabajados") = val(Me.TxtMeses.Text)
    DtaBajas.Recordset("DiasTrabajados") = val(TxtDias.Text)
    DtaBajas.Recordset("montovaca") = Format(VACACIONES, "##,##0.00")
    DtaBajas.Recordset("monto13mes") = Format(Mes13, "##,##0.00")
    DtaBajas.Recordset("HorasExtra") = HE
    DtaBajas.Recordset("MontoHorasExtra") = MontoHRSExtras
    DtaBajas.Recordset("MontoAntiguedad") = MontoAntiguedad
  
    'DtaBajas.Recordset.montoCargoConfianza = CargoConfianza
    DtaBajas.Recordset("Otro") = TextOtro
    DtaBajas.Recordset("montootro") = Otro
  If Not FechaIniVaca = "0:00:00" Then
     DtaBajas.Recordset("FechaIniVaca") = CDate(FechaIniVaca)
    End If
    If Not FechaFinVaca = "0:00:00" Then
    DtaBajas.Recordset("FechaFinVaca") = CDate(FechaFinVaca)
    End If
    If Not FechaIniAgui = "0:00:00" Then
    DtaBajas.Recordset("FechaIniAgui") = FechaIniAgui
    End If
    If Not FechaFinAgui = "0:00:00" Then
     DtaBajas.Recordset("FechaFinAgui") = FechaFinAgui
    End If
    If Not Dias = 0 Then
    DtaBajas.Recordset("DiasAguinaldo") = Format(Dias, "##,##0.00")
    End If
    If Not CantRegistros = 0 Then
    DtaBajas.Recordset("DiasVacaciones") = CantRegistros * 1.25
    End If
    If Not DiasDescuento = 0 Then
    DtaBajas.Recordset("DiasMenosVaca") = DiasDescuento
    End If
    DtaBajas.Recordset("SalarioMensual") = SalMayor
    
    If OptFinContrato Then
      DtaBajas.Recordset("tipobaja") = "Fin de Contrato"
    ElseIf OptDespido Then
      DtaBajas.Recordset("tipobaja") = "Despido"
    Else
      DtaBajas.Recordset("tipobaja") = "Renuncia"
    End If
    DtaBajas.Recordset("MOTIVOBAJA") = TxtMotivo.Text
    DtaBajas.Recordset("Prestamo") = Prestamo
    DtaBajas.Recordset("Deducciones") = Deducciones
    DtaBajas.Recordset("MontoInss") = MontoInss
    DtaBajas.Recordset("MontoIR") = MontoIr

 Me.DtaBajas.Recordset.Update
End If







'///////Imprimo la Liquidacion//////////////////////////////////////
k% = MsgBox("Desea Imprimir las bajas?", vbYesNo)
If k% = 6 Then

    CodEmpleado = TxtCodEmpleado.Text
    Me.DtaConsulta.RecordSource = "SELECT [Bajas].[MontoNomPropor]+[Bajas].[MontoVaca]+[Bajas].[Monto13Mes]+[Bajas].[MontoAnosTrab]+[Bajas].[MontoCargoConfianza]+[Bajas].[MontoAntiguedad]+[Bajas].[MontoOtro] AS Ingresos, [Bajas].[Prestamo]+[Bajas].[Deducciones]+[Bajas].[MontoINSS]+[Bajas].[MontoIR] AS Egresos, Bajas.CodEmpleado From Bajas Where (((Bajas.CodEmpleado) = '" & CodEmpleado & "'))"
    Me.DtaConsulta.Refresh
    
    ArepBajas.DataControl1.ConnectionString = ConexionReporte
    ArepBajas.lbltitulo.Caption = Titulo
    ArepBajas.LblSubtitulo.Caption = SubTitulo
    ArepBajas.ImgLogo.Picture = LoadPicture(RutaLogo)
    Cadena = "SELECT Empleado.CodEmpleado, [Empleado].[Nombre1]+'" & Espacio & "'+[Empleado].[Nombre2]+'" & Espacio & "'+[Empleado].[Apellido1]+'" & Espacio & "'+[Empleado].[Apellido2] AS Nombres, Bajas.FechaBaja, Bajas.AnnosTrabajados, Bajas.MesesTrabajados, Bajas.DiasTrabajados, Bajas.MontoNomPropor, Bajas.MontoVaca, Bajas.Monto13Mes, Bajas.MontoAnosTrab, Bajas.MontoCargoConfianza, Bajas.MontoAntiguedad, Bajas.MotivoBaja, Bajas.TipoBaja, Bajas.Otro, Bajas.MontoOtro, Bajas.Prestamo, Bajas.Deducciones, Bajas.MontoINSS, Bajas.MontoIR, Cargo.Cargo, Departamento.Departamento, Historico.FechaContrato, Bajas.SalarioMensual,Bajas.HorasExtra, Bajas.Viaticos " & vbLf
    Cadena = Cadena & ",Bajas.MontoHorasExtra FROM ((Departamento INNER JOIN (Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) ON Departamento.CodDepartamento = Empleado.CodDepartamento) INNER JOIN Bajas ON Empleado.CodEmpleado = Bajas.CodEmpleado) LEFT JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado Where (((Empleado.CodEmpleado) = '" & CodEmpleado & "'))"
    ArepBajas.DataControl1.Source = Cadena
'    ArepBajas.LblFechaFinVaca = FechaFinVaca
'    ArepBajas.LblFechaIniVaca = FechaIniVaca
'    ArepBajas.LblFechaIniAguinaldo = FechaIniAgui
'    ArepBajas.LblFechaFinAguinaldo = FechaFinAgui
'    ArepBajas.LblDiasAguinaldo = Format(Dias, "##,##0.00")
'    ArepBajas.LblDiasBruto = CantRegistros * 1.25
'    ArepBajas.LblDiasMenos = DiasDescuento
'    ArepBajas.LblDiasNetos = (CantRegistros * 1.25) - DiasDescuento
    
  
     NetoPagar = TotalIngresos - TotalEgresos
     ArepBajas.LblTotalEgresos.Caption = Format(TotalEgresos, "##,##0.00")
     ArepBajas.LblTotalIngresos.Caption = Format(TotalIngresos, "##,##0.00")
     ArepBajas.LblNetoPagar.Caption = Format(NetoPagar, "##,##0.00")
    

    ArepBajas.Show 1
    Me.CmdEfectuar.Enabled = True
    Me.CmdRenovar.Enabled = True
End If



End Sub

Private Sub ChkOtro_Click()

If ChkOtro.Value = 1 Then
    LblPrestacion.Visible = True
    LblMonto.Visible = True
    TxtOtrPrestacion.Visible = True
    TxtMontoOtrPrestacion.Visible = True
Else
    
    LblPrestacion.Visible = False
    LblMonto.Visible = False
    TxtOtrPrestacion.Visible = False
    TxtMontoOtrPrestacion.Visible = False
End If

End Sub

Private Sub CmdBuscarEmpleado_Click()
'Quien = "Despido"
'FrmBuscaEmpleado.Show 1
QueProducto = "CodigoEmpleado"
FrmConsulta.Show 1
Me.txtCodEmpleado1.Text = FrmConsulta.CodigoEmpleado1
End Sub

Private Sub CmdCalculos_Click()
Dim SalarioHoraAgui As Double, SalarioDiarioAgui As Double, SalarioMesAgui As Double, MontoIr As Double, MontoIRPatronal As Double
Dim TipoNomina As String, sql As String, CodigoEmpleado As String, CodTipoNomina As String, MontoBrutoMensual As Double
Dim Años As Double, PAntiguedad As Double, AñoActual As Integer, FechaInicio As Date, Prestamo As Double
Dim DiasTrabajados As Double, FechaEgreso As Date, TotalAguinaldo As Double, TotalVacaciones As Double
Dim DiasAntiguedad As Double, TotalAntiguedad As Double, TotalInss As Double, TotalIr As Double
Dim MontoHRSExtras As Double, HE As Double, Adelanto13vo As Double, AdelantoVaca As Double, TotalOtrosSalarios As Double
Dim SalarioHoraVaca As Double, SalarioDiarioVaca As Double, SalarioMesVaca As Double
Dim DiasTrabajadosVaca As Double, DiasTrabajadosAgui As Double, FechaInicioVaca As Date, FechaInicioAgui As Date
Dim DiasMes As Double, DiasProporcional As Double, MesesVacaciones As Double, DiasVacaciones As Double
Dim MesesAguinaldo As Double, DiasAguinaldo As Double, DiasVacacionesReal As Double, DiasAguinaldoReal As Double
Dim DiasAntiguedadPro As Double, SQlOtrosIngresos As String, SalarioPromedio As Double, SalarioBasico As Double
Dim AntiguedadMenorAno As Boolean, TotalDiasVaca As Double, TasaCambio As Double, DiasDescuentos As Double
Dim fPreview As New FrmPreview, SueldoActual As Double, SalarioAlto As Double, DiasTrabajadosAntigue As Double
Dim rs As New ADODB.Recordset, TasaInss As Double, Fecha1 As Date, Fecha2 As Date, Dias As Double, FechaAguinaldo As Date


CodigoEmpleado = Me.TxtCodEmpleado.Text

res = Bitacora(Now, NombreUsuario, "Liquidacion", "Se Calculo Liquidacion: " & Me.txtCodEmpleado1.Text & " " & Me.TxtNombre1.Text)

'/////////CONSULTA EL SALARIO Y TIPO DE NOMINA DEL EMPLEADO//////////////////////////

 sql = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.SueldoActualBasico, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento," & vbLf
 sql = sql & "Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.OtrosIngresos, Empleado.DescripOtrIngre," & vbLf
 sql = sql & "Empleado.ExentoIr, Empleado.PagoInssPatronal, Empleado.Activo, Empleado.Liquidado, Empleado.Ausente, Empleado.SalarioFijo," & vbLf
 sql = sql & "Empleado.SumarSubsidio , Empleado.PorcientoIncentivo, Empleado.Gravidez, TipoNomina.Periodo, TipoNomina.TasaInss" & vbLf
 sql = sql & "FROM Empleado INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina" & vbLf
 sql = sql & "WHERE     (Empleado.CodEmpleado = '" & CodigoEmpleado & "') AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
 Me.DtaConsulta.RecordSource = sql
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  TipoNomina = Me.DtaConsulta.Recordset("Periodo")
  CodTipoNomina = Me.DtaConsulta.Recordset("CodTipoNomina")
  TasaInss = Me.DtaConsulta.Recordset("TasaInss") / 100
 Else
  MsgBox "Este Empleado no Existe", vbCritical, "Sistema de Nominas"
  Exit Sub
 End If
 
'      If Not IsNull(DtaEmpleado.Recordset("SueldoActualBasico")) = True Then
'         If DtaEmpleado.Recordset("SueldoActualBasico") = True Then
'          Me.ChkSueldoActual.Value = 1
'          If Not IsNull(DtaEmpleado.Recordset("SueldoActual")) Then
'             SueldoActual = DtaEmpleado.Recordset("SueldoActual")
'          End If
'        Else
'          Me.ChkSueldoActual.Value = 0
'          SueldoActual = 0
'         End If
'        End If
 
 If Me.txtSalarioBasico.Text <> "" Then
   SalarioBasico = Me.txtSalarioBasico.Text
 Else
   SalarioBasico = 0
 End If
 
 If Me.TxtSalarioPromedio.Text <> "" Then
   SalarioPromedio = Me.TxtSalarioPromedio.Text
 Else
   SalarioPromedio = 0
 End If
 
 If Me.TxtSalarioAlto.Text <> "" Then
  SalarioAlto = Me.TxtSalarioAlto.Text
 Else
  SalarioAlto = 0
 End If
 
         MDIPrimero.DtaControles.Refresh
        If Not MDIPrimero.DtaControles.Recordset.EOF Then
         DiasMes = MDIPrimero.DtaControles.Recordset("DiasMes")
        End If
        
        If Not MDIPrimero.DtaControles.Recordset.EOF Then
         If Not IsNull(MDIPrimero.DtaControles.Recordset("AntiguedadMenor")) Then
           If MDIPrimero.DtaControles.Recordset("AntiguedadMenor") = True Then
               AntiguedadMenorAno = True
           Else
               AntiguedadMenorAno = False
           End If
         Else
           AntiguedadMenorAno = False
         End If
        End If
 
 '///FACTOR  1/12 =  0.083333
 '/////////30.4167= 365/12/////////////////
    
    '/////////////VERIFICO SI SE UTILIZA LA ANTIGUEDAD COMO BASE//////////////////////
    '/////////////PARA EL CALCULO DE LA LIQUIDACION///////////////////////////////////
  If ChkAntiguedad.Value = 1 Then
        If TxtAnnos = "" Then
         Años = 0
        Else
         Años = Int(Me.TxtAnnos.Text)
        End If
            Me.AdoAntiguedad.RecordSource = "SELECT años_acum, porcent From Antiguedad Where (años_acum = " & Años & ")"
            Me.AdoAntiguedad.Refresh
        If Not Me.AdoAntiguedad.Recordset.EOF Then
         PAntiguedad = 1 + Me.AdoAntiguedad.Recordset("porcent")
         SalarioMes = (Me.DtaConsulta.Recordset("SueldoPeriodo") * 2) * PAntiguedad
        End If
'

        
        
        If SalarioAlto > SalarioBasico Then
          SalarioMesAgui = Me.TxtSalarioAlto.Text
          SalarioHoraAgui = (SalarioMesAgui / DiasMes) / 8
          SalarioDiarioAgui = SalarioMesAgui / DiasMes
        Else
          SalarioMesAgui = Me.txtSalarioBasico.Text
          SalarioHoraAgui = (SalarioMesAgui / DiasMes) / 8
          SalarioDiarioAgui = SalarioMesAgui / DiasMes
        End If
        
        '//////////////////////ESTE CAMBIO SOLICITADO POR DOÑA MELBA, 09-03-2020  si el empleado tiene menos de 6 meses se le paga con el basico.
        
          If DateDiff("d", CDate(Me.TxtFechaContrato.Text), CDate(Me.TxtUltFechaNomina.Value)) >= 180 Then
                If val(Me.TxtSalarioPromedio.Text) > val(Me.txtSalarioBasico.Text) Then
                 SalarioMesVaca = Me.TxtSalarioPromedio.Text
                Else
                 SalarioMesVaca = Me.txtSalarioBasico.Text
                End If
          Else
          
            If Not IsNull(DtaEmpleado.Recordset("SueldoActualBasico")) = True Then
                 If SalarioPromedio > SalarioBasico Then
                  SalarioMesVaca = SalarioPromedio
                 Else
                  SalarioMesVaca = SalarioBasico
                 End If
             Else
                 SalarioMesVaca = Me.txtSalarioBasico.Text
             End If
          
                 
          End If
        SalarioHoraVaca = (SalarioMesVaca / DiasMes) / 8
        SalarioDiarioVaca = SalarioMesVaca / DiasMes
       
   Else
       
        MDIPrimero.DtaControles.Refresh
        If Not MDIPrimero.DtaControles.Recordset.EOF Then
         DiasMes = MDIPrimero.DtaControles.Recordset("DiasMes")
        End If
'        SalarioMes = Me.DtaConsulta.Recordset("SueldoPeriodo") * 2
        If SalarioAlto > SalarioBasico Then
          SalarioMesAgui = Me.TxtSalarioAlto.Text
          SalarioHoraAgui = (SalarioMesAgui / DiasMes) / 8
          SalarioDiarioAgui = SalarioMesAgui / DiasMes
        Else
          SalarioMesAgui = Me.txtSalarioBasico.Text
          SalarioHoraAgui = (SalarioMesAgui / DiasMes) / 8
          SalarioDiarioAgui = SalarioMesAgui / DiasMes
        End If
        
        
        '//////////////////////ESTE CAMBIO SOLICITADO POR DOÑA MELBA, 09-03-2020  si el empleado tiene menos de 6 meses se le paga con el basico.
        
        If DateDiff("d", CDate(Me.TxtFechaContrato.Text), CDate(Me.TxtUltFechaNomina.Value)) >= 180 Then
            If SalarioPromedio > SalarioBasico Then
             SalarioMesVaca = SalarioPromedio
            Else
             SalarioMesVaca = SalarioBasico
            End If
            
        Else
        
            If Not IsNull(DtaEmpleado.Recordset("SueldoActualBasico")) = True Then
                If SalarioPromedio > SalarioBasico Then
                 SalarioMesVaca = SalarioPromedio
                Else
                 SalarioMesVaca = SalarioBasico
                End If
            Else
                SalarioMesVaca = SalarioBasico
            End If
          
        
            
        End If
        
        

        SalarioHoraVaca = (SalarioMesVaca / DiasMes) / 8
        SalarioDiarioVaca = SalarioMesVaca / DiasMes
       End If
       
       If val(TxtDias.Text) > 0 Then
       If Not IsNull(DtaEmpleado.Recordset("SueldoActualBasico")) = True Then
         If SSueldoActual = 0 Then
           MontoNomPropor = (SalarioBasico / DiasMes) * val(TxtDias.Text)
         Else
           MontoNomPropor = (SSueldoActual / DiasMes) * val(TxtDias.Text)
         End If
       Else
           MontoNomPropor = (SalarioBasico / DiasMes) * val(TxtDias.Text)
        End If
       Else
           MontoNomPropor = 0
       End If
       
       '/////////////BUSCO EL INICIO DEL PERIODO/////////////////////
       AñoActual = Year(Me.TxtUltFechaNomina.Value)
       Me.AdoInicioAño.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (año = " & AñoActual & ") AND (Periodo = 1)"
       Me.AdoInicioAño.Refresh
       If Not Me.AdoInicioAño.Recordset.EOF Then
        FechaInicio = Me.AdoInicioAño.Recordset("Inicio")
       Else
        FechaInicio = Format(Now, "dd/mmm/yyyy")
       End If
       
       '//////////////////////////////////////////////////////////////////
       '//////////CALCULO CUANTOS DIAS TIENE TRABAJADOS////////////////////
       '////////////////////////////////////////////////////////////////////
       '//////SUMO 1 PARA AJUSTAR QUE LA RESTA DA SIEMPRE 1 DIA MENOS/////
       FechaEgreso = Me.TxtUltFechaNomina.Value
       FechaInicioAgui = Me.DTPFechaIniAgui.Value
       FechaInicioVaca = Me.DTPFechaIniVaca.Value
       DiasTrabajados = DateDiff("d", CDate(Me.TxtFechaContrato.Text), FechaEgreso) + 1 + val(Me.TxtDias.Text)
'       DiasTrabajados = CDbl(FechaEgreso) - CDbl(FechaInicio) + val(Me.TxtDias.Text) + 1
'       DiasTrabajadosVaca = CDbl(FechaEgreso) - CDbl(FechaInicioVaca) + 1
       'DiasTrabajadosVaca = CalcularDiasVaca(FechaInicioVaca, FechaEgreso)
'      If DiasTrabajados <= 30 Then
'        DiasTrabajadosAntigue = DiasTrabajados * 0.0833
'        DiasTrabajadosVaca = DiasTrabajados * 0.0833
'        DiasTrabajadosAgui = CalcularDiasAguinaldo(Me.txtCodEmpleado1.Text, FechaInicioAgui, FechaEgreso)
'      Else
'       DiasTrabajadosAntigue = CalcularDiasAntiguedad(Me.TxtFechaContrato.Text, FechaEgreso)
'       DiasTrabajadosVaca = CalcularDiasAguinaldo(Me.txtCodEmpleado1.Text, FechaInicioVaca, FechaEgreso)
'       DiasTrabajadosAgui = CalcularDiasAguinaldo(Me.txtCodEmpleado1.Text, FechaInicioAgui, FechaEgreso)
'      End If

FechaContrato = Me.TxtFechaContrato.Text
If Month(CDate(Me.TxtFechaContrato.Text)) = 2 Then
  Fecha1 = DateSerial(Year(FechaContrato), Month(FechaContrato) + 1, 1 - 1)
  If Day(CDate(Fecha1)) = 28 Then
    FechaContrato = CDate(Me.TxtFechaContrato.Text) - 2
  ElseIf Day(CDate(Fecha1)) = 29 Then
    FechaContrato = CDate(Me.TxtFechaContrato.Text) - 1
  End If
End If

      If DateDiff("d", CDate(FechaContrato), FechaEgreso) + 1 < 30 Then
         If Month(CDate(FechaContrato)) = Month(FechaEgreso) Then
              If DiasMes = 30 Then
                    Dias = Day(DateSerial(Year(FechaEgreso), Month(FechaEgreso) + 1, 0))
                    If Dias = 31 Then
                      Fecha2 = "31/" & Month(FechaEgreso) & " / " & Year(FechaEgreso)
                    Else
                      Fecha2 = FechaEgreso
                    End If
                
                    If Fecha2 = Me.TxtUltFechaNomina.Value Then
                       '--------CUANDO SE SELECCIONA UN EMPLEADO Y TIENE 2 DIAS TRABAJADOS, CON ESTE CODIGO LE AUMENTA A 30 DIAS
    '                   FechaEgreso = "30/ " & Month(FechaEgreso) & " / " & Year(FechaEgreso)
    '                   Me.TxtUltFechaNomina.Value = FechaEgreso
                    End If
              End If
               Dias = DateDiff("d", CDate(Me.TxtFechaContrato.Text), FechaEgreso) + 1
              DiasTrabajadosAntigue = 0
              DiasTrabajadosVaca = 0
             DiasTrabajadosAgui = 0
         Else
          
                '////////////////////////////VERIFICO EL MES ////////////////////////////////////////
                If Month(CDate(Me.TxtFechaContrato.Text)) = 2 Then
                   Fecha1 = DateSerial(Year(CDate(Me.TxtFechaContrato.Text)), Month(CDate(Me.TxtFechaContrato.Text)) + 1, 1 - 1) - 1
                ElseIf Month(CDate(FechaEgreso)) = 2 Then
                   Fecha1 = DateSerial(Year(CDate(FechaEgreso)), Month(CDate(FechaEgreso)) + 1, 1 - 1) - 1
                Else
                   Fecha1 = "30/ " & Month(FechaEgreso) & " / " & Year(FechaEgreso)
                End If
            
'                Fecha2 = DateSerial(Year(FechaContrato), Month(FechaContrato), 1)
          
'            Fecha1 = "30/ " & Month(CDate(Me.TxtFechaContrato.Text)) & "/" & Year(CDate(Me.TxtFechaContrato.Text))
            Fecha2 = "01/ " & Month(FechaEgreso) & " / " & Year(FechaEgreso)
''            DiasTrabajadosAntigue = (DateDiff("d", CDate(Me.TxtFechaContrato.Text), Fecha1) + 1) + (DateDiff("d", Fecha2, FechaEgreso) + 1)
''            DiasTrabajadosVaca = (DateDiff("d", CDate(Me.TxtFechaContrato.Text), Fecha1) + 1) + (DateDiff("d", Fecha2, FechaEgreso) + 1)
''            DiasTrabajadosAgui = CalcularDiasAguinaldo(Me.txtCodEmpleado1.Text, FechaInicioAgui, FechaEgreso)
            
         DiasTrabajadosAntigue = 0
         DiasTrabajadosVaca = 0
         DiasTrabajadosAgui = 0
         
         End If

         
        ElseIf DateDiff("d", CDate(FechaContrato), FechaEgreso) + 1 >= 30 And DateDiff("d", CDate(FechaContrato), FechaEgreso) + 1 <= 31 Then
                Dias = Day(DateSerial(Year(FechaEgreso), Month(FechaEgreso) + 1, 0))
                If Dias = 31 Then
                  Fecha2 = "31/" & Month(FechaEgreso) & " / " & Year(FechaEgreso)
                Else
                  Fecha2 = FechaEgreso
                End If
            
            DiasTrabajadosAntigue = 2.5
            DiasTrabajadosVaca = 2.5
            DiasTrabajadosAgui = 2.5
        Else
        
           '//////////////////////////SI ES MAYOR QUE UN MES EL CALCULO DE ANTIGUEDAD SERA NORMAL SEGUN SU FECHA DE CONTRATO ///////////////////
           '/////////////////////////  VALIDACION CON ERICK CHOI CHIN 17/06/2020 //////////////////////////////////////
           
           FechaContrato = Me.TxtFechaContrato.Text
        
            DiasTrabajadosAntigue = CalcularDiasAntiguedad(FechaContrato, FechaEgreso)
            If DateDiff("d", CDate(FechaInicioVaca), FechaEgreso) < 30 Then
              DiasTrabajadosVaca = (DateDiff("d", CDate(FechaInicioVaca), FechaEgreso) + 1) * 0.0833
            Else
            DiasTrabajadosVaca = CalcularDiasAguinaldo(Me.txtCodEmpleado1.Text, FechaInicioVaca, FechaEgreso)
            End If
            
            DiasTrabajadosAgui = CalcularDiasAguinaldo(Me.txtCodEmpleado1.Text, FechaInicioAgui, FechaEgreso)
        End If
    
       '////////////////////////////////////////////////////////////////////////
       '////////////////CALCULO AGUINALDO Y VACACIONES//////////////////////////
       '////////////////////////////////////////////////////////////////////////
       'aguinaldo= Diastrabajado*1/12*salariodiario
       
       
       If Me.Chk13mes.Value = 1 Then
         DiasProporcional = Day(FechaEgreso)
         MesesAguinaldo = DiasTrabajadosAgui / DiasMes
         
'         If (MesesAguinaldo - Int(MesesAguinaldo)) < 1 Then
'          DiasAguinaldo = Int(MesesAguinaldo) * 2.5 + (DiasProporcional * 0.083333)
'          DiasAguinaldoReal = Int(MesesAguinaldo) * 2.5 + (DiasProporcional * 0.083333)
'          TotalAguinaldo = DiasAguinaldo * SalarioDiarioAgui
'         End If


        FechaAguinaldo = Me.DTPFechaIniAgui.Value
        If Month(CDate(FechaContrato)) = 2 Then
          Fecha1 = DateSerial(Year(FechaContrato), Month(FechaContrato) + 1, 1 - 1)
          If Day(CDate(Fecha1)) = 28 Then
            FechaAguinaldo = CDate(Me.TxtFechaContrato.Text) - 2
          ElseIf Day(CDate(Fecha1)) = 29 Then
            FechaAguinaldo = CDate(Me.TxtFechaContrato.Text) - 1
          End If
        End If

         
         
         If DateDiff("d", CDate(FechaAguinaldo), FechaEgreso) < 30 Then
                  Dias = DateDiff("d", CDate(FechaAguinaldo), FechaEgreso) + 1
                  If Dias < 0 Then
                    Dias = 0
                  End If
                  
                  DiasAguinaldo = Dias * 0.08333
                  DiasAguinaldoReal = Dias * 0.08333
                  
                  
         Else
                  DiasAguinaldo = DiasTrabajadosAgui '* 0.083333333
                  DiasAguinaldoReal = DiasTrabajadosAgui '* 0.083333333
         End If
         
         

             TotalAguinaldo = Format(Format(DiasAguinaldo, "##,##0.00") * Format(SalarioDiarioAgui, "##,##0.0000"), "##,##0.00") '* 0.083333333
    
             
             '////////////////////////////////////////////////////////////////////
             '/////////////Busco el Adelanto de 13vo mes Registrados//////////////
             '////////////////////////////////////////////////////////////////////////
    
            Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo WHERE (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Adelanto13vo.FechaAdelanto) Between  '" & Format(FechaInicio, "yyyymmdd") & "' And '" & Format(FechaEgreso, "yyyymmdd") & "') AND ((Adelanto13vo.TipoAdelanto)='13vo Mes'))"
            Me.DtaAdelanto.Refresh
            Adelanto13vo = 0
            Do While Not DtaAdelanto.Recordset.EOF
             Adelanto13vo = Adelanto13vo + DtaAdelanto.Recordset("MontoAdelanto")
             DtaAdelanto.Recordset.MoveNext
            Loop
             
             TotalAguinaldo = TotalAguinaldo - Adelanto13vo
       Else
            TotalAguinaldo = 0
       End If
       
       
       
       'vacaciones
       If Me.ChkVaca.Value = 1 Then

        DiasProporcional = Day(FechaEgreso)
        MesesVacaciones = DiasTrabajadosVaca
        If Me.TxtDescuentoDias.Text <> "" Then
          DiasDescuentos = Me.TxtDescuentoDias.Text
        End If
        DiasDescuentos = Format(DiasDescuentos, "##,##0.00")
        DiasTrabajadosVaca = Format(DiasTrabajadosVaca, "##,##0.00")
        

        
        ''''''''''  DiasVacaciones = (DiasTrabajadosVaca * 0.083333333) - CDbl(val(Me.TxtDescuentoDias.Text))
          DiasVacaciones = DiasTrabajadosVaca - DiasDescuentos
          '''''''''DiasVacacionesReal = (DiasTrabajadosVaca * 0.083333333)
          DiasVacacionesReal = (DiasTrabajadosVaca)
          TotalVacaciones = DiasVacaciones * SalarioDiarioVaca
       
       '//////////////////////////////////////////////////////////////////////
       '/////////////Busco el Adelanto de Vacaciones Registrados//////////////
       '////////////////////////////////////////////////////////////////////////
         
        Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo WHERE (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Adelanto13vo.FechaAdelanto) Between  '" & Format(FechaInicio, "yyyymmdd") & "' And '" & Format(FechaEgreso, "yyyymmdd") & "') AND ((Adelanto13vo.TipoAdelanto)='Vacaciones'))"
        Me.DtaAdelanto.Refresh
        AdelantoVaca = 0
        
        Do While Not DtaAdelanto.Recordset.EOF
         AdelantoVaca = AdelantoVaca + DtaAdelanto.Recordset("MontoAdelanto")
         DtaAdelanto.Recordset.MoveNext
        Loop
       
        TotalVacaciones = TotalVacaciones - AdelantoVaca
       Else
        TotalVacaciones = 0
       End If
       

       
     If Me.ChkExtra.Value = 1 Then
       '////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////// HORAS EXTRAS /////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////
        'calculo el monto por hora dependiendo del tipo de nomina
        SalarioHora = SalarioHoraVaca
        
        SqlHrsExtras = "SELECT HorasExtras.CodEmpleado, HorasExtras.NumNomina, HorasExtras.CantHoras, HorasExtras.Pagada From HorasExtras WHERE  HorasExtras.Pagada=0 AND HorasExtras.CodEmpleado='" & CodEmpleado & "'"
        Me.DtaHrsExtras.RecordSource = SqlHrsExtras
        DtaHrsExtras.Refresh
       
            If Not DtaHrsExtras.Recordset.EOF Then
             HE = DtaHrsExtras.Recordset("canthoras")
             MontoHRSExtras = HE * SalarioHora * 2
            Else
             MontoHRSExtras = 0
             HE = 0
            End If
      End If
         
    TotalAntiguedad = 0
       '////////////////////////////////////////////////////////////////////////////////
       '//////////////CALCULO ANTIGUEDAD/////////////////////////////////////////////////
       '/////////////////////////////////////////////////////////////////////////////////
       Dim Dia As Double, Mes As Double, Meses As Double, DiaMes As Double
       MDIPrimero.DtaControles.Refresh
       DiasAntiguedad = Me.TxtDiasTrabajados.Text
       DiasMes = MDIPrimero.DtaControles.Recordset("DiasMes")
       Meses = DiasAntiguedad / DiasMes
       Mes = Int(Meses)
       Dia = Format((Meses - Mes) * DiasMes, "####0")
       
       
       
       '/////////////VERIFICO SI SE UTILIZA LA ANTIGUEDAD COMO BASE//////////////////////
       '/////////////PARA EL CALCULO DE LA LIQUIDACION///////////////////////////////////
       
'        If DiasAntiguedad >= 365 Then
'           TotalAntiguedad = DiasAntiguedad * 0.083333 * SalarioDiarioVaca
'        Else
'            TotalAntiguedad = 0
'        End If

      If (Meses / 12) >= 3 Then    '////// EL AÑO TIENE 360 POR QUE SON 30DIAS * 12 MESES
          If Int(Meses / 12) >= 6 Then
            TotalAntiguedad = SalarioDiarioVaca * DiasMes * 5
          Else
            '////////////////////////////CALCULO LOS DIAS PARA 3 AÑOS //////////////////
            If Meses >= 36 Then
              DiasTrabajadosAntigue = 90 + (Mes - 36) * 1.67 + (Dia / DiasMes) * 1.67
            Else
              DiasTrabajadosAntigue = 90  ' 36 meses multiplicado por 2.5 dias
            End If
          
            TotalAntiguedad = Format(Format(SalarioDiarioVaca, "##,##0.0000") * Format(DiasTrabajadosAntigue, "##,##0.00"), "##,##0.00")
          End If
      ElseIf Int(Meses / 12) >= 1 Then
          TotalAntiguedad = Format(Format(SalarioDiarioVaca, "##,##0.0000") * Format(DiasTrabajadosAntigue, "##,##0.00"), "##,##0.00")
       
       ElseIf AntiguedadMenorAno = True Then
        If DiasAntiguedad > DiasMes Then
         DiasAntiguedad = DiasTrabajadosAntigue
         TotalAntiguedad = SalarioDiarioVaca * DiasTrabajadosAntigue
        Else
         TotalAntiguedad = 0
        End If
       Else
          TotalAntiguedad = 0
      End If
       
      '//////////////////////////////////////////////////////////////////////////
      '/////////CALCULO LOS OTROS INGRESOS//////////////////////////////////////
      If Me.ChkOtro.Value = 1 Then
        TotalOtrosSalarios = val(Me.TxtMontoOtrPrestacion.Text)
      Else
        TotalOtrosSalarios = 0
      
      End If
      
       
      '////////////////////////////////////////////////////////////////////////////////////
      '//////////CALCULO LOS OTROS INGRESOS DE LA PLANILLA/////////////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////
      
      If Me.ChkOtroPlanilla.Value = 1 Then
        SQlOtrosIngresos = "SELECT  CodEmpleado, SUM(OtrosIngresos) AS OtrosIngresos From DetalleNomina GROUP BY CodEmpleado Having (CodEmpleado = " & CodEmpleado & ")"
        Me.DtaConsulta.RecordSource = SQlOtrosIngresos
        Me.DtaConsulta.Refresh
        If Not Me.DtaConsulta.Recordset.EOF Then
          TotalOtrosSalarios = TotalOtrosSalarios + val(Me.DtaConsulta.Recordset("OtrosIngresos"))
             
        End If
      End If
      
      
'*************************************************************************
'*************************************************************************
'//////////////////DEDUCCIONES DEL EMPLEADO///////////////////////////////
'***************************************************************************
'***************************************************************************

        '/////////////////////////////////////////////////////////////////////////
        '//////////////Busco si el empleado tiene Prestamo////////////////////////
        '/////////////////////////////////////////////////////////////////////////
        
        TasaCambio = BuscaTasaCambio(Me.TxtUltFechaNomina.Value)
        If Me.ChkPrestamo.Value = 1 Then
        '///////////////Prestamos//////////////////////////
        SQlPrestamo = "SELECT MovPrestamo.NumPrestamo, MovPrestamo.NumCuota, MovPrestamo.CuotaIgual, MovPrestamo.Cancelado, MovPrestamo.NumNomina, Prestamo.CodEmpleado, Moneda FROM Prestamo INNER JOIN MovPrestamo ON Prestamo.NumPrestamo = MovPrestamo.NumPrestamo WHERE MovPrestamo.Cancelado=0 AND Prestamo.CodEmpleado='" & CodEmpleado & "'"
        Dtaprestamo.RecordSource = SQlPrestamo
        Dtaprestamo.Refresh
        Prestamo = 0
        Do While Not Dtaprestamo.Recordset.EOF
        
         If Me.Dtaprestamo.Recordset("Moneda") = "US" Then
           Prestamo = (Dtaprestamo.Recordset("CuotaIgual") * TasaCambio) + Prestamo
         Else
           Prestamo = Dtaprestamo.Recordset("CuotaIgual") + Prestamo
         End If
         
         
        
        Me.Dtaprestamo.Recordset.MoveNext
        Loop
        
        
              
        

        End If
      
        If Me.ChkDeducciones.Value = 1 Then
        '/////////////////////////////////////////////////////////////
        '////////////Cancelo deducciones///////////////////////////////
        '///////////////////////////////////////////////////////////////
        '/////////////// Deducciones //////////////////////////
            SQlDeducciones = "SELECT Deduccion.CodTipoDeduccion, DetalleDeduccion.NumDeduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.Pagado, DetalleDeduccion.NumNomina, Deduccion.CodEmpleado FROM Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion WHERE DetalleDeduccion.Pagado=0 AND Deduccion.CodEmpleado= '" & CodEmpleado & "'"
            DtaDeducciones.RecordSource = SQlDeducciones
            DtaDeducciones.Refresh
            Do While Not DtaDeducciones.Recordset.EOF
                If DtaDeducciones.Recordset("CodEmpleado") = TxtCodEmpleado Then
                    Deducciones = Deducciones + DtaDeducciones.Recordset("valor")
                Else
                    Deducciones = 0
                End If
            DtaDeducciones.Recordset.MoveNext
            Loop
        End If
      
      
      
             
       '//////////////CALCULO DEL INSS///////////////////////////////////////
       TotalInss = (TotalVacaciones + MontoHRSExtras + TotalOtrosSalarios + MontoNomPropor) * TasaInss
       
       '////////////////////////////////////////////////////////////////////////////////////
       '////////////////////////CALCULO DEL IR//////////////////////////////////////////////
       '////////////////////////////////////////////////////////////////////////////////////
       Dim IrUltimaSemana As Boolean

        MontoIr = 0
        MontoIRPatronal = 0
        
'        CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
        IrUltimaSemana = DtaTipoNomina.Recordset("IrUltimaSemana")

        'Hago el Calcul del nuevo Techo para el Ir
         MontoBrutoMensual = (MontoHRSExtras + TotalOtrosSalarios + MontoNomPropor) - (MontoHRSExtras + TotalOtrosSalarios + MontoNomPropor) * TasaInss

        'agregar IR laboral y patronal
        MontoIr = 0
        MontoIRPatronal = 0
        
        
        MontoIr = CalcularMontoIrBajas(CodEmpleado, IrUltimaSemana, CodTipoNomina, Me.TxtUltFechaNomina.Value, MontoBrutoMensual)

        
'        DtaIR.Refresh
'        DtaIR.Recordset.MoveNext
'        MinIR = DtaIR.Recordset("desde")
'        MinIR = MinIR - 1
'        MinIR = (MinIR / 12)
'     '   MsgBox MinIR
'        Do While Not DtaIR.Recordset.EOF
'
'           'ubicar la linea
'         If DtaTipoNomina.Recordset("Periodo") = "Semanal Viernes" Then
'            If (MontoBrutoMensual) >= MinIR Then
'            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'               MontoIr = Format(MontoIr / CantSabados / 12, "###,##0.00")
'               MontoIRPatronal = MontoIr
'               Exit Do
'            End If
'            End If
'
'         ElseIf DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
'            If (MontoBrutoMensual) >= MinIR Then
'            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'               MontoIr = Format(MontoIr / CantSabados / 12, "###,##0.00")
'               MontoIRPatronal = MontoIr
'               Exit Do
'
'            End If
'            End If
'
'        ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Viernes" Then
'            If (MontoBrutoMensual) >= MinIR Then
'            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'  '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
'
'                MontoIrMensual = Format(MontoIr / 1 / 12, "###,##0.00")
'                MontoIr = MontoIrMensual - MontoIrAnterior
'                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
'
'            End If
'            Else
'               MontoIrMensual = 0
'               MontoIr = MontoIrMensual - MontoIrAnterior
'               MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
'            End If
'         ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
'            If (MontoBrutoMensual) >= MinIR Then
'            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'  '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
'
'                MontoIrMensual = Format(MontoIr / 1 / 12, "###,##0.00")
'                MontoIr = MontoIrMensual - MontoIrAnterior
'                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
'
'            End If
'            Else
'               MontoIrMensual = 0
'                MontoIr = MontoIrMensual - MontoIrAnterior
'                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
'            End If
'         ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
'            If (MontoBrutoMensual) >= MinIR Then
'            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
''///////Verfico si el la Ultima Quincena para hacer ajustes////////////
'
'                MontoIrMensual = Format(MontoIr / 1 / 12, "###,##0.00")
'                MontoIr = MontoIrMensual - MontoIrAnterior
'                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
'
'            End If
'            Else
'               MontoIrMensual = 0
'                MontoIr = MontoIrMensual - MontoIrAnterior
'                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
'            End If
'
'         ElseIf DtaTipoNomina.Recordset("Periodo") = "Mensual" Then
'           If (MontoBrutoMensual) >= MinIR Then
'            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'               MontoIr = Format(MontoIr / 12, "###,##0.00")
'               MontoIRPatronal = MontoIr
'               Exit Do
'            End If
'         End If
'         ElseIf DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
'           If (MontoBrutoMensual) >= MinIR Then
'            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'               MontoIr = Format(MontoIr / 4, "###,##0.00")
'               MontoIRPatronal = MontoIr
'               Exit Do
'            End If
'           End If
'         ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
'             If (MontoBrutoMensual) >= MinIR Then
'            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'               MontoIr = Format(MontoIr / 2, "###,##0.00")
'               MontoIRPatronal = MontoIr
'               Exit Do
'            End If
'            End If
'         End If
'        DtaIR.Recordset.MoveNext
'        Loop
'fin del calculo del ir

        '////////////Registro los Cambios en la tabla de Bajas/////////////////
        Criterio = "CodEmpleado='" & CodEmpleado & "'"
        
        '////////Busco si existe registro en las bajas///////////////
        'e.DtaBajas.Recordset.Find Criterio
        Me.DtaBajas.RecordSource = "SELECT Id  From Bajas "
        Me.DtaBajas.Refresh
        If DtaBajas.Recordset.EOF Then
            Id = 0
        Else
            Me.DtaBajas.Recordset.MoveLast
            Id = Me.DtaBajas.Recordset("id") + 1
        End If

        Me.DtaBajas.RecordSource = "SELECT Id,  AnnosTrabajados,CodEmpleado, FechaBaja, MesesTrabajados, DiasTrabajados, MontoNomPropor, MontoVaca, Monto13Mes, MontoAnosTrab, MontoCargoConfianza,MontoAntiguedad, TipoBaja, MotivoBaja, Otro, MontoOtro, Prestamo, Deducciones, SalarioMensual, MontoINSS, MontoIR, FechaIniAgui, FechaFinAgui,DiasAguinaldo , DiasVacaciones, DiasMenosVaca, FechaIniVaca, FechaFinVaca, HorasExtra, Viaticos, MontoHorasExtra,  FechaEgreso, FechaHistorial, SueldoActualBasicoLiquida, PrestamoOpt, DeduccionesOpt, AguinaldoOpt, AntiguedadOpt, ViaticosOpt, VacacionesOpt, HorasExtraOpt, OtrosIngresosOpt , OtrosIngresosPlanillaOpt, Prestacion, MontoPagarPrestacion, Calculada, Procesada  From Bajas WHERE     (CodEmpleado = '" & CodEmpleado & "')"
        Me.DtaBajas.Refresh

        '////////////////////////Sumo el total de los ingresos y Egresoso//////////////////////////
        TotalIngresos = MontoNomPropor + VACACIONES + Mes13 + MontoAntiguedad + MontoHRSExtras + TotalOtrosSalarios
        TotalEgresos = Deducciones + Prestamo + TotalInss + MontoIr
     If Me.DtaBajas.Recordset.EOF Then
        '////////Agrego un nuevo Registro////////////////
        DtaBajas.Recordset.AddNew
        DiasTrabajados = Me.TxtDiasTrabajados.Text
        DtaBajas.Recordset("id") = Id
        DtaBajas.Recordset("CodEmpleado") = CodEmpleado
        DtaBajas.Recordset("fechabaja") = Me.TxtUltFechaNomina.Value
        DtaBajas.Recordset("MontoNomPropor") = MontoNomPropor
        DtaBajas.Recordset("annostrabajados") = val(Me.TxtAnnos.Text)
        DtaBajas.Recordset("mesestrabajados") = val(Me.TxtMeses.Text)
        DtaBajas.Recordset("DiasTrabajados") = Format(DiasTrabajadosAntigue, "##,##0.00")
        DtaBajas.Recordset("montovaca") = Format(TotalVacaciones, "##,##0.00")
        DtaBajas.Recordset("monto13mes") = Format(TotalAguinaldo, "##,##0.00")
        DtaBajas.Recordset("MontoAntiguedad") = Format(TotalAntiguedad, "##,##0.00")
        DtaBajas.Recordset("HorasExtra") = HE
        DtaBajas.Recordset("MontoHorasExtra") = MontoHRSExtras
        DtaBajas.Recordset("Otro") = TextOtro
        DtaBajas.Recordset("montootro") = TotalOtrosSalarios
        
'             OtrosIngresosOpt , OtrosIngresosPlanillaOpt, Prestacion, MontoPagarPrestacion

            
            DtaBajas.Recordset("FechaIniVaca") = CDate(FechaInicioVaca)
            DtaBajas.Recordset("FechaFinVaca") = CDate(Me.TxtUltFechaNomina.Value)
            DtaBajas.Recordset("FechaIniAgui") = FechaInicioAgui
            DtaBajas.Recordset("FechaFinAgui") = Me.TxtUltFechaNomina.Value
            DtaBajas.Recordset("DiasAguinaldo") = Format(DiasAguinaldo, "##,##0.00")
            DtaBajas.Recordset("DiasVacaciones") = Format(DiasVacaciones, "##,##0.00")
            DtaBajas.Recordset("DiasMenosVaca") = val(Me.TxtDescuentoDias.Text)
            DtaBajas.Recordset("SalarioMensual") = SalarioMesVaca    'SalarioMesAgui   Modificado Doña Melba 13-05-2020
            DtaBajas.Recordset("FechaEgreso") = Me.TxtUltFechaNomina.Value
            DtaBajas.Recordset("FechaHistorial") = Me.TxtFechaHistorial.Value
            
            If Me.ChkSueldoActual.Value = xtpChecked Then
               DtaBajas.Recordset("SueldoActualBasicoLiquida") = 1
            Else
               DtaBajas.Recordset("SueldoActualBasicoLiquida") = 0
            End If
            
            
            If Me.ChkSueldoActual.Value = xtpChecked Then
               DtaBajas.Recordset("PrestamoOpt") = 1
            Else
               DtaBajas.Recordset("PrestamoOpt") = 0
            End If
            
            If Me.ChkPrestamo.Value = xtpChecked Then
               DtaBajas.Recordset("PrestamoOpt") = 1
            Else
               DtaBajas.Recordset("PrestamoOpt") = 0
            End If
            
            If Me.ChkDeducciones.Value = xtpChecked Then
               DtaBajas.Recordset("DeduccionesOpt") = 1
            Else
               DtaBajas.Recordset("DeduccionesOpt") = 0
            End If
    
    
            If Me.ChkDeducciones.Value = xtpChecked Then
               DtaBajas.Recordset("DeduccionesOpt") = 1
            Else
               DtaBajas.Recordset("DeduccionesOpt") = 0
            End If
    
            If Me.Chk13mes.Value = xtpChecked Then
               DtaBajas.Recordset("AguinaldoOpt") = 1
            Else
               DtaBajas.Recordset("AguinaldoOpt") = 0
            End If
            
            If Me.ChkAntiguedad.Value = xtpChecked Then
               DtaBajas.Recordset("AntiguedadOpt") = 1
            Else
               DtaBajas.Recordset("AntiguedadOpt") = 0
            End If
            
            If Me.ChkCargo.Value = xtpChecked Then
               DtaBajas.Recordset("ViaticosOpt") = 1
            Else
               DtaBajas.Recordset("ViaticosOpt") = 0
            End If
            
            If Me.ChkVaca.Value = xtpChecked Then
               DtaBajas.Recordset("VacacionesOpt") = 1
            Else
               DtaBajas.Recordset("VacacionesOpt") = 0
            End If
           
            
            If Me.ChkExtra.Value = xtpChecked Then
               DtaBajas.Recordset("HorasExtraOpt") = 1
            Else
               DtaBajas.Recordset("HorasExtraOpt") = 0
            End If
            
            If Me.ChkExtra.Value = xtpChecked Then
               DtaBajas.Recordset("HorasExtraOpt") = 1
            Else
               DtaBajas.Recordset("HorasExtraOpt") = 0
            End If
    
        If OptFinContrato Then
            DtaBajas.Recordset("tipobaja") = "Fin de Contrato"
        ElseIf OptDespido Then
            DtaBajas.Recordset("tipobaja") = "Despido"
        Else
            DtaBajas.Recordset("tipobaja") = "Renuncia"
        End If
        DtaBajas.Recordset("MOTIVOBAJA") = TxtMotivo.Text
        DtaBajas.Recordset("Prestamo") = Prestamo
        DtaBajas.Recordset("Deducciones") = Deducciones
        DtaBajas.Recordset("MontoInss") = TotalInss
        DtaBajas.Recordset("MontoIR") = MontoIr
        DtaBajas.Recordset("Calculada") = 1
        DtaBajas.Recordset("Procesada") = 0
        
        DtaBajas.Recordset.Update

    Else
        '/////////////Edito el registro existente.///////////////
 'Me.DtaBajas.Recordset.Edit
        DiasTrabajados = Me.TxtDiasTrabajados.Text
        DtaBajas.Recordset("MontoNomPropor") = MontoNomPropor
        DtaBajas.Recordset("annostrabajados") = val(Me.TxtAnnos.Text)
        DtaBajas.Recordset("mesestrabajados") = val(Me.TxtMeses.Text)
        DtaBajas.Recordset("DiasTrabajados") = Format(DiasTrabajadosAntigue, "##,##0.00")
        DtaBajas.Recordset("montovaca") = Format(TotalVacaciones, "##,##0.00")
        DtaBajas.Recordset("monto13mes") = Format(TotalAguinaldo, "##,##0.00")
        DtaBajas.Recordset("MontoAntiguedad") = Format(TotalAntiguedad, "##,##0.00")
        DtaBajas.Recordset("fechabaja") = Me.TxtUltFechaNomina.Value
        
        DtaBajas.Recordset("HorasExtra") = HE
        DtaBajas.Recordset("MontoHorasExtra") = MontoHRSExtras
        
         If Me.ChkExtra.Value = xtpChecked Then
               DtaBajas.Recordset("HorasExtraOpt") = 1
         Else
               DtaBajas.Recordset("HorasExtraOpt") = 0
         End If
         
    'DtaBajas.Recordset.montoCargoConfianza = CargoConfianza
    DtaBajas.Recordset("Otro") = TextOtro
    DtaBajas.Recordset("montootro") = TotalOtrosSalarios

     DtaBajas.Recordset("FechaIniVaca") = CDate(FechaInicioVaca)

    DtaBajas.Recordset("FechaFinVaca") = CDate(Me.TxtUltFechaNomina.Value)

    DtaBajas.Recordset("FechaIniAgui") = FechaInicioAgui

     DtaBajas.Recordset("FechaFinAgui") = Me.TxtUltFechaNomina.Value

    DtaBajas.Recordset("DiasAguinaldo") = Format(DiasAguinaldo, "####0.00")

    DtaBajas.Recordset("DiasVacaciones") = Format(DiasVacaciones, "####0.00")


    DtaBajas.Recordset("DiasMenosVaca") = val(Me.TxtDescuentoDias.Text)

    DtaBajas.Recordset("SalarioMensual") = SalarioMesVaca  'SalarioMesAgui
    
    If OptFinContrato Then
      DtaBajas.Recordset("tipobaja") = "Fin"
    ElseIf OptDespido Then
      DtaBajas.Recordset("tipobaja") = "Despido"
    Else
      DtaBajas.Recordset("tipobaja") = "Renuncia"
    End If
    DtaBajas.Recordset("MOTIVOBAJA") = TxtMotivo.Text
    DtaBajas.Recordset("Prestamo") = Prestamo
    DtaBajas.Recordset("Deducciones") = Deducciones
    DtaBajas.Recordset("MontoInss") = Format(TotalInss, "##,##0.00")
    DtaBajas.Recordset("MontoIR") = MontoIr
    DtaBajas.Recordset("Calculada") = 1
    DtaBajas.Recordset("Procesada") = 0

 Me.DtaBajas.Recordset.Update
End If







'///////Imprimo la Liquidacion//////////////////////////////////////
k% = MsgBox("Desea Imprimir las bajas?", vbYesNo)
If k% = 6 Then
    
    CodEmpleado = TxtCodEmpleado.Text
    Me.DtaConsulta.RecordSource = "SELECT [Bajas].[MontoNomPropor]+[Bajas].[MontoVaca]+[Bajas].[Monto13Mes]+[Bajas].[MontoAnosTrab]+[Bajas].[MontoCargoConfianza]+[Bajas].[MontoAntiguedad]+[Bajas].[MontoOtro] AS Ingresos, [Bajas].[Prestamo]+[Bajas].[Deducciones]+[Bajas].[MontoINSS]+[Bajas].[MontoIR] AS Egresos, Bajas.CodEmpleado From Bajas Where (((Bajas.CodEmpleado) = '" & CodEmpleado & "'))"
    Me.DtaConsulta.Refresh
    
    ARBaja.DataControl1.ConnectionString = ConexionReporte
    ARBaja.lbltitulo.Caption = Titulo
    ARBaja.LblSubtitulo.Caption = SubTitulo
   
    If Dir(RutaLogo) <> "" Then
    ARBaja.ImgLogo.Picture = LoadPicture(RutaLogo)
    End If

    ARBaja.LblSalarioAlto.Caption = Me.TxtSalarioAlto.Text
    Cadena = "SELECT     Empleado.CodEmpleado1,Empleado.CodEmpleado, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Bajas.FechaBaja, " & vbLf
    Cadena = Cadena & "                  Bajas.AnnosTrabajados, Bajas.MesesTrabajados, Bajas.DiasTrabajados, Bajas.MontoNomPropor, Bajas.MontoVaca, Bajas.Monto13Mes," & vbLf
    Cadena = Cadena & "                  Bajas.MontoAnosTrab, Bajas.MontoCargoConfianza, Bajas.MontoAntiguedad, Bajas.MotivoBaja, Bajas.TipoBaja, Bajas.Otro, Bajas.MontoOtro," & vbLf
    Cadena = Cadena & "                  Bajas.Prestamo, Bajas.Deducciones, Bajas.MontoINSS, Bajas.MontoIR, Cargo.Cargo, Departamento.Departamento, Historico.FechaContrato," & vbLf
    Cadena = Cadena & "                  Bajas.SalarioMensual, Bajas.HorasExtra, Bajas.Viaticos, Bajas.MontoHorasExtra, Bajas.FechaIniAgui, Bajas.FechaFinAgui, Bajas.DiasAguinaldo," & vbLf
    Cadena = Cadena & "                  Bajas.DiasVacaciones , Bajas.DiasMenosVaca, Bajas.FechaIniVaca, Bajas.FechaFinVaca" & vbLf
    Cadena = Cadena & "FROM         Departamento INNER JOIN" & vbLf
    Cadena = Cadena & "                  Cargo INNER JOIN" & vbLf
    Cadena = Cadena & "                  Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento INNER JOIN" & vbLf
    Cadena = Cadena & "                  Bajas ON Empleado.CodEmpleado = Bajas.CodEmpleado LEFT OUTER JOIN" & vbLf
    Cadena = Cadena & "                  Historico ON Empleado.CodEmpleado = Historico.Codempleado" & vbLf
    Cadena = Cadena & "Where (Empleado.CodEmpleado = '" & CodEmpleado & "')"
    
    ARBaja.DataControl1.Source = Cadena
    TotalVacaciones = Format(TotalVacaciones, "####0.00")
    MontoNomPropor = Format(MontoNomPropor, "####0.00")
    TotalAguinaldo = Format(TotalAguinaldo, " ####0.00")
    MontoHRSExtras = Format(MontoHRSExtras, "####0.00")
    
'    ARBaja.LblDiasNeto = Format(DiasVacaciones, "##,##0.00")
    TotalIngresos = (TotalVacaciones + MontoNomPropor + TotalAguinaldo + TotalAntiguedad + MontoHRSExtras + TotalOtrosSalarios)
    TotalEgresos = TotalInss + MontoIr + Prestamo + Deducciones
     NetoPagar = TotalIngresos - TotalEgresos
     ARBaja.LblTotalEgresos.Caption = Format(TotalEgresos, "##,##0.00")
     ARBaja.LblTotalIngresos.Caption = Format(TotalIngresos, "##,##0.00")
     ARBaja.LblNetoPagar.Caption = Format(NetoPagar, "##,##0.00")
     ARBaja.LblSalarioBasico.Caption = Me.txtSalarioBasico.Text
     ARBaja.LblDiasAgui.Caption = Format(DiasAguinaldo, "##,##0.00")
     ARBaja.LblDiasVaca.Caption = Format(DiasVacacionesReal, "##,##0.00")
     ARBaja.txtSalPorDia.Text = Format(CDbl(Me.TxtSalarioAlto.Text) / 30, "##,##0.00")
     
    ''''''' ARBaja.FldDiasVaca.Text = Format(DiasTrabajadosVaca * 0.083333333, "##,##0.00")
    
    ARBaja.FldDiasVaca.Text = Format(DiasTrabajadosVaca, "##,##0.00")
    ARBaja.Show 1

'     Set rpt = New ARBaja
'     rpt.DataControl1.ConnectionString = ConexionReporte
'     rpt.DataControl1.Source = Cadena
'     fPreview.RunReport rpt
'     fPreview.Show 1
     
    Me.CmdEfectuar.Enabled = True
    Me.CmdRenovar.Enabled = True
End If

   
       
  

End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdDetalle_Click()
On Error GoTo TipoErrs
Dim SqlSalarios As String, FechaEgreso As Date, FechaContrato As Date, FechaHistorico As Date, FechaBusqueda As Date
Dim CodEmpleado As String, NumeroEmpleado As Integer, i As Integer
Dim rpt As Object
Dim fPreview As New FrmPreview

    CodEmpleado = TxtCodEmpleado.Text
  

  
  '///////////Busco la Fecha para la Busqueda////////////////////////////
  
FechaEgreso = Me.TxtFechaHistorial.Value
'FechaContrato = Me.TxtFechaContrato.Text


NumeroEmpleado = Me.TxtCodEmpleado.Text

SqlSalarios = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo,SUM(DetalleNomina.Incentivos) AS Incentivos, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.Comisiones) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes AS MES, Nomina.Ano AS AÑO,SUM(DetalleNomina.Comisiones) As Comisiones " & _
              "FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano  " & _
              "HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") AND (MAX(Nomina.FechaNomina) <= CONVERT(DATETIME, '" & Format(FechaEgreso, "yyyy/mm/dd") & "', 102))"

'SQLSalarios = "SELECT DISTINCT" & vbLf
'SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
'SQLSalarios = SQLSalarios & "AS Destajo, SUM(dbo.DetalleNomina.Incentivos) AS Incentivos, SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo)" & vbLf
'SQLSalarios = SQLSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes AS MES," & vbLf
'SQLSalarios = SQLSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
'SQLSalarios = SQLSalarios & "FROM   dbo.DetalleNomina INNER JOIN" & vbLf
'SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
'SQLSalarios = SQLSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
'SQLSalarios = SQLSalarios & "Having (dbo.DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") And (Sum(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0) AND (MAX(Nomina.FechaNomina) <= CONVERT(DATETIME, '" & Format(FechaEgreso, "yyyy/mm/dd") & "', 102))"
Me.DtaConsulta.RecordSource = SqlSalarios
Me.DtaConsulta.Refresh


Me.DtaConsulta.Recordset.MoveLast
i = 0
Do While Not Me.DtaConsulta.Recordset.BOF
  If i = 1 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")

  ElseIf i = 5 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    Exit Do
  ElseIf i = 0 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    FechaHistorico = Me.DtaConsulta.Recordset("FechaFin")
  Else
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
  End If
  i = i + 1

  Me.DtaConsulta.Recordset.MovePrevious
Loop

  
  
  

    FechaEgreso = Me.TxtUltFechaNomina.Value

    FechaContrato = Me.TxtFechaContrato.Text


    Año = Year(FechaBusqueda)
    Mes = Month(FechaBusqueda)
    
'    SQLSalarios = "SELECT DISTINCT" & vbLf
'    SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, dbo.DetalleNomina.SalarioBasico AS SalarioBasico, dbo.DetalleNomina.Destajo AS Destajo," & vbLf
'    SQLSalarios = SQLSalarios & "dbo.DetalleNomina.SeptimoDia AS Septimo, dbo.DetalleNomina.OtrosIngresos AS Otros, dbo.DetalleNomina.Incentivos AS Incentivos," & vbLf
'    SQLSalarios = SQLSalarios & "dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.OtrosIngresos AS TotalIngresos," & vbLf
'    SQLSalarios = SQLSalarios & "dbo.Nomina.FechaNominaINI AS FechaInicio, dbo.Nomina.FechaNomina AS FechaFin, dbo.Nomina.Mes, dbo.Nomina.Ano AS AÑO" & vbLf
'    SQLSalarios = SQLSalarios & "FROM         dbo.DetalleNomina INNER JOIN" & vbLf
'    SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
'    SQLSalarios = SQLSalarios & "WHERE     (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo <> 0) AND (dbo.DetalleNomina.CodEmpleado = '" & Me.txtCodEmpleado.Text & "') AND" & vbLf
'    SQLSalarios = SQLSalarios & "(dbo.Nomina.FechaNomina BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "')" & vbLf
'    SQLSalarios = SQLSalarios & "ORDER BY dbo.Nomina.Ano, dbo.Nomina.Mes,dbo.Nomina.FechaNomina"
'

         If Me.AdoEmpresa.Recordset("FormatoNomina") = "Nomina Bono Produccion" Then
            If Me.ChkExtra.Value = 1 Then
                     SqlSalarios = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico AS SalarioBasico, DetalleNomina.Destajo AS Destajo,DetalleNomina.SeptimoDia AS Septimo, DetalleNomina.OtrosIngresos AS Otros, DetalleNomina.Incentivos AS Incentivos,DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones+DetalleNomina.BonoProduccion+DetalleNomina.HorasExtras+DetalleNomina.IncetivoProduccion AS TotalIngresos,Nomina.FechaNominaINI AS FechaInicio, Nomina.FechaNomina AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO,DetalleNomina.Comisiones AS Comisiones, DetalleNomina.BonoProduccion, DetalleNomina.HorasExtras, DetalleNomina.IncetivoProduccion, DetalleNomina.MontoINSS, DetalleNomina.MontoIR " & _
                                   "FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina  " & _
                                   "WHERE (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.OtrosIngresos <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (Nomina.FechaNomina BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') " & _
                                   "ORDER BY Nomina.Ano, Nomina.Mes, Nomina.FechaNomina "
                    
                    ArepDetalleLiquidaBono.DataControl1.ConnectionString = ConexionReporte
                    ArepDetalleLiquidaBono.lbltitulo.Caption = Titulo
                    ArepDetalleLiquidaBono.LblSubtitulo.Caption = SubTitulo
                    ArepDetalleLiquidaBono.ImgLogo.Picture = LoadPicture(RutaLogo)
                    
                    ArepDetalleLiquidaBono.lbltitulo.Caption = Titulo
                    ArepDetalleLiquidaBono.LblSubtitulo.Caption = SubTitulo
                    ArepDetalleLiquidaBono.ImgLogo.Picture = LoadPicture(RutaLogo)
                    ArepDetalleLiquidaBono.DataControl1.ConnectionString = ConexionReporte
                    ArepDetalleLiquidaBono.DataControl1.Source = SqlSalarios
                    
                    ArepDetalleLiquidaBono.LblCodEmpleado.Caption = Me.txtCodEmpleado1.Text
                    ArepDetalleLiquidaBono.LblNombreEmpleado.Caption = Me.TxtNombre1.Text + " " + Me.TxtNombre2.Text + " " + Me.TxtApellido1.Text + " " + Me.TxtApellido2.Text
                    ArepDetalleLiquidaBono.LblDepartamento.Caption = Me.txtDepartamento.Text
                    ArepDetalleLiquidaBono.LblCargo.Caption = Me.txtCargo.Text
                    ArepDetalleLiquidaBono.LblAños.Caption = Me.TxtAnnos.Text
                    ArepDetalleLiquidaBono.LblDias.Caption = Me.TxtDiasTrabajados.Text
                    ArepDetalleLiquidaBono.LblMeses.Caption = Me.TxtMeses.Text
                    
                    ArepDetalleLiquidaBono.Show 1
        '                   fPreview.arv.ReportSource = ArepDetalleLiquidaBono
        '                    fPreview.Show 1
           Else
                     SqlSalarios = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico AS SalarioBasico, DetalleNomina.Destajo AS Destajo,DetalleNomina.SeptimoDia AS Septimo, DetalleNomina.OtrosIngresos AS Otros, DetalleNomina.Incentivos AS Incentivos,DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones+DetalleNomina.BonoProduccion+DetalleNomina.IncetivoProduccion AS TotalIngresos,Nomina.FechaNominaINI AS FechaInicio, Nomina.FechaNomina AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO,DetalleNomina.Comisiones AS Comisiones, DetalleNomina.BonoProduccion, DetalleNomina.HorasExtras, DetalleNomina.IncetivoProduccion, DetalleNomina.MontoINSS, DetalleNomina.MontoIR " & _
                                   "FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina  " & _
                                   "WHERE (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.OtrosIngresos <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (Nomina.FechaNomina BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') " & _
                                   "ORDER BY Nomina.Ano, Nomina.Mes, Nomina.FechaNomina "
                    
                    ArepDetalleLiquidaBono2.DataControl1.ConnectionString = ConexionReporte
                    ArepDetalleLiquidaBono2.lbltitulo.Caption = Titulo
                    ArepDetalleLiquidaBono2.LblSubtitulo.Caption = SubTitulo
                    ArepDetalleLiquidaBono2.ImgLogo.Picture = LoadPicture(RutaLogo)
                    
                    ArepDetalleLiquidaBono2.lbltitulo.Caption = Titulo
                    ArepDetalleLiquidaBono2.LblSubtitulo.Caption = SubTitulo
                    ArepDetalleLiquidaBono2.ImgLogo.Picture = LoadPicture(RutaLogo)
                    ArepDetalleLiquidaBono2.DataControl1.ConnectionString = ConexionReporte
                    ArepDetalleLiquidaBono2.DataControl1.Source = SqlSalarios
                    
                    ArepDetalleLiquidaBono2.LblCodEmpleado.Caption = Me.txtCodEmpleado1.Text
                    ArepDetalleLiquidaBono2.LblNombreEmpleado.Caption = Me.TxtNombre1.Text + " " + Me.TxtNombre2.Text + " " + Me.TxtApellido1.Text + " " + Me.TxtApellido2.Text
                    ArepDetalleLiquidaBono2.LblDepartamento.Caption = Me.txtDepartamento.Text
                    ArepDetalleLiquidaBono2.LblCargo.Caption = Me.txtCargo.Text
                    ArepDetalleLiquidaBono2.LblAños.Caption = Me.TxtAnnos.Text
                    ArepDetalleLiquidaBono2.LblDias.Caption = Me.TxtDiasTrabajados.Text
                    ArepDetalleLiquidaBono2.LblMeses.Caption = Me.TxtMeses.Text
                    
                    ArepDetalleLiquidaBono2.Show 1
           
           End If
         
         
         
         Else
             SqlSalarios = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico AS SalarioBasico, DetalleNomina.Destajo AS Destajo,DetalleNomina.SeptimoDia AS Septimo, DetalleNomina.OtrosIngresos AS Otros, DetalleNomina.Incentivos AS Incentivos,DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones + DetalleNomina.Incentivos AS TotalIngresos,Nomina.FechaNominaINI AS FechaInicio, Nomina.FechaNomina AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO,DetalleNomina.Comisiones AS Comisiones " & _
                           "FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina  " & _
                           "WHERE (DetalleNomina.SalarioBasico + DetalleNomina.Destajo <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (Nomina.FechaNomina BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') " & _
                           "ORDER BY Nomina.Ano, Nomina.Mes, Nomina.FechaNomina "
            
            ArepDetalleLiquida.DataControl1.ConnectionString = ConexionReporte
            ArepDetalleLiquida.lbltitulo.Caption = Titulo
            ArepDetalleLiquida.LblSubtitulo.Caption = SubTitulo
            ArepDetalleLiquida.ImgLogo.Picture = LoadPicture(RutaLogo)
            
            ArepDetalleLiquida.lbltitulo.Caption = Titulo
            ArepDetalleLiquida.LblSubtitulo.Caption = SubTitulo
            ArepDetalleLiquida.ImgLogo.Picture = LoadPicture(RutaLogo)
            ArepDetalleLiquida.DataControl1.ConnectionString = ConexionReporte
            ArepDetalleLiquida.DataControl1.Source = SqlSalarios
            
            ArepDetalleLiquida.LblCodEmpleado.Caption = Me.txtCodEmpleado1.Text
            ArepDetalleLiquida.LblNombreEmpleado.Caption = Me.TxtNombre1.Text + " " + Me.TxtNombre2.Text + " " + Me.TxtApellido1.Text + " " + Me.TxtApellido2.Text
            ArepDetalleLiquida.LblDepartamento.Caption = Me.txtDepartamento.Text
            ArepDetalleLiquida.LblCargo.Caption = Me.txtCargo.Text
            ArepDetalleLiquida.LblAños.Caption = Me.TxtAnnos.Text
            ArepDetalleLiquida.LblDias.Caption = Me.TxtDiasTrabajados.Text
            ArepDetalleLiquida.LblMeses.Caption = Me.TxtMeses.Text
            
        '    ArepDetalleLiquida.Show 1
                   fPreview.arv.ReportSource = ArepDetalleLiquida
                   fPreview.Show 1
                   
         End If
        

Exit Sub
TipoErrs:
MsgBox Err.Description

End Sub

Private Sub CmdEfectuar_Click()
'On Error GoTo TipoErr
Dim i As Integer, NumNomina As Integer
Dim SalMayor As Double, sql As String
Dim SalTemp As Double
Dim SalBrutoTemp As Double
Dim SalBrutoMayor As Double
Dim Mes As Byte
Dim DiaMes As Double
Dim DiaSemana As Double
Dim Mes13 As Double
Dim VACACIONES As Double
Dim MontoAntiguedad As Double
Dim CargoConfianza As Double
Dim TextOtro As String
Dim Otro As Double
Dim Prestamo As Double
Dim Deducciones As Double
Dim MontoNomPropor As Double 'salario de nomina proporcional
Dim rs As New ADODB.Recordset

CodEmpleado = TxtCodEmpleado.Text
'extraigo el salario mayor de los ultimos seis meses
SalMayor = 0
SalBrutoMayor = 0

k% = MsgBox("Desea realmente dar de baja a este Empleado?", vbYesNo)
If k <> 6 Then Exit Sub
Anno = Year(Now)


res = Bitacora(Now, NombreUsuario, "Liquidacion", "Se Proceso Liquidacion: " & Me.txtCodEmpleado1.Text & " " & Me.TxtNombre1.Text)

'DtaEmpleado.Recordset.Edit
DtaEmpleado.Recordset("activo") = False
DtaEmpleado.Recordset.Update

 '/////////////// Deducciones //////////////////////////
        SQlDeducciones = "SELECT Deduccion.CodTipoDeduccion, DetalleDeduccion.NumDeduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.Pagado, DetalleDeduccion.NumNomina, Deduccion.CodEmpleado FROM Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion WHERE DetalleDeduccion.Pagado=0 AND Deduccion.CodEmpleado= '" & CodEmpleado & "'"
        DtaDeducciones.RecordSource = SQlDeducciones
        DtaDeducciones.Refresh

If Me.ChkPrestamo.Value = 1 Then
'/////////////////////////////////////////////////////////////////////////
'//////////////Busco si el empleado tiene Prestamo////////////////////////
'/////////////////////////////////////////////////////////////////////////
Me.Dtaprestamo.RecordSource = "SELECT Prestamo.CodEmpleado, Prestamo.Cancelado From Prestamo Where (((Prestamo.CodEmpleado) = " & CodigoEmpleado & "))"
Dtaprestamo.Refresh
    Prestamo = 0
Do While Not Dtaprestamo.Recordset.EOF

   'DtaPrestamo.Recordset.Edit
   Dtaprestamo.Recordset("cancelado") = True
   Dtaprestamo.Recordset.Update


 Me.Dtaprestamo.Recordset.MoveNext
Loop

End If

If Me.ChkDeducciones.Value = 1 Then
'/////////////////////////////////////////////////////////////
'////////////Cancelo deducciones///////////////////////////////
'///////////////////////////////////////////////////////////////
DtaDeducciones.Refresh
Do While Not DtaDeducciones.Recordset.EOF
    If DtaDeducciones.Recordset("CodEmpleado") = TxtCodEmpleado Then
       Deducciones = Deducciones + DtaDeducciones.Recordset("valor")
       'DtaDeducciones.Recordset.Edit
       DtaDeducciones.Recordset("Pagado") = True
       DtaDeducciones.Recordset.Update
    Else
     Deducciones = 0
    End If
  If Not Me.DtaDeducciones.Recordset.EOF Then
    DtaDeduccion.Recordset.MoveNext
  End If
Loop

End If

'/////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////BUSCO SI TIENE NOMINAS ACTIVAS PARA BORRAR EL REGISTRO//////
'////////////////////////////////////////////////////////////////////////////////////////

sql = "SELECT   DetalleNomina.NumNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HE, " & vbLf
sql = sql & "DetalleNomina.DD, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre," & vbLf
sql = sql & "DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
sql = sql & "DetalleNomina.MontoINSS, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.INATEC, DetalleNomina.Mes13," & vbLf
sql = sql & "DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.TotalSubsidio, DetalleNomina.VacacionesPagadas," & vbLf
sql = sql & "DetalleNomina.DiasVacaciones, DetalleNomina.AdelantosVacaciones, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia," & vbLf
sql = sql & "DetalleNomina.IncetivoProduccion , DetalleNomina.TarifaHoraria, Nomina.Activa" & vbLf
sql = sql & "FROM         DetalleNomina INNER JOIN" & vbLf
sql = sql & "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina" & vbLf
sql = sql & "Where (DetalleNomina.CodEmpleado = " & CodEmpleado & ") And (Nomina.Activa = 1)"
Me.DtaConsulta.RecordSource = sql
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
  NumNomina = Me.DtaConsulta.Recordset("NumNomina")
Else
  NumNomina = -1
End If

rs.Open "DELETE FROM DetalleNomina Where (NumNomina = " & NumNomina & ") And (CodEmpleado = " & CodEmpleado & ")", Conexion

'//////////////////////////////////////////////////////////////////////////////////////////
'/////////////BUSCO LA NOMINA DE SUBSIDIOS PARA BORRAR EL REGISTRO DEL EMPLEADO///////////
'///////////////////////////////////////////////////////////////////////////////////////////
sql = "SELECT DetalleNomSubsidio.NumNominaSubsidio, DetalleNomSubsidio.CodEmpleado, DetalleNomSubsidio.Subsidio, NomSubsidio.Activa" & vbLf
sql = sql & "FROM DetalleNomSubsidio INNER JOIN NomSubsidio ON DetalleNomSubsidio.NumNominaSubsidio = NomSubsidio.NumNomina" & vbLf
sql = sql & "Where (NomSubsidio.Activa = 1) And (DetalleNomSubsidio.CodEmpleado = " & CodEmpleado & ")"
Me.DtaConsulta.RecordSource = sql
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
  NumNomina = Me.DtaConsulta.Recordset("NumNominaSubsidio")
Else
  NumNomina = -1
End If

rs.Open "DELETE FROM DetalleNomSubsidio Where (NumNominaSubsidio = " & NumNomina & ") And (CodEmpleado = " & CodEmpleado & ")", Conexion


'//////////////////////////////////////////////////////////////////////////////////////////
'/////////////BUSCO SI TIENE NOMINA DE VACACIONES PARA BORRAR EL REGISTRO/////////////////
'//////////////////////////////////////////////////////////////////////////////////////////
sql = "SELECT     DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, " & vbLf
sql = sql & "                      DetalleNomVaca.DiasDescuento , DetalleNomVaca.AdelantoVacaciones, DetalleNomVaca.Inss, DetalleNomVaca.TarifaHoraria, NomVaca.Activa" & vbLf
sql = sql & "FROM         DetalleNomVaca INNER JOIN" & vbLf
sql = sql & "                      NomVaca ON DetalleNomVaca.NumNomVaca = NomVaca.NumNomVaca" & vbLf
sql = sql & "Where (DetalleNomVaca.CodEmpleado = " & CodEmpleado & ") And (NomVaca.Activa = 1)"
Me.DtaConsulta.RecordSource = sql
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
  NumNomina = Me.DtaConsulta.Recordset("NumNomVaca")
Else
  NumNomina = -1
End If

rs.Open "DELETE FROM DetalleNomVaca Where (NumNomVaca = " & NumNomina & ") And (CodEmpleado = " & CodEmpleado & ")", Conexion


'////////////////////////////////////////////////////////////////////////////////////////////////
'////////BUSCO LA NOMINA DE 13VO MES PARA BORRAR EL REGISTRO DEL EMPLEADO///////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////
sql = "SELECT     DetalleNom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.SalarioAPagar, " & vbLf
sql = sql & "DetalleNom13Mes.DiasAPagar , DetalleNom13Mes.Adelanto13vo, Nom13Mes.Activa" & vbLf
sql = sql & "FROM         DetalleNom13Mes INNER JOIN" & vbLf
sql = sql & "Nom13Mes ON DetalleNom13Mes.NumNom13Mes = Nom13Mes.NumNom13Mes" & vbLf
sql = sql & "Where (Nom13Mes.Activa = 1) And (DetalleNom13Mes.CodEmpleado = " & CodEmpleado & ")"
Me.DtaConsulta.RecordSource = sql
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
  NumNomina = Me.DtaConsulta.Recordset("NumNom13Mes")
Else
  NumNomina = -1
End If

rs.Open "DELETE FROM DetalleNom13Mes Where (NumNom13Mes = " & NumNomina & ") And (CodEmpleado = " & CodEmpleado & ")", Conexion



    

    MsgBox "Se Proceso con Existo!!", vbExclamation, "Zeus Nominas"
   
    Me.CmdEfectuar.Enabled = False
    Me.CmdRenovar.Enabled = False

'CmdCancelar.Caption = "Cerrar"
Exit Sub
TipoErr:
ControlErrores
End Sub

Private Sub ChkVaca_Click()
If ChkVaca.Value = 1 Then
Me.TxtDescuentoDias.Visible = True
Me.TxtDiasDescuento.Visible = True
Else
    

Me.TxtDiasDescuento.Visible = False
Me.TxtDescuentoDias.Visible = False
End If

End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdImprimirHistorial_Click()
On Error GoTo TipoErrs
Dim SqlSalarios As String, FechaEgreso As Date, FechaContrato As Date, FechaHistorico As Date, FechaBusqueda As Date
Dim CodEmpleado As String, NumeroEmpleado As Integer, i As Integer
Dim rpt As Object
Dim fPreview As New FrmPreview

    CodEmpleado = TxtCodEmpleado.Text
  
  '///////////Busco la Fecha para la Busqueda////////////////////////////
  
FechaEgreso = Me.TxtFechaHistorial.Value
'FechaContrato = Me.TxtFechaContrato.Text


NumeroEmpleado = Me.TxtCodEmpleado.Text

SqlSalarios = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.Incentivos) AS Incentivos, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.Comisiones) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes AS MES, Nomina.Ano AS AÑO, SUM(DetalleNomina.Comisiones) As Comisiones  " & _
              "FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
              "HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") AND (MAX(Nomina.FechaNomina)<= CONVERT(DATETIME, '" & Format(FechaEgreso, "yyyy/mm/dd") & "', 102))"

'SQLSalarios = "SELECT DISTINCT" & vbLf
'SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
'SQLSalarios = SQLSalarios & "AS Destajo, SUM(dbo.DetalleNomina.Incentivos) AS Incentivos, SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo)" & vbLf
'SQLSalarios = SQLSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes AS MES," & vbLf
'SQLSalarios = SQLSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
'SQLSalarios = SQLSalarios & "FROM   dbo.DetalleNomina INNER JOIN" & vbLf
'SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
'SQLSalarios = SQLSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
'SQLSalarios = SQLSalarios & "Having (dbo.DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") And (Sum(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0) AND (MAX(Nomina.FechaNomina) <= CONVERT(DATETIME, '" & Format(FechaEgreso, "yyyy/mm/dd") & "', 102))"
Me.DtaConsulta.RecordSource = SqlSalarios
Me.DtaConsulta.Refresh

If Not Me.DtaConsulta.Recordset.EOF Then
Me.DtaConsulta.Recordset.MoveLast
End If
i = 0
Do While Not Me.DtaConsulta.Recordset.BOF
  If i = 1 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")

  ElseIf i = 5 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    Exit Do
  ElseIf i = 0 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    FechaHistorico = Me.DtaConsulta.Recordset("FechaFin")
  Else
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
  End If
  i = i + 1

  Me.DtaConsulta.Recordset.MovePrevious
Loop

  
  
  

    FechaEgreso = Me.TxtUltFechaNomina.Value
'    FechaHistorico = DateSerial(Year(FechaEgreso), Month(FechaEgreso), 1 - 1)
    FechaContrato = Me.TxtFechaContrato.Text
'    FechaBusqueda = DateSerial(Year(FechaEgreso), Month(FechaEgreso) - 6, 1)

    Año = Year(FechaBusqueda)
    Mes = Month(FechaBusqueda)
    
    
'    SQLSalarios = "SELECT DISTINCT" & vbLf
'    SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
'    SQLSalarios = SQLSalarios & "AS Destajo, SUM(dbo.DetalleNomina.SeptimoDia) AS Septimo, SUM(dbo.DetalleNomina.OtrosIngresos) AS Otros, SUM(dbo.DetalleNomina.Incentivos)" & vbLf
'    SQLSalarios = SQLSalarios & "AS Incentivos," & vbLf
'    SQLSalarios = SQLSalarios & "SUM (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.OtrosIngresos)" & vbLf
'    SQLSalarios = SQLSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes," & vbLf
'    SQLSalarios = SQLSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
'    SQLSalarios = SQLSalarios & "FROM    dbo.DetalleNomina INNER JOIN" & vbLf
'    SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
'    SQLSalarios = SQLSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
'    SQLSalarios = SQLSalarios & "HAVING(SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0) And (DetalleNomina.CodEmpleado = '" & Me.txtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND" & vbLf
'    SQLSalarios = SQLSalarios & "'" & Format(FechaHistorico, "yyyymmdd") & "')" & vbLf
'    SQLSalarios = SQLSalarios & "ORDER BY dbo.Nomina.Ano, dbo.Nomina.Mes"


        If Me.AdoEmpresa.Recordset("FormatoNomina") = "Nomina Bono Produccion" Then
        
          If Me.ChkExtra.Value = 1 Then
                   SqlSalarios = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion) AS Incentivos, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.BonoProduccion + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.HorasExtras + DetalleNomina.IncetivoProduccion) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.BonoProduccion) AS BonoProduccion, SUM(DetalleNomina.HorasExtras) AS HorasExtras, " & _
                                 "SUM(DetalleNomina.MontoIR)AS MontoIR FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                                 "HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.OtrosIngresos) <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) Between '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') ORDER BY Nomina.Ano, Nomina.Mes"
                   
        
                    
                    ArepHistorialLiquidaBono.lbltitulo.Caption = Titulo
                    ArepHistorialLiquidaBono.LblSubtitulo.Caption = SubTitulo
                    ArepHistorialLiquidaBono.ImgLogo.Picture = LoadPicture(RutaLogo)
                
                    ArepHistorialLiquidaBono.LblCodEmpleado.Caption = Me.txtCodEmpleado1.Text
                    ArepHistorialLiquidaBono.LblNombreEmpleado.Caption = Me.TxtNombre1.Text + " " + Me.TxtNombre2.Text + " " + Me.TxtApellido1.Text + " " + Me.TxtApellido2.Text
                    ArepHistorialLiquidaBono.LblDepartamento.Caption = Me.txtDepartamento.Text
                    ArepHistorialLiquidaBono.LblCargo.Caption = Me.txtCargo.Text
                    ArepHistorialLiquidaBono.LblAños.Caption = Me.TxtAnnos.Text
                    ArepHistorialLiquidaBono.LblDias.Caption = Me.TxtDiasTrabajados.Text
                    ArepHistorialLiquidaBono.LblMeses.Caption = Me.TxtMeses.Text
                    
                    ArepHistorialLiquidaBono.LblSalarioAlto.Caption = Me.TxtSalarioAlto.Text
                    ArepHistorialLiquidaBono.LblSalarioBasico.Caption = Me.txtSalarioBasico.Text
                    ArepHistorialLiquidaBono.LblSalarioPromedio.Caption = Me.TxtSalarioPromedio.Text
                    ArepHistorialLiquidaBono.LblTarifaHoraria.Caption = Me.TxtTarifa.Text
                    ArepHistorialLiquidaBono.lblVacaciones.Caption = Me.DTPFechaIniVaca.Value
                    ArepHistorialLiquidaBono.LblAguinaldo.Caption = Me.DTPFechaIniAgui.Value
                
                    ArepHistorialLiquidaBono.DataControl1.ConnectionString = ConexionReporte
                    ArepHistorialLiquidaBono.DataControl1.Source = SqlSalarios
                    ArepHistorialLiquidaBono.Show 1
           Else
                   SqlSalarios = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion) AS Incentivos, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.BonoProduccion + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos +  DetalleNomina.IncetivoProduccion) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.BonoProduccion) AS BonoProduccion,SUM(DetalleNomina.MontoIR)AS MontoIR FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                                 "HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.OtrosIngresos) <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) Between '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') ORDER BY Nomina.Ano, Nomina.Mes"
                   
        
                    
                    ArepHistorialLiquidaBono2.lbltitulo.Caption = Titulo
                    ArepHistorialLiquidaBono2.LblSubtitulo.Caption = SubTitulo
                    ArepHistorialLiquidaBono2.ImgLogo.Picture = LoadPicture(RutaLogo)
                
                    ArepHistorialLiquidaBono2.LblCodEmpleado.Caption = Me.txtCodEmpleado1.Text
                    ArepHistorialLiquidaBono2.LblNombreEmpleado.Caption = Me.TxtNombre1.Text + " " + Me.TxtNombre2.Text + " " + Me.TxtApellido1.Text + " " + Me.TxtApellido2.Text
                    ArepHistorialLiquidaBono2.LblDepartamento.Caption = Me.txtDepartamento.Text
                    ArepHistorialLiquidaBono2.LblCargo.Caption = Me.txtCargo.Text
                    ArepHistorialLiquidaBono2.LblAños.Caption = Me.TxtAnnos.Text
                    ArepHistorialLiquidaBono2.LblDias.Caption = Me.TxtDiasTrabajados.Text
                    ArepHistorialLiquidaBono2.LblMeses.Caption = Me.TxtMeses.Text
                    
                    ArepHistorialLiquidaBono2.LblSalarioAlto.Caption = Me.TxtSalarioAlto.Text
                    ArepHistorialLiquidaBono2.LblSalarioBasico.Caption = Me.txtSalarioBasico.Text
                    ArepHistorialLiquidaBono2.LblSalarioPromedio.Caption = Me.TxtSalarioPromedio.Text
                    ArepHistorialLiquidaBono2.LblTarifaHoraria.Caption = Me.TxtTarifa.Text
                    ArepHistorialLiquidaBono2.lblVacaciones.Caption = Me.DTPFechaIniVaca.Value
                    ArepHistorialLiquidaBono2.LblAguinaldo.Caption = Me.DTPFechaIniAgui.Value
                
                    ArepHistorialLiquidaBono2.DataControl1.ConnectionString = ConexionReporte
                    ArepHistorialLiquidaBono2.DataControl1.Source = SqlSalarios
                    ArepHistorialLiquidaBono2.Show 1
           End If
'
'           ArepHistorialLiquidaBono.Refresh
'           fPreview.arv.ReportSource = ArepHistorialLiquidaBono
'
'
'           fPreview.Show 1
'

        
        Else
        
            SqlSalarios = " SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo,SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones + DetalleNomina.Incentivos) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Comisiones) As Comisiones " & _
                          "FROM DetalleNomina INNER JOIN  Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                          "HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) Between '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "')" & _
                          "ORDER BY Nomina.Ano, Nomina.Mes "
            
            ArepHistorialLiquida.DataControl1.ConnectionString = ConexionReporte
            ArepHistorialLiquida.lbltitulo.Caption = Titulo
            ArepHistorialLiquida.LblSubtitulo.Caption = SubTitulo
            ArepHistorialLiquida.ImgLogo.Picture = LoadPicture(RutaLogo)

            ArepHistorialLiquida.LblCodEmpleado.Caption = Me.txtCodEmpleado1.Text
            ArepHistorialLiquida.LblNombreEmpleado.Caption = Me.TxtNombre1.Text + " " + Me.TxtNombre2.Text + " " + Me.TxtApellido1.Text + " " + Me.TxtApellido2.Text
            ArepHistorialLiquida.LblDepartamento.Caption = Me.txtDepartamento.Text
            ArepHistorialLiquida.LblCargo.Caption = Me.txtCargo.Text
            ArepHistorialLiquida.LblAños.Caption = Me.TxtAnnos.Text
            ArepHistorialLiquida.LblDias.Caption = Me.TxtDiasTrabajados.Text
            ArepHistorialLiquida.LblMeses.Caption = Me.TxtMeses.Text

            ArepHistorialLiquida.LblSalarioAlto.Caption = Me.TxtSalarioAlto.Text
            ArepHistorialLiquida.LblSalarioBasico.Caption = Me.txtSalarioBasico.Text
            ArepHistorialLiquida.LblSalarioPromedio.Caption = Me.TxtSalarioPromedio.Text
            ArepHistorialLiquida.LblTarifaHoraria.Caption = Me.TxtTarifa.Text
            ArepHistorialLiquida.lblVacaciones.Caption = Me.DTPFechaIniVaca.Value
            ArepHistorialLiquida.LblAguinaldo.Caption = Me.DTPFechaIniAgui.Value

            ArepHistorialLiquida.DataControl1.ConnectionString = ConexionReporte
            ArepHistorialLiquida.DataControl1.Source = SqlSalarios
'            ArepHistorialLiquida.Show 1
           fPreview.arv.ReportSource = ArepHistorialLiquida
           fPreview.Show 1

         End If
    
    
 Exit Sub
TipoErrs:
 MsgBox Err.Description
 
End Sub


Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub CmdRenovar_Click()
Dim CodEmpleado As Double, CodEmpleado1 As String, CodEmpleadoAnterior As Double
Dim FechaNacimiento As Date, Id As Double, Contador As Double, Contador2 As Double
Dim Mes As Integer, CodigoNuevo As String



 
  Respuesta = MsgBox("Desea reutilizar Codigos?", vbYesNo, "Zeus Facturacion")
 
 
 If Respuesta = 6 Then
    CodEmpleado = Me.TxtCodEmpleado.Text
    CodEmpleado1 = Me.txtCodEmpleado1.Text
 Else
    Mes = Month(Now)
    CodigoNuevo = "S1" & Mid(Year(Now), Len(Year(Now)) - 1, Len(Year(Now)) - 2) & Format(Mes, "0#")
    Me.DtaConsulta.RecordSource = "SELECT CodEmpleado1 From Empleado WHERE (CodEmpleado1 LIKE '%" & CodigoNuevo & "%') ORDER BY CodEmpleado1 DESC"
    Me.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
      Numero = Mid(Me.DtaConsulta.Recordset("CodEmpleado1"), 7, 4)
      CodigoNuevo = CodigoNuevo & Format(Numero + 1, "000#")
    Else
      CodigoNuevo = CodigoNuevo & "0001"
    End If
    
      CodEmpleado = Me.TxtCodEmpleado.Text
      CodEmpleado1 = CodigoNuevo
      
 End If
 
 
 
 Me.CmdEfectuar.Value = True
 
 FrmFechaIngresoBaja.Show 1
 
 res = Bitacora(Now, NombreUsuario, "Liquidacion", "Se Renovo Contrato: " & Me.txtCodEmpleado1.Text & " " & Me.TxtNombre1.Text)
 
 Me.DtaConsulta.RecordSource = "SELECT * From Empleado Where (CodEmpleado = " & CodEmpleado & ")"
 Me.DtaConsulta.Refresh
 If Not Me.DtaConsulta.Recordset.EOF Then
 
 
       adoEmpleado.Recordset.AddNew
'            CodEmpleado1 = Me.DtaConsulta.Recordset("CodEmpleado1")
'            AdoEmpleado.Recordset("CodEmpleado1") = Me.DtaConsulta.Recordset("CodEmpleado1")
            adoEmpleado.Recordset("CodEmpleado1") = CodEmpleado1
            adoEmpleado.Recordset("Nombre1") = Me.DtaConsulta.Recordset("Nombre1")
            adoEmpleado.Recordset("Nombre2") = Me.DtaConsulta.Recordset("Nombre2")
            adoEmpleado.Recordset("Apellido1") = Me.DtaConsulta.Recordset("Apellido1")
            adoEmpleado.Recordset("Apellido2") = Me.DtaConsulta.Recordset("Apellido2")
            adoEmpleado.Recordset("Direccion") = Me.DtaConsulta.Recordset("Direccion")
            adoEmpleado.Recordset("Nacionalidad") = Me.DtaConsulta.Recordset("Nacionalidad")
            adoEmpleado.Recordset("CodigoPostal") = Me.DtaConsulta.Recordset("CodigoPostal")
            adoEmpleado.Recordset("numcedula") = Me.DtaConsulta.Recordset("numcedula")
            adoEmpleado.Recordset("sexo") = Me.DtaConsulta.Recordset("sexo")
            adoEmpleado.Recordset("NumeroInss") = Me.DtaConsulta.Recordset("NumeroInss")
            adoEmpleado.Recordset("numeroruc") = Me.DtaConsulta.Recordset("numeroruc")
            adoEmpleado.Recordset("CodDepartamento") = Me.DtaConsulta.Recordset("CodDepartamento")
            adoEmpleado.Recordset("CodCargo") = Me.DtaConsulta.Recordset("CodCargo")
            adoEmpleado.Recordset("Codgrupo") = Me.DtaConsulta.Recordset("Codgrupo")
            adoEmpleado.Recordset("Sindicalista") = Me.DtaConsulta.Recordset("Sindicalista")
            adoEmpleado.Recordset("CodTipoNomina") = Me.DtaConsulta.Recordset("CodTipoNomina")
            adoEmpleado.Recordset("numhijos") = Me.DtaConsulta.Recordset("numhijos")
            adoEmpleado.Recordset("PorcientoIncentivo") = 0
            adoEmpleado.Recordset("SueldoPeriodo") = Me.DtaConsulta.Recordset("SueldoPeriodo")
            adoEmpleado.Recordset("TarifaHoraria") = Me.DtaConsulta.Recordset("TarifaHoraria")
            adoEmpleado.Recordset("PorcentajeComision") = Me.DtaConsulta.Recordset("PorcentajeComision")
            adoEmpleado.Recordset("OtrosIngresos") = Me.DtaConsulta.Recordset("OtrosIngresos")
            adoEmpleado.Recordset("salariominimo") = Me.DtaConsulta.Recordset("salariominimo")
            adoEmpleado.Recordset("ExentoInss") = Me.DtaConsulta.Recordset("ExentoInss")
            adoEmpleado.Recordset("ExentoIr") = Me.DtaConsulta.Recordset("ExentoIr")
            adoEmpleado.Recordset("PagoInssPatronal") = Me.DtaConsulta.Recordset("PagoInssPatronal")
            adoEmpleado.Recordset("CuentaBanco") = Me.DtaConsulta.Recordset("CuentaBanco")
      adoEmpleado.Recordset.Update
      
      '///////////////////////////////////////////////////////////////////////////////////////////////
      '//////////////////////////BUSCO LA FECHA DE NACIMIENTO DE EMPLEADO DADO DE BAJA /////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////////////
      CodEmpleadoAnterior = CodEmpleado
      Me.DtaConsulta.RecordSource = "SELECT  * From Historico Where (CodEmpleado = " & CodEmpleado & ")"
      Me.DtaConsulta.Refresh
      If Not Me.DtaConsulta.Recordset.EOF Then
        FechaNacimiento = Me.DtaConsulta.Recordset("FechaNacimiento")
      End If

      
      '///////////////////////////////////////////////////////////////////////////////////////////////
      '//////////////////////////BUSCO EL CODIGO INTERNO PARA EL EMPLEADO QUE ACABO DE GRABAR /////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////////////
      Me.DtaConsulta.RecordSource = "SELECT * From Empleado WHERE (CodEmpleado1 = '" & CodEmpleado1 & "') AND (Activo = 1) ORDER BY CodEmpleado"
      Me.DtaConsulta.Refresh
      If Not Me.DtaConsulta.Recordset.EOF Then
        CodEmpleado = Me.DtaConsulta.Recordset("CodEmpleado")
      End If
      
      
      '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      '/////////////////////////////GRABO LA FECHA DE INGRESO Y VACACIONES ///////////////////////////////////////////
      '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      
       Me.DtaConsulta.RecordSource = "SELECT  * From Historico"
       Me.DtaConsulta.Refresh
       If Not Me.DtaConsulta.Recordset.EOF Then
         Me.DtaConsulta.Recordset.MoveLast
         Id = Me.DtaConsulta.Recordset("id") + 1
       Else
        Id = 1
       End If
      
       Me.AdoHistoricos.RecordSource = "SELECT  * From Historico Where (CodEmpleado = " & CodEmpleado & ")"
       Me.AdoHistoricos.Refresh
       If Me.AdoHistoricos.Recordset.EOF Then
         Me.AdoHistoricos.Recordset.AddNew
          Me.AdoHistoricos.Recordset("id") = Id
          Me.AdoHistoricos.Recordset("Codempleado") = CodEmpleado
          Me.AdoHistoricos.Recordset("FechaNacimiento") = FechaNacimiento
          Me.AdoHistoricos.Recordset("FechaContrato") = FechaIngreso
          Me.AdoHistoricos.Recordset("FechaContratoVac") = FechaIngreso
         Me.AdoHistoricos.Recordset.Update
       Else
          Me.AdoHistoricos.Recordset("FechaNacimiento") = FechaNacimiento
          Me.AdoHistoricos.Recordset("FechaContrato") = FechaIngreso
          Me.AdoHistoricos.Recordset("FechaContratoVac") = FechaIngreso
         Me.AdoHistoricos.Recordset.Update
       End If
       
       
       '//////////////////////////////////////////////////////////////////////////////////////////////////////////////77
       '///////////////////////////TRASLADO TODAS LAS NOMINAS QUE TENGA CON FECHA MAYOR A SU INGRESO ///////////////////
       '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
       Me.AdoTraslado.RecordSource = "SELECT  * FROM   DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina  " & _
                                     "WHERE (DetalleNomina.CodEmpleado = " & CodEmpleadoAnterior & ") AND (Nomina.FechaNomina >= CONVERT(DATETIME, '" & Format(FechaIngreso, "yyyy-mm-dd") & "', 102)) ORDER BY DetalleNomina.NumNomina"
       Me.AdoTraslado.Refresh
       Contador = 0
       Do While Not Me.AdoTraslado.Recordset.EOF
       
         Me.AdoTraslado.Recordset("CodEmpleado") = CodEmpleado
         Me.AdoTraslado.Recordset.Update
         Contador = Contador + 1
         Me.AdoTraslado.Recordset.MoveNext
       Loop
       
       
       '//////////////////////////////////////////////////////////////////////////////////////////////////////////////77
       '///////////////////////////TRASLADO TODAS LAS MARCADAS QUE TENGA CON FECHA MAYOR A SU INGRESO ///////////////////
       '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
       Me.AdoTraslado.RecordSource = "SELECT  * From AsistenciaEmpleado WHERE (FechaEntrada >= CONVERT(DATETIME, '" & Format(FechaIngreso, "yyyy-mm-dd") & "', 102)) AND (CodEmpleado = " & CodEmpleadoAnterior & ")"
       Me.AdoTraslado.Refresh
       Contador2 = 0
       Do While Not Me.AdoTraslado.Recordset.EOF
     
         Me.AdoTraslado.Recordset("CodEmpleado") = CodEmpleado
         Me.AdoTraslado.Recordset.Update
         Contador2 = Contador + 1
         Me.AdoTraslado.Recordset.MoveNext
       Loop
 
       Me.DtaTurnos.ConnectionString = Conexion
       Me.DtaTurnos.RecordSource = "SELECT CodEmpleado, LEntrada, LSalida, MEntrada, MSalida, MCEntrada, MCSalida, JEntrada, JSalida, VEntrada, VSalida, TComida, TurnoLunes,TurnoMartes , TurnoMiercoles, TurnoJueves, TurnoViernes, TurnoSabado, TurnoDomingo, SEntrada, SSalida, DEntrada, DSalida From dbo.HorarioEmpleado "
       Me.DtaTurnos.Refresh
  
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////GRABO EL HORARIO DE EMPLEADOS /////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////
       Me.DtaHorarioEmpleado.RecordSource = "SELECT CodEmpleado, LEntrada, LSalida, MEntrada, MSalida, MCEntrada, MCSalida, JEntrada, JSalida, VEntrada, VSalida, TComida, TurnoLunes,TurnoMartes , TurnoMiercoles, TurnoJueves, TurnoViernes, TurnoSabado, TurnoDomingo, SEntrada, SSalida, DEntrada, DSalida From dbo.HorarioEmpleado WHERE(CodEmpleado ='" & CodEmpleado1 & "')"
       Me.DtaHorarioEmpleado.Refresh
       If Me.DtaHorarioEmpleado.Recordset.EOF Then
         Me.DtaTurnos.Refresh
         If Not Me.DtaTurnos.Recordset.EOF Then
'           CodTurno = Me.DtaTurnos.Recordset("CodTurno")
         Me.DtaHorarioEmpleado.Recordset.AddNew
           Me.DtaHorarioEmpleado.Recordset("CodEmpleado") = CodEmpleado1
           Me.DtaHorarioEmpleado.Recordset("LEntrada") = Me.DtaTurnos.Recordset("LEntrada")
           Me.DtaHorarioEmpleado.Recordset("LSalida") = Me.DtaTurnos.Recordset("LSalida")
           Me.DtaHorarioEmpleado.Recordset("MEntrada") = Me.DtaTurnos.Recordset("MEntrada")
           Me.DtaHorarioEmpleado.Recordset("MSalida") = Me.DtaTurnos.Recordset("MSalida")
           Me.DtaHorarioEmpleado.Recordset("MCEntrada") = Me.DtaTurnos.Recordset("MCEntrada")
           Me.DtaHorarioEmpleado.Recordset("MCSalida") = Me.DtaTurnos.Recordset("MCSalida")
           Me.DtaHorarioEmpleado.Recordset("JEntrada") = Me.DtaTurnos.Recordset("JEntrada")
           Me.DtaHorarioEmpleado.Recordset("JSalida") = Me.DtaTurnos.Recordset("JSalida")
           Me.DtaHorarioEmpleado.Recordset("VEntrada") = Me.DtaTurnos.Recordset("VEntrada")
           Me.DtaHorarioEmpleado.Recordset("VSalida") = Me.DtaTurnos.Recordset("VSalida")
           Me.DtaHorarioEmpleado.Recordset("TComida") = Me.DtaTurnos.Recordset("TComida")
           Me.DtaHorarioEmpleado.Recordset("TurnoLunes") = Me.DtaTurnos.Recordset("TurnoLunes")
           Me.DtaHorarioEmpleado.Recordset("TurnoMartes") = Me.DtaTurnos.Recordset("TurnoMartes")
           Me.DtaHorarioEmpleado.Recordset("TurnoMiercoles") = Me.DtaTurnos.Recordset("TurnoMiercoles")
           Me.DtaHorarioEmpleado.Recordset("TurnoJueves") = Me.DtaTurnos.Recordset("TurnoJueves")
           Me.DtaHorarioEmpleado.Recordset("TurnoViernes") = Me.DtaTurnos.Recordset("TurnoViernes")
           Me.DtaHorarioEmpleado.Recordset("TurnoSabado") = Me.DtaTurnos.Recordset("TurnoSabado")
           Me.DtaHorarioEmpleado.Recordset("TurnoDomingo") = Me.DtaTurnos.Recordset("TurnoDomingo")
           Me.DtaHorarioEmpleado.Recordset("SEntrada") = Me.DtaTurnos.Recordset("SEntrada")
           Me.DtaHorarioEmpleado.Recordset("SSalida") = Me.DtaTurnos.Recordset("SEntrada")
           Me.DtaHorarioEmpleado.Recordset("DEntrada") = Me.DtaTurnos.Recordset("SEntrada")
           Me.DtaHorarioEmpleado.Recordset("DSalida") = Me.DtaTurnos.Recordset("SEntrada")
    
         Me.DtaHorarioEmpleado.Recordset.Update
         End If
       End If

 
  MsgBox "Se trasladoron " & Contador & " Nominas" & " y " & Contador2 & " Marcadas "
 End If
 
 
End Sub

Private Sub Form_Activate()
'txtCodEmpleado.Text = CodEmpleado
End Sub

Public Function CalcularMontoIrBajas(CodEmpleado As String, IrUltimaSemana As Boolean, CodTipoNomina As String, FechaBaja As Date, MontoBaja As Double) As Double
  Dim sql As String, Pmes As Double, PAno As Double, pPeriodo As Double, Periodo As Double
  Dim MontoNomina As Double
  '//////////////////////////////////////////////////////////////////////////////
  '///////////////////////CON ESTA CONSULTA UBICO EL MES Y EL AÑO /////////////////
  '/////////////////////////////////////////////////////////////////////////////////////
   sql = "SELECT NumNomina, Periodo, año, mes, Inicio, Final, CodTipoNomina From Fecha_Planilla " & _
         "WHERE  (Inicio <= CONVERT(DATETIME, '" & Format(FechaBaja, "yyyy-mm-dd") & "', 102)) AND (Final >= CONVERT(DATETIME, '" & Format(FechaBaja, "yyyy-mm-dd") & "', 102)) AND (CodTipoNomina = '" & CodTipoNomina & "')"
   Me.DtaConsulta.RecordSource = sql
   Me.DtaConsulta.Refresh
   If Not Me.DtaConsulta.Recordset.EOF Then
     Pmes = Me.DtaConsulta.Recordset("Mes")
     PAno = Me.DtaConsulta.Recordset("año")
     PFechaNomina = FechaBaja
   End If
  

                        If IrUltimaSemana = True Then
                         
                         ' AND (Inicio > CONVERT(DATETIME, '" & Format(PFechaNomina, "MM-dd-yyyy") & " 00:00:00', 102))
                             '/////////////////////////////////////////////////////////////////////////////////////////////////
                             '///////////////////////////////////BUSCO LA NOMINA ACTUAL /////////////////////////////////////////////////
                             '///////////////////////////////////////////////////////////////////////////////////////////////////////
                             Me.DtaConsulta.RecordSource = "SELECT     Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina   FROM  Fecha_Planilla  WHERE  (año = " & PAno & ") AND (mes ='" & Format(Pmes, "00") & "') AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"     'AND  (NumNomina IS NULL)"
                             Me.DtaConsulta.Refresh
                             If Not Me.DtaConsulta.Recordset.EOF Then
                             
                               '////////////////////////SI ENCUENTRA UNA NOMINA SIGNIFICA QUE EXISTEN NOMINAS ACUMULADAS PARA EL MES
                               '///////////////////nomina cartorcenal y quincenal
                               If Me.DtaConsulta.Recordset.RecordCount >= 1 Then
                               
                               
                             
                             
                                     pPeriodo = Me.DtaConsulta.Recordset("Periodo")
                                     
                                     'Periodo Actual
                                     'MontoComisiones  no sumo comisiones, guardo viaticos
'                                         MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'                                         MontoBrutoMensual = (MontoBruto * 26) / 12
                                         
        '                                 Me.DtaConsulta.RecordSource = "SELECT  SUM(DetalleNomina.SalarioBasico +  DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas  + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasTurno + DetalleNomina.Antiguedad + DetalleNomina.Reembolso) AS TotalDevengado,    SUM(DetalleNomina.MontoINSS) AS MontoINSS, Nomina.NumNomina  FROM         DetalleNomina INNER JOIN    Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina  WHERE     (Nomina.Mes = " & Pmes & ") AND (Nomina.Ano = " & PAno & ") AND (Nomina.CodTipoNomina = '" & CodTipoNomina & "') AND (DetalleNomina.CodEmpleado = " & CodEmpleado & ") AND (NOT (Nomina.Periodo = " & pPeriodo & ")) GROUP BY Nomina.NumNomina"
                                         Me.DtaConsulta.RecordSource = "SELECT        SUM(DetalleNomina.SalarioBasico + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasTurno + DetalleNomina.Antiguedad + DetalleNomina.Reembolso) AS TotalDevengado, SUM(DetalleNomina.MontoINSS) AS MontoINSS, MAX(Nomina.NumNomina) AS NumNomina FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina WHERE (Nomina.Mes = " & Pmes & ") AND (Nomina.Ano = " & PAno & ") AND (Nomina.CodTipoNomina = '" & CodTipoNomina & "') AND (DetalleNomina.CodEmpleado = " & CodEmpleado & ") AND (NOT (Nomina.Periodo = " & pPeriodo & "))"
                                         Me.DtaConsulta.Refresh
                                          'DetalleNomina.Comisiones
                                        
                                        If Not Me.DtaConsulta.Recordset.EOF Then
                                                Dim Tdevenga, tInss As Double
                                                If IsNull(DtaConsulta.Recordset("TotalDevengado")) Then
                                                     Tdevenga = 0
                                                Else
                                                     Tdevenga = DtaConsulta.Recordset("TotalDevengado")
                                                     NumeroNominaAnt = Me.DtaConsulta.Recordset("NumNomina")
                                                     
                                                End If
                                                
                                                 If IsNull(DtaConsulta.Recordset("MontoINSS")) Then
                                                     tInss = 0
                                                Else
                                                     tInss = DtaConsulta.Recordset("MontoINSS")
                                                End If
                                                
                                                If IsNull(Me.DtaConsulta.Recordset("NumNomina")) Then
                                                  NumeroNominaAnt = -1
                                                Else
                                                NumeroNominaAnt = Me.DtaConsulta.Recordset("NumNomina")
                                                End If
                                        End If
                                        '/////////////////////////////////////////////////////////////////////////////////////
                                        '////////////////////////BUSCO LOS INCENTIVOS EXCENTEOS PARA RESTAR EL DEVENGADO ///
                                        '///////////////////////////////////////////////////////////////////////////////////
                                          '/////////////////////////////BUSCO LOS INCENTIVOS /////////////////////////////////////////////
'                                        Me.DtaConsulta.RecordSource = "SELECT  Nomina.* From Nomina WHERE (Mes = " & pMes & ") AND (Ano = " & pAno & ") AND (Periodo = " & pPeriodo & ") AND (CodTipoNomina = '" & CodTipoNomina & "')"
'                                        Me.DtaConsulta.Refresh
'                                        If Not Me.DtaConsulta.Recordset.EOF Then
'                                          NumeroNominaAnt = Me.DtaConsulta.Recordset("NumNomina")
'                                        End If
                                        
        
                                        MDIPrimero.AdoConsulta.ConnectionString = Conexion
                                        MDIPrimero.AdoConsulta.RecordSource = "SELECT MAX(DetalleIncentivo.NumIncentivo) AS NumIncentivo, SUM(DetalleIncentivo.Valor) AS Valor FROM DetalleIncentivo INNER JOIN Incentivo ON DetalleIncentivo.NumIncentivo = Incentivo.NumIncentivo INNER JOIN Empleado ON Incentivo.CodEmpleado = Empleado.CodEmpleado  " & _
                                                                              "WHERE (Incentivo.CodTipoIncentivo = '14') AND (Empleado.CodEmpleado = " & CodEmpleado & ") AND (DetalleIncentivo.NumNomina = " & NumeroNominaAnt & ")"
                                        MDIPrimero.AdoConsulta.Refresh
                                        If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
                                          If Not IsNull(MDIPrimero.AdoConsulta.Recordset("Valor")) Then
                                                Viaticos = Format(MDIPrimero.AdoConsulta.Recordset("Valor"), "##,##0.00")
                                          End If
                                        End If
                                                                       
                                        
                                        
                                         'MontoBrutoMensual = (MontoBruto + (Tdevenga - tInss) * 26) / 12
        '                                 MontoBrutoMensual = MontoBrutoMensual + (((Tdevenga - tInss) * 26) / 12)
                                          MontoNomina = (Tdevenga - tInss - Viaticos)
                                         
                                     Else
                                        MontoNomina = 0
                                         
                                         
                                     End If
 
                          End If
                        End If
                        
 
   Referencia = "Monto Nomina " & Format(MontoNomina, "##,##0.00") & " Monto INSS " & Format(tInss, "##,##0.00")

   MontoBrutoMensual = MontoNomina + MontoBaja


        DtaIR.Refresh
        DtaIR.Recordset.MoveNext
        MinIR = DtaIR.Recordset("desde")
        MinIR = MinIR - 1
        MinIR = (MinIR / 12)
     '   MsgBox MinIR
        Do While Not DtaIR.Recordset.EOF

           'ubicar la linea
         If DtaTipoNomina.Recordset("Periodo") = "Semanal Viernes" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIr = Format(MontoIr / CantSabados / 12, "###,##0.00")
               MontoIRPatronal = MontoIr
               Exit Do
            End If
            End If

         ElseIf DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIr = Format(MontoIr / CantSabados / 12, "###,##0.00")
               MontoIRPatronal = MontoIr
               Exit Do

            End If
            End If

        ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Viernes" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
  '///////Verfico si el la Ultima Quincena para hacer ajustes////////////

                MontoIrMensual = Format(MontoIr / 1 / 12, "###,##0.00")
                MontoIr = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior

            End If
            Else
               MontoIrMensual = 0
               MontoIr = MontoIrMensual - MontoIrAnterior
               MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
            End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
  '///////Verfico si el la Ultima Quincena para hacer ajustes////////////

                MontoIrMensual = Format(MontoIr / 1 / 12, "###,##0.00")
                MontoIr = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior

            End If
            Else
               MontoIrMensual = 0
                MontoIr = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
            End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'///////Verfico si el la Ultima Quincena para hacer ajustes////////////

                MontoIrMensual = Format(MontoIr / 1 / 12, "###,##0.00")
                MontoIr = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior

            End If
            Else
               MontoIrMensual = 0
                MontoIr = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
            End If

         ElseIf DtaTipoNomina.Recordset("Periodo") = "Mensual" Then
           If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIr = Format(MontoIr / 12, "###,##0.00")
               MontoIRPatronal = MontoIr
               Exit Do
            End If
         End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
           If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIr = Format(MontoIr / 4, "###,##0.00")
               MontoIRPatronal = MontoIr
               Exit Do
            End If
           End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
             If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIr = Format(MontoIr / 2, "###,##0.00")
               MontoIRPatronal = MontoIr
               Exit Do
            End If
            End If
         End If
        DtaIR.Recordset.MoveNext
        Loop
        
  CalcularMontoIrBajas = MontoIr

End Function



'Public Function CalcularMontoIr(CodTipoNomina As String, TotalDevengado As Double, CodEmpleado As String, TipoCalculoIr As String, IrUltimaSemana As Boolean, MontoInss As Double, NumNomina As String) As Double
'Dim i As Integer, sql As String, TarifaHorariaBasico As Double, FechaIni As String, TotalSalarioxHora As Double, TasaCambio As Double, SQLNominaEmpleado As String, TarifaHoraria As Double, SQLNomina As String, TotalHoras As Double, CodDepartamento As String, MontoIncentivoHoras As Double, PorcientoIncentivo As Double, TotalPuntualidad As Double, SQlIncentivos As String, Septimo As Double, SQlDeducciones As String, MontoIrAcumulado As Double, SQlPrestamo As String, SQlComisiones As String, SQlDestajo As String, SqlHrsExtras As String, SueldoPeriodo As Double, TasaInss As Double, TasaInssPatronal As Double, MontoIncentivos As Double, TasaIr As Double, MontoDeduccion As Double, MontoPrestamo As Double, MontoComisiones As Double, FechaContrato As Date, MontoDestajos As Double, annos As Date, MontoHRSExtras As Double, Antiguedad As Double
'Dim MontoOtrosIngresos As Double, PorcientoAntiguedad As Double, DescripOtrIngre As String, NumFecha1 As Date, MontoHora As Double, CodProceso As String, CantEmpleados As Long, CodReferencia As String, MontoIr As Double, UnidadesProducidas As Double, Rango As Double, Monto As Double, MontoIRPatronal As Double, NumeroDeduccion As Double, MontoInssPatronal As Double, MontoBrutoAnual As Double, MontoVacaciones As Double, Nombres As String, MontoMes13 As Double, FechaNomina As Date, DeduccionPorFalta As Double, SeptimoAnterior As Double, MinIR As Double, AñoFiscal As Double, SalarioMensual As Double, RentaGravable As Double, DiasMes As Double, TotalDevengadoAcumulado As Double, DiasSemana As Double, IncentivoProduccion As Double, CantSabados As Byte, IdDeduccion As Double, PagoProduccion As Double, SalHora As Double, NumeroPeriodo As Double, PeriodoFiscal As Double, Factor As Double, NQuincenas As Double, INATEC As Double, FechaInicialIr As Date
'Dim FechaFinalIr As Date, DevengadoSinHrsExtras As Double, VacacionesAcumuladas As Double, HE As Single, HoraPuntualidad As Double, MontoPuntualidad As Double, DD As Single, HoraSeptimo As Double, HoraBasico As Double, FormatoNomina As String, Adelantos As Double, Anos As Double, Moneda As String, MontoDolares As Double, MontoProduccion As Double, agregar As Boolean, FechaIngreso As Date, PeriodoIngreso As Double, BonoProduccion As Double, MontoViaticos As Double, NumIncentivo As Double, FechaInicio As String, FechaFin As String, Mes As Double, Fecha As String, Calcular7mo As Boolean, Dolarizado As Boolean, cn As New ADODB.Connection, ValorPunto As Double, SalarioMinimo As Double, SalarioPorciento As Double, rs As New ADODB.Recordset, TotalPuntos As Double, SalarioBasico As Double, CalcularPuntos As Boolean, MontoInssBasico As Double, AjusteINSS As Double, cmd As New ADODB.Command, HT As Double, MontoHorasTurno As Double
'Dim TipoVacaciones As Boolean, MontoTipoVacaciones As Double, CalcularHorasTurno As Boolean
'Dim AnoIni As Double, Viaticos As Double, NumeroNominaAnt As Double
'
'AnoIni = Year(Me.DtaNomina.Recordset("FechaNomina"))
'Mes = (Me.DtaNomina.Recordset("Mes"))
'
'        '-------------------------------------------------------------------------------------------------------
'        '------------------------------BUSCO EL SALARIO BASICO DEL EMPLEADO-------------------------------------
'        '-------------------------------------------------------------------------------------------------------
'
'
'
'        Me.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, PorcientoInss, TasaInss, PorcientoIr, TasaIr From TipoNomina WHERE (PorcientoIr = 1) AND (CodTipoNomina = '" & CodTipoNomina & "' )"
'        Me.DtaConsulta.Refresh
'        If DtaConsulta.Recordset.EOF Then
'
'        End If
'
'        '//////////////////////////////////////////////////
'        '///PRIMERO BUSCO EL NUMERO DEL PERIODO PARA CALCULAR IR
'        '////////////////////////////////////////////////////////
'        '///////////////////////Verifico si Tiene Ir Porcentual//////////////////////////////
''        CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
'        Me.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, PorcientoInss, TasaInss, PorcientoIr, TasaIr From TipoNomina WHERE (PorcientoIr = 1) AND (CodTipoNomina = '" & CodTipoNomina & "' )"
'        Me.DtaConsulta.Refresh
'        If DtaConsulta.Recordset.EOF Then
'
'         Select Case DtaTipoNomina.Recordset("Periodo")
'              Case "Catorcenal los Sabados"
'                Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (CodTipoNomina = " & CodTipoNomina & ") AND (Inicio = '" & Format(CDate(Me.TxtFechaIni.Text), "DD/MM/YYYY") & "') ORDER BY Periodo"
'                Me.AdoPeriodoFiscal.Refresh
'                If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
'                   PeriodoFiscal = Me.AdoPeriodoFiscal.Recordset("Periodo") ' formula = n
'                   NumeroPeriodo = 24 - (PeriodoFiscal - 1) 'formula = 24-(n-1)
'                   AñoFiscal = Me.AdoPeriodoFiscal.Recordset("Año")
'                End If
'              Case "Quincenal"
'                Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (CodTipoNomina = " & CodTipoNomina & ") AND (Inicio = '" & Format(CDate(Me.TxtFechaIni.Text), "DD/MM/YYYY") & "') ORDER BY Periodo"
'                Me.AdoPeriodoFiscal.Refresh
'                If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
'                   PeriodoFiscal = Me.AdoPeriodoFiscal.Recordset("Periodo") ' formula = n
'                   NumeroPeriodo = 24 - (PeriodoFiscal - 1) 'formula = 24-(n-1)
'                   AñoFiscal = Me.AdoPeriodoFiscal.Recordset("Año")
'                End If
'
'               Case "Mensual"
'                Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (CodTipoNomina = " & CodTipoNomina & ") AND (Inicio = '" & Format(CDate(Me.TxtFechaIni.Text), "DD/MM/YYYY") & "') ORDER BY Periodo"
'                Me.AdoPeriodoFiscal.Refresh
'                If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
'                   PeriodoFiscal = Me.AdoPeriodoFiscal.Recordset("Periodo") ' formula = n
'                   NumeroPeriodo = 12 - (PeriodoFiscal - 1) 'formula = 12-(n-1)
'                   AñoFiscal = Me.AdoPeriodoFiscal.Recordset("Año")
'                End If
'
'
'          End Select
'        End If
'
'        '//////////////////////////////////////////////////
'        '///BUSCO LA FECHA INICIAL DEL AÑO FISCAL
'        '////////////////////////////////////////////////////////
'                Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (Año = " & AñoFiscal & ") AND (CodTipoNomina = " & CodTipoNomina & ") AND (Periodo = 1)ORDER BY Periodo"
'                Me.AdoPeriodoFiscal.Refresh
'                If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
'                  FechaInicialIr = Me.AdoPeriodoFiscal.Recordset("Inicio")
'                End If
'
'        '//////////////////////////////////////////////////
'        '///BUSCO LA FECHA DE LA ULTIMA NOMINA DEL AÑO FISCAL CALCULADA
'        '////////////////////////////////////////////////////////
'                PeriodoFiscal = PeriodoFiscal - 1
'                Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (Año = " & AñoFiscal & ") AND (CodTipoNomina = " & CodTipoNomina & ") AND (Periodo = " & PeriodoFiscal & ") ORDER BY Periodo"
'                Me.AdoPeriodoFiscal.Refresh
'                PeriodoFiscal = PeriodoFiscal + 1
'                If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
'                  FechaFinalIr = Me.AdoPeriodoFiscal.Recordset("Final")
'                End If
'
'        '/////////////////////////////////////////////////////////////////
'        '///BUSCO LAS NOMINAS ACUMULADAS/////////////////////////////////
'        '///////////////////////////////////////////////////////////////////
'
'        sql = "SELECT     DetalleNomina.CodEmpleado AS CodEmpleado, SUM(DetalleNomina.MontoINSS) AS MontoINSS, SUM(DetalleNomina.MontoIR) AS MontoIR, " & _
'             "SUM(DetalleNomina.VacacionesPagadas) AS Vacaciones, SUM(DetalleNomina.INSSPatronal) AS INSSPatronal, SUM(DetalleNomina.IRPatronal) AS IRPatronal, " & _
'             "SUM(DetalleNomina.INATEC) AS INATEC, " & _
'             "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos " & _
'             " + DetalleNomina.VacacionesPagadas + DetalleNomina.AdelantosVacaciones) AS TotalDevengado, COUNT(DetalleNomina.NumNomina) AS NQuincenas, MIN(Nomina.FechaNominaINI) AS FechaIngreso FROM DetalleNomina INNER JOIN " & _
'             "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina " & _
'             "WHERE     (Nomina.FechaNomina <= CONVERT(DATETIME, '" & Format(FechaFinalIr, "yyyy/mm/dd") & "', 102)) AND (Nomina.FechaNominaINI >= CONVERT(DATETIME, '" & Format(FechaInicialIr, "yyyy/mm/dd") & "', 102)) " & _
'             "GROUP BY DetalleNomina.CodEmpleado " & _
'             "Having (DetalleNomina.CodEmpleado = " & CodEmpleado & ") "
'
'        '+ DetalleNomina.Incentivos   PANAM LO QUITE
'
'            Me.DtaConsulta.RecordSource = sql
'            Me.DtaConsulta.Refresh
'            TotalDevengadoAcumulado = 0
'            MontoIrAcumulado = 0
'            VacacionesPagadas = 0
'            If Not Me.DtaConsulta.Recordset.EOF Then
'               MontoIrAcumulado = Me.DtaConsulta.Recordset("MontoIR")
'               TotalDevengadoAcumulado = Me.DtaConsulta.Recordset("TotalDevengado") - Me.DtaConsulta.Recordset("MontoINSS")
'               NQuincenas = Me.DtaConsulta.Recordset("NQuincenas") + 1
'               FechaIngreso = Me.DtaConsulta.Recordset("FechaIngreso")
'               VacacionesAcumuladas = Me.DtaConsulta.Recordset("Vacaciones")
'            Else
'               NQuincenas = 1
'               FechaIngreso = Me.TxtFechaIni.Text
'
'            End If
'
'
'        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'        '/////////////////////////////////////BUSCO SI EXISTE NOMINA ACUMULADA //////////////////////////////////////////////
'        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'         sql = "SELECT id, NumNomina, CodEmpleado, SalarioBasico, Destajo, HE, DD, HorasExtras, Comisiones, OtrosIngresos, DescripOtrIngre, Incentivos, Deducciones, Prestamo, MontoINSS, MontoIR, Vacaciones, INSSPatronal, IRPatronal, INATEC, Mes13, DiasDescuento, Adelantos, TotalSubsidio, VacacionesPagadas, DiasVacaciones, AdelantosVacaciones, HTrabajada, SeptimoDia, IncetivoProduccion, TarifaHoraria, produjo, BonoProduccion, Viaticos, Ajuste, TIngresos, TGastos, SalarioBasico + Destajo + HorasExtras + Comisiones + OtrosIngresos + Incentivos + SeptimoDia + IncetivoProduccion + BonoProduccion AS TotalDevengado,NQuincenaAcumulada From DetalleNominaAcumulada Where (NumNomina = 0) And (CodEmpleado = " & CodEmpleado & ")"
'         Me.DtaConsulta.RecordSource = sql
'         Me.DtaConsulta.Refresh
'             If Not Me.DtaConsulta.Recordset.EOF Then
'               MontoIrAcumulado = Me.DtaConsulta.Recordset("MontoIR") + MontoIrAcumulado
'               TotalDevengadoAcumulado = TotalDevengadoAcumulado + Me.DtaConsulta.Recordset("TotalDevengado") - Me.DtaConsulta.Recordset("MontoINSS")
'               VacacionesAcumuladas = Me.DtaConsulta.Recordset("Vacaciones")
'               If Not IsNull(Me.DtaConsulta.Recordset("NQuincenaAcumulada")) Then
'                 NQuincenas = Me.DtaConsulta.Recordset("NQuincenaAcumulada") + NQuincenas
'               End If
'             End If
'
'
'
'
'        '//////////////////////////////////////////////////
'        '///BUSCO EL PERIODO DE INGRESO DEL EMPLEADO
'        '////////////////////////////////////////////////////////
'                Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (CodTipoNomina = " & CodTipoNomina & ") AND (Inicio = '" & Format(CDate(FechaIngreso), "DD/MM/YYYY") & "') ORDER BY Periodo"
'                Me.AdoPeriodoFiscal.Refresh
'                If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
'                   PeriodoIngreso = Me.AdoPeriodoFiscal.Recordset("Periodo")
'                End If
'
'
'
'        '///////////////////////Verifico si Tiene Ir Porcentual//////////////////////////////
''        CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
'        Me.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, PorcientoInss, TasaInss, PorcientoIr, TasaIr From TipoNomina WHERE (PorcientoIr = 1) AND (CodTipoNomina = '" & CodTipoNomina & "' )"
'        Me.DtaConsulta.Refresh
'        If DtaConsulta.Recordset.EOF Then
'         'Hago el Calcul del nuevo Techo para el Ir
'         Select Case DtaTipoNomina.Recordset("Periodo")
'                        Case "Semanal Viernes"
'
'                            If BuscaUltimaSemana(CDbl(CantSabados), CDbl(NumNomina), Format(Mes, "0#"), CDbl(AnoIni)) = True Then
'                             MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'                             MontoBrutoMensual = MontoBruto + TotalSueldoAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes)) - TotalInssAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes))
'                            ElseIf IrUltimaSemana = False Then
'                                MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'                                MontoBrutoMensual = MontoBruto * CantSabados
'                            Else
'                                MontoBrutoMensual = 0
'                            End If
'
'                        Case "Semanal Sabado"
'
'                            If BuscaUltimaSemana(CDbl(CantSabados), CDbl(NumNomina), Format(Mes, "0#"), CDbl(AnoIni)) = True Then
'                             MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'                             MontoBrutoMensual = MontoBruto + TotalSueldoAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes)) - TotalInssAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes))
'                            ElseIf IrUltimaSemana = False Then
'                                MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'                                MontoBrutoMensual = MontoBruto * CantSabados
'                            Else
'                                MontoBrutoMensual = 0
'                            End If
'
'                        Case "Catorcenal los Viernes"
'                            If DiaFin < 28 Then
'                             MontoBruto = (TotalDevengado + MontoOtrosIngresos + MontoTipoVacaciones) - MontoInss
'                             MontoBrutoMensual = ((MontoBruto * 15) / 14) * 2
'                            Else
'                             MontoBrutoMensual = SalarioMensual - MontoInssMensual
'                            End If
'                        Case "Catorcenal los Sabados"
'                        'EMPIEZO A BUSCAR SI EN EL PERIODO EN EL QUE ESTOY ES LA ULTIMA SEMANA, SI LO ES ENTONCES CALCULO
'                        'SI NO EXISTEN FILAS/ROWS ENTONCES SE CALCULA IR
'                        'SELECT     Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina   FROM         Fecha_Planilla  WHERE     (año = 2017) AND (mes = N'04') AND (CodTipoNomina = N'04') AND (Inicio > CONVERT(DATETIME, '04-17-2017 00:00:00', 102))
'
'
'
'                        'CodTipoNomina
'                        'DtaNomina.Recordset ("FechaNomina")
'                        'Mes = (DtaNomina.Recordset("Mes"))
'
'
'                        If IrUltimaSemana = False Then
'                              MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'                              MontoBrutoMensual = (MontoBruto * 26) / 12
'                        Else
'
'                             Me.DtaConsulta.RecordSource = "SELECT     Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina   FROM         Fecha_Planilla  WHERE     (año = " & PAno & ") AND (mes ='" & Format(Pmes, "00") & "') AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Inicio > CONVERT(DATETIME, '" & Format(PFechaNomina, "MM-dd-yyyy") & " 00:00:00', 102))"
'                             Me.DtaConsulta.Refresh
'                             If Me.DtaConsulta.Recordset.EOF Then
'
'                             Dim pPeriodo As Integer
'                             pPeriodo = Periodo
'
'                             'Periodo Actual
'                             'MontoComisiones  no sumo comisiones, guardo viaticos
'                                 MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'                                 MontoBrutoMensual = (MontoBruto * 26) / 12
'
''                                 Me.DtaConsulta.RecordSource = "SELECT  SUM(DetalleNomina.SalarioBasico +  DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas  + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasTurno + DetalleNomina.Antiguedad + DetalleNomina.Reembolso) AS TotalDevengado,    SUM(DetalleNomina.MontoINSS) AS MontoINSS, Nomina.NumNomina  FROM         DetalleNomina INNER JOIN    Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina  WHERE     (Nomina.Mes = " & Pmes & ") AND (Nomina.Ano = " & PAno & ") AND (Nomina.CodTipoNomina = '" & CodTipoNomina & "') AND (DetalleNomina.CodEmpleado = " & CodEmpleado & ") AND (NOT (Nomina.Periodo = " & pPeriodo & ")) GROUP BY Nomina.NumNomina"
'                                 Me.DtaConsulta.RecordSource = "SELECT        SUM(DetalleNomina.SalarioBasico + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasTurno + DetalleNomina.Antiguedad + DetalleNomina.Reembolso) AS TotalDevengado, SUM(DetalleNomina.MontoINSS) AS MontoINSS, MAX(Nomina.NumNomina) AS NumNomina FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina WHERE (Nomina.Mes = " & Pmes & ") AND (Nomina.Ano = " & PAno & ") AND (Nomina.CodTipoNomina = '" & CodTipoNomina & "') AND (DetalleNomina.CodEmpleado = " & CodEmpleado & ") AND (NOT (Nomina.Periodo = " & pPeriodo & "))"
'                                 Me.DtaConsulta.Refresh
'                                  'DetalleNomina.Comisiones
'
'                                If Not Me.DtaConsulta.Recordset.EOF Then
'                                        Dim Tdevenga, tInss As Double
'                                        If IsNull(DtaConsulta.Recordset("TotalDevengado")) Then
'                                             Tdevenga = 0
'                                        Else
'                                             Tdevenga = DtaConsulta.Recordset("TotalDevengado")
'                                             NumeroNominaAnt = Me.DtaConsulta.Recordset("NumNomina")
'
'                                        End If
'
'                                         If IsNull(DtaConsulta.Recordset("MontoINSS")) Then
'                                             tInss = 0
'                                        Else
'                                             tInss = DtaConsulta.Recordset("MontoINSS")
'                                        End If
'                                End If
'                                '/////////////////////////////////////////////////////////////////////////////////////
'                                '////////////////////////BUSCO LOS INCENTIVOS EXCENTEOS PARA RESTAR EL DEVENGADO ///
'                                '///////////////////////////////////////////////////////////////////////////////////
'                                  '/////////////////////////////BUSCO LOS INCENTIVOS /////////////////////////////////////////////
''                                Me.DtaConsulta.RecordSource = "SELECT  Nomina.* From Nomina WHERE (Mes = " & Pmes & ") AND (Ano = " & PAno & ") AND (Periodo = " & pPeriodo & ") AND (CodTipoNomina = '" & CodTipoNomina & "')"
''                                Me.DtaConsulta.Refresh
''                                If Not Me.DtaConsulta.Recordset.EOF Then
''                                  NumeroNominaAnt = Me.DtaConsulta.Recordset("NumNomina")
''                                End If
'
'
'                                MDIPrimero.AdoConsulta.ConnectionString = Conexion
'                                MDIPrimero.AdoConsulta.RecordSource = "SELECT MAX(DetalleIncentivo.NumIncentivo) AS NumIncentivo, SUM(DetalleIncentivo.Valor) AS Valor FROM DetalleIncentivo INNER JOIN Incentivo ON DetalleIncentivo.NumIncentivo = Incentivo.NumIncentivo INNER JOIN Empleado ON Incentivo.CodEmpleado = Empleado.CodEmpleado  " & _
'                                                                      "WHERE (Incentivo.CodTipoIncentivo = '14') AND (Empleado.CodEmpleado = " & CodEmpleado & ") AND (DetalleIncentivo.NumNomina = " & NumeroNominaAnt & ")"
'                                MDIPrimero.AdoConsulta.Refresh
'                                If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
'                                  If Not IsNull(MDIPrimero.AdoConsulta.Recordset("Valor")) Then
'                                        Viaticos = Format(MDIPrimero.AdoConsulta.Recordset("Valor"), "##,##0.00")
'                                  End If
'                                End If
'
'
'
'                                 'MontoBrutoMensual = (MontoBruto + (Tdevenga - tInss) * 26) / 12
''                                 MontoBrutoMensual = MontoBrutoMensual + (((Tdevenga - tInss) * 26) / 12)
'                                  MontoBrutoMensual = MontoBruto + (Tdevenga - tInss - Viaticos)
'
'                             Else
'                                 MontoBrutoMensual = 0
'
'
'                             End If
'
'
'                        End If
'
'
'
'
'
'
'
'                        Case "Quincenal"
'                          If TipoCalculoIr = "Calcular IR x 12" Then
'                            If DiaFin < 28 Then
'                              If IrUltimaSemana = False Then
'                                MontoBruto = (TotalDevengado) - MontoInss
'                                MontoBrutoMensual = MontoBruto * 2
'                                MontoBrutoAnual = MontoBrutoMensual * 12
'                              Else
'                                MontoBruto = 0
'                                MontoBrutoMensual = 0
'                                MontoBrutoAnual = 0
'                              End If
'                            ElseIf IrUltimaSemana = False Then
'                                MontoBruto = (TotalDevengado) - MontoInss
'                                MontoBrutoMensual = MontoBruto * 2
'                                MontoBrutoAnual = MontoBrutoMensual * 12
'                                '                        If TotalDevengadoAnterior = 0 Then
'        '                           MontoBrutoMensual = (SalarioMensual - MontoInssMensual) * 2
'        '                           MontoBrutoAnual = MontoBrutoMensual * 12
'        '                        Else
'        '                           MontoBrutoMensual = SalarioMensual - MontoInssMensual
'        '                           MontoBrutoAnual = MontoBrutoMensual * 12
'        '                        End If
'                            ElseIf IrUltimaSemana = True Then
'                             MontoBruto = (TotalDevengado) - MontoInss
'                             MontoBrutoMensual = MontoBruto + TotalSueldoAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes)) - TotalInssAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes))
'                             MontoBrutoAnual = MontoBrutoMensual * 12
'                            End If
'                          Else
'                           MontoBruto = (TotalDevengado) - MontoInss '+ MontoOtrosIngresos
'                           RentaGravable = ((TotalDevengadoAcumulado + MontoBruto) / NQuincenas) * 24
'
'                           MontoBrutoAnual = RentaGravable '+ MontoVacaciones + VacacionesAcumuladas
'                           MontoBrutoMensual = MontoBruto * 2
'                          End If
'
'                        Case "Mensual"
'
'                           MontoBruto = (TotalDevengado) - MontoInss
'                           RentaGravable = ((TotalDevengadoAcumulado + MontoBruto) * (12 - (PeriodoIngreso - 1))) / NQuincenas
'        '                   MontoBrutoAnual = RentaGravable + MontoVacaciones + VacacionesAcumuladas
'                           MontoBrutoMensual = MontoBruto
'                           MontoBrutoAnual = MontoBrutoMensual * 12
'        '                    MontoBruto = SalarioMensual - MontoInssMensual
'        '                    MontoBrutoMensual = MontoBruto
'                        Case "Trimestral"
'
'                            MontoBruto = SalarioMensual - MontoInssMensual
'                            MontoBrutoMensual = MontoBruto / 3
'                        Case "Semestral"
'
'                            MontoBruto = SalarioMensual - MontoInssMensual
'                            MontoBrutoMensual = MontoBruto / 6
'        End Select
'
'
'          '//////////////////////////////////////////////////////////////////////////
'          '///////////////////BUSCO EL TIPO DE MONEDA DE LA NOMINA///////////////////
'          '//////////////////////////////////////////////////////////////////////////
'           Me.AdoBusca.RecordSource = "SELECT CodTipoNomina, Nomina, Periodo, UltFecha, TipoPago, Moneda, MantValor, Activa, PorcientoInss, TasaInss, PorcientoIr, TasaIr,TasaInssPatronal From TipoNomina WHERE (CodTipoNomina = '" & CodTipoNomina & "')"
'           Me.AdoBusca.Refresh
'           If Not Me.AdoBusca.Recordset.EOF Then
'              Moneda = Me.AdoBusca.Recordset("Moneda")
'           Else
'              Moneda = "C$"
'           End If
'
'        If DtaEmpleados.Recordset("ExentoIr") = False Then
'                'agregar IR laboral y patronal
'
'                MontoIr = 0
'                MontoIRPatronal = 0
'                MontoDolares = 0
'                If Moneda = "US" Then
'                 MontoDolares = MontoBrutoMensual
'                 MontoBrutoMensual = MontoBrutoMensual * TasaCambio
'                End If
'
'
'                DtaIR.Refresh
'                DtaIR.Recordset.MoveNext
'                MinIR = DtaIR.Recordset("desde")
'                MinIR = MinIR - 1
'                MinIR = (MinIR / 12)
'                Do While Not DtaIR.Recordset.EOF
'
'                   'ubicar la linea
'                 If DtaTipoNomina.Recordset("Periodo") = "Semanal Viernes" Then
'                    If (MontoBrutoMensual) >= MinIR Then
'                    If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'                       MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'                       MontoIr = Format(MontoIr / 12, "###,##0.00")  'MontoIr = Format(MontoIr / CantSabados / 12, "###,##0.00")
'                       MontoIRPatronal = MontoIr
'                       Exit Do
'                    End If
'                    End If
'
'                 ElseIf DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
'                    If (MontoBrutoMensual) >= MinIR Then
'                    If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'                       MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'                       MontoIr = Format(MontoIr / 12, "###,##0.00")
'                       MontoIRPatronal = MontoIr
'                       Exit Do
'
'                    End If
'                    End If
'
'                ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Viernes" Then
'                    If (MontoBrutoMensual) >= MinIR Then
'                    If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'                       MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'          '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
'                       If DiaFin < 28 Then
'                        MontoIr = Format(MontoIr / 2 / 12, "###,##0.00")
'                        MontoIRPatronal = MontoIr
'                        Exit Do
'                       Else
'                        MontoIrMensual = Format(MontoIr / 1 / 12, "###,##0.00")
'                        MontoIr = MontoIrMensual - MontoIrAnterior
'                        MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
'                       End If
'                    End If
'                    Else
'                       MontoIrMensual = 0
'                       MontoIr = MontoIrMensual - MontoIrAnterior
'                       MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
'                    End If
'                 ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
'                    If (MontoBrutoMensual) >= MinIR Then
'                    If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'                       MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'          '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
'                       If DiaFin < 20 Then
'                            If IrUltimaSemana = False Then
'                                MontoIr = Format(MontoIr / 26, "###,##0.00")
'                                MontoIRPatronal = MontoIr
'                            Else
'                             MontoIr = Format(MontoIr / 12, "###,##0.00")
'                             MontoIRPatronal = MontoIr
'                            End If
'                       Else
'                            If IrUltimaSemana = False Then
'                                MontoIr = Format(MontoIr / 26, "###,##0.00")
'                                MontoIRPatronal = MontoIr
'                            Else
'                             MontoIr = Format(MontoIr / 12, "###,##0.00")
'                             MontoIRPatronal = MontoIr
'                            End If
'                       End If
'                    End If
'                    Else
'                       MontoIrMensual = 0
'                        MontoIr = MontoIrMensual - MontoIrAnterior
'                        MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
'                    End If
'
'
'                 ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
'                     If DtaIR.Recordset("desde") <= (MontoBrutoAnual) And DtaIR.Recordset("Hasta") >= (MontoBrutoAnual) Then
'                       MontoIr = ((MontoBrutoAnual) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'        '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
'
'                        If TipoCalculoIr = "Calcular IR x 12" Then
'                            If DiaFin < 28 Then
'                                MontoIr = Format(MontoIr / 2 / 12, "###,##0.00")
'                                MontoIRPatronal = MontoIr
'                                Exit Do
'                            ElseIf IrUltimaSemana = False Then
''                                MontoIrAcumulado = TotalIrAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes))
''                                MontoIr = Format((MontoIr / 12) - MontoIrAcumulado, "###,##0.00")
''                                MontoIRPatronal = MontoIr
'                                MontoIr = Format(MontoIr / 24, "###,##0.00")
'                                MontoIRPatronal = MontoIr
'                            ElseIf IrUltimaSemana = True Then
'                                MontoIrAcumulado = TotalIrAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes))
'                                MontoIr = Format((MontoIr / 12) - MontoIrAcumulado, "###,##0.00")
'                                MontoIRPatronal = MontoIr
'                            End If
'                        Else
'                        If Not NumeroPeriodo = 0 Then
'                          'NumeroPeriodo = 24-(NQuincenas-1)
'                         MontoIr = (MontoIr - MontoIrAcumulado) / NumeroPeriodo
'                         ' MontoIr = ((MontoIr / 24) * NQuincenas) - MontoIrAcumulado
'                        Else
'                         MontoIr = 0
'                        End If
'                        End If
'
'                        MontoIRPatronal = MontoIr - MontoIrPatronalAnterior
'                        Exit Do
'        '               End If
'                     End If
'        '            Else
'        '               MontoIrMensual = 0
'
'        '                MontoIR = MontoIrMensual - MontoIrAnterior
'        '                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
'        '            End If
'
'
'
'                 ElseIf DtaTipoNomina.Recordset("Periodo") = "Mensual" Then
'        '           If (MontoBrutoAnual) >= MinIR Then
'                    If DtaIR.Recordset("desde") <= (MontoBrutoAnual) And DtaIR.Recordset("Hasta") >= (MontoBrutoAnual) Then
'
'                       MontoIr = ((MontoBrutoAnual) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'
'                        MontoIr = (MontoIr - MontoIrAcumulado) / 12
'                        MontoIRPatronal = MontoIr - MontoIrPatronalAnterior
'                        Exit Do
'
'                       Exit Do
'                    End If
'        '         End If
'                 ElseIf DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
'                   If (MontoBrutoMensual) >= MinIR Then
'                    If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'                       MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'                       MontoIr = Format(MontoIr / 4, "###,##0.00")
'                       MontoIRPatronal = MontoIr
'                       Exit Do
'                    End If
'                   End If
'                 ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
'                     If (MontoBrutoMensual) >= MinIR Then
'                    If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'                       MontoIr = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'                       MontoIr = Format(MontoIr / 2, "###,##0.00")
'                       MontoIRPatronal = MontoIr
'                       Exit Do
'                    End If
'                    End If
'                 End If
'          DtaIR.Recordset.MoveNext
'          Loop
'
'            If Moneda = "US" Then
'               MontoBrutoMensual = MontoDolares
'               If TasaCambio <> 0 Then
'                MontoIr = MontoIr / TasaCambio
'                MontoIRPatronal = MontoIRPatronal / TasaCambio
'               End If
'            End If
'
'          End If 'del if que pregunta si esta excento de IR
'                'TotalDevengado = TotalDevengado + MontoDestajo + MontoHRSExtras + MontoComisiones + MontoIncentivos
'        Else
'
'
'
'        End If
'
'
'CalcularMontoIr = MontoIr
'
'
'End Function





Private Sub Form_Load()
Me.TxtUltFechaNomina.Value = Format(Now, "dd/mm/yyyy")
Me.TxtFechaHistorial.Value = Format(Now, "dd/mm/yyyy")
Me.CmdDetalle.BackColor = RGB(219, 226, 242)
Me.CmdImprimirHistorial.BackColor = RGB(219, 226, 242)
MDIPrimero.Skin1.ApplySkin hWnd
 Me.TDBGridSalarios.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.TDBGridSalarios.OddRowStyle.BackColor = &H80000005
 Me.TDBGridSalarios.AlternatingRowStyle = True
 
  Me.TDBGridBonos.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.TDBGridBonos.OddRowStyle.BackColor = &H80000005
 Me.TDBGridBonos.AlternatingRowStyle = True

MDIPrimero.DtaControles.Refresh
If Not IsNull(MDIPrimero.DtaControles.Recordset("SalarioPromedioReal")) Then
   If MDIPrimero.DtaControles.Recordset("SalarioPromedioReal") = True Then
     Me.Check1.Value = 1
   Else
     Me.Check1.Value = 0
   End If
End If

With Me.DtaTurnos
   .ConnectionString = Conexion
   .RecordSource = "Turno"
   .Refresh
End With
 
 With DtaHorarioEmpleado
     .ConnectionString = Conexion
 End With
 
  With AdoAuxiliar
     .ConnectionString = Conexion
 End With
 
 With Me.AdoHistoricos
     .ConnectionString = Conexion
 End With
 
 With Me.AdoTraslado
     .ConnectionString = Conexion
 End With
 
 With Me.AdoEmpresa
   .ConnectionString = Conexion
   .RecordSource = "DatosEmpresa"
   .Refresh
End With

With Me.DtaIR
  '.DatabaseName = Ruta
  .ConnectionString = Conexion
  .RecordSource = "IR"
  .Refresh
End With

With Me.DtaInss
  '.DatabaseName = Ruta
  .ConnectionString = Conexion
  .RecordSource = "INSS"
  .Refresh
End With

With Me.AdoInicioAño
  '.DatabaseName = Ruta
  .ConnectionString = Conexion
End With

With Me.AdoElimina
  '.DatabaseName = Ruta
  .ConnectionString = Conexion
End With

With Me.AdoSalarios
  .ConnectionString = Conexion
End With


With Me.AdoAntiguedad
  '.DatabaseName = Ruta
  .ConnectionString = Conexion
End With

With Me.DtaHrsExtras
  '.DatabaseName = Ruta
  .ConnectionString = Conexion
End With

With Me.DtaConsulta
  '.DatabaseName = Ruta
  .ConnectionString = Conexion
End With

With Me.DtaAdelanto
  '.DatabaseName = Ruta
  .ConnectionString = Conexion
End With

With Me.DtaControles
  '.DatabaseName = Ruta
  .ConnectionString = Conexion
  .RecordSource = "Controles"
  .Refresh
End With

With Me.DtaBajas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaDeduccion
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaDeducciones
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaEmpleado
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With



With Me.adoEmpleado
      .ConnectionString = Conexion
      .RecordSource = "SELECT  * From Empleado"
      .Refresh
End With

With Me.DtaHistorico
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaNominas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.Dtaprestamo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With


With Me.DtaTipoNomina
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "TipoNomina"
   .Refresh
End With


End Sub



Private Sub txtCodEmpleado_Change()
On Error GoTo TipoErr
Dim FechaContrato As Date, FechaInicio As Date, FechaEgreso As Date
Dim FechaHoy As Date, SueldoPeriodo As Double, FechaHistorico As Date
Dim FechaUltNomina As Date, i As Integer, NumeroEmpleado As Double
Dim annos As Date
Dim SQlEmpleado As String
Dim SQlPrestamo As String
Dim SQlDeducciones As String
Dim SqlNominas As String, DiasMes As Double, DiasReales As Double, MesReal As Double
Dim FechaBusqueda As Date, Año As Integer, Mes As Integer
Dim Contador As Integer, TotalSalario As Double, Salario As Double, SalarioAlto As Double, SalarioPromedio As Double
Dim SueldoActual As Double, Dia As Double, Meses As Double

'SQlEmpleado = "SELECT  Empleado.SalarioFijo, Empleado.SueldoPeriodo, Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion, Empleado.Direccion AS Expr1, Empleado.Sexo, Empleado.Activo, Empleado.TarifaHoraria, Empleado.SueldoActualBasico FROM Departamento INNER JOIN Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento WHERE  (Empleado.CodEmpleado = " & TxtCodEmpleado.Text & ")"
SQlEmpleado = "SELECT Empleado.SalarioFijo, Empleado.SueldoPeriodo, Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion, Empleado.Direccion AS Expr1, Empleado.Sexo, Empleado.Activo, Empleado.TarifaHoraria, Empleado.SueldoActualBasico, Historico.SueldoActual FROM Departamento INNER JOIN Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado Where (Empleado.CodEmpleado = " & TxtCodEmpleado.Text & ")"
DtaEmpleado.RecordSource = SQlEmpleado
DtaEmpleado.Refresh

SQlPrestamo = "SELECT Prestamo.NumPrestamo, Prestamo.CodEmpleado, Prestamo.Monto, Prestamo.CantCuotas, Prestamo.Interes, Prestamo.Saldo, Prestamo.FechaInicial, Prestamo.Cancelado From Prestamo WHERE Prestamo.Cancelado=0 AND Prestamo.CodEmpleado='" & TxtCodEmpleado.Text & "'"
Dtaprestamo.RecordSource = SQlPrestamo
Dtaprestamo.Refresh

SQlDeducciones = "SELECT Deduccion.NumDeduccion, Deduccion.CodEmpleado, Deduccion.CodTipoDeduccion, DetalleDeduccion.NumDeduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.pagado  FROM Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion WHERE DetalleDeduccion.pagado=0 AND Deduccion.CodEmpleado='" & TxtCodEmpleado.Text & "'"
DtaDeducciones.RecordSource = SQlDeducciones
DtaDeducciones.Refresh

DoEvents

If Not DtaEmpleado.Recordset.EOF Then
       If DtaEmpleado.Recordset("activo") = False Then
           MsgBox "Este empleado ya fue dado de Baja"
           Exit Sub
        End If
        
    Me.txtCodEmpleado1.Text = DtaEmpleado.Recordset("CodEmpleado1")
    CodEmpleado = DtaEmpleado.Recordset("CodEmpleado")
    TxtNombre1 = DtaEmpleado.Recordset("Nombre1")
    TxtNombre2 = DtaEmpleado.Recordset("Nombre2")
    TxtApellido1 = DtaEmpleado.Recordset("Apellido1")
    TxtApellido2 = DtaEmpleado.Recordset("Apellido2")
    TxtDireccion = DtaEmpleado.Recordset("Direccion")
    txtCargo = DtaEmpleado.Recordset("Cargo")
    txtDepartamento = DtaEmpleado.Recordset("departamento")
    txtSexo = DtaEmpleado.Recordset("sexo")
    SalarioBasico = DtaEmpleado.Recordset("SueldoPeriodo")
    If DtaEmpleado.Recordset("SalarioFijo") = "S" Then
     SueldoFijo = True
    Else
     SueldoFijo = False
    End If
    
        If Not IsNull(DtaEmpleado.Recordset("SueldoActualBasico")) = True Then
         If DtaEmpleado.Recordset("SueldoActualBasico") = True Then
          Me.ChkSueldoActual.Value = 1
          If Not IsNull(DtaEmpleado.Recordset("SueldoActual")) Then
             SueldoActual = DtaEmpleado.Recordset("SueldoActual")
             SSueldoActual = DtaEmpleado.Recordset("SueldoActual")
          End If
        Else
          Me.ChkSueldoActual.Value = 0
          SueldoActual = 0
          SSueldoActual = 0
         End If
        End If
Else
    TxtNombre1 = ""
    TxtNombre2 = ""
    TxtApellido1 = ""
    TxtApellido2 = ""
    TxtDireccion = ""
    txtCargo = ""
    txtDepartamento = ""
    txtSexo = ""
    TxtAnnos.Text = ""
    TxtMeses.Text = ""
    TxtFechaContrato.Text = ""
    TxtUltFechaNomina.Value = Now
    TxtMotivo.Text = ""
   Exit Sub
End If

DiasMes = 0

DtaHistorico.RecordSource = "SELECT Id, Codempleado, FechaBaja, MotivoBaja, FechaAumento, MotivoAumento, FechaInicialSusp, FechaFinalSusp, MotivoSuspencion, FechaNacimiento,FechaContrato , CargoInicial, CargoActual, CargoAnterior, SueldoInicial, SueldoAnterior, SueldoActual, CuentaDebito, CuentaCredito From Historico"
DtaHistorico.Refresh
Do While Not DtaHistorico.Recordset.EOF
    If TxtCodEmpleado.Text = DtaHistorico.Recordset("CodEmpleado") Then
        If Not IsNull(DtaHistorico.Recordset("FechaContrato")) Then
           TxtFechaContrato.Text = DtaHistorico.Recordset("FechaContrato")
           FechaContrato = DtaHistorico.Recordset("FechaContrato")
        Else
        MsgBox "No ha sido grabada la fecha de contrato no puede realizarse la baja"
        CmdEfectuar.Enabled = False
        Exit Sub
        End If
           Exit Do
        
    End If
DtaHistorico.Recordset.MoveNext
Loop
FechaHoy = Format(Now, "dd/mm/yyyy")

SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[HorasExtras]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos] AS TotalBruto, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[HorasExtras]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]-[Deducciones]-[Prestamo]-[MontoINSS]-[MontoIR] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
DtaNominas.RecordSource = SqlNominas
DtaNominas.Refresh

If Not DtaNominas.Recordset.EOF Then
   DtaNominas.Recordset.MoveLast
   FechaUltNomina = DtaNominas.Recordset("FechaNomina")
   TxtUltFechaNomina.Value = FechaUltNomina
   NumFecha1 = FechaUltNomina
Else
   TxtUltFechaNomina = TxtFechaContrato
   'MsgBox "No ha sido Grabada Ninguna nómina a este empleado, Se le realizará la baja desde su contrato"
End If







'///////////Busco la Fecha para la Busqueda////////////////////////////

NumeroEmpleado = Me.TxtCodEmpleado.Text

If Me.ChkSueldoActual.Value = xtpUnchecked Then
SqlSalarios = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo,SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO,SUM(DetalleNomina.Comisiones) As Comisiones " & _
              "FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
              "Having (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.Comisiones) <> 0) And (DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") " & _
              "ORDER BY Nomina.Ano, Nomina.Mes "
Else
SqlSalarios = "SELECT DISTINCT TOP (100) PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SeptimoDia) AS Septimos, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,  SUM(DetalleNomina.Destajo + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones + DetalleNomina.Incentivos) + AVG(Historico.SueldoActual) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Comisiones) AS Comisiones, AVG(Historico.SueldoActual) AS SalarioBasico FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
              "Having (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.Comisiones) <> 0) And (DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") ORDER BY AÑO, Nomina.Mes"
End If



Me.DtaConsulta.RecordSource = SqlSalarios
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
 Me.DtaConsulta.Recordset.MoveLast
Else
 FechaHistorico = Format(Now, "dd/mm/yyyy")
 FechaBusqueda = Format(Now, "dd/mm/yyyy")
End If
i = 0
Do While Not Me.DtaConsulta.Recordset.BOF
  If i = 1 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")

  ElseIf i = 5 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    Exit Do
  ElseIf i = 0 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    FechaHistorico = Me.DtaConsulta.Recordset("FechaFin")
  Else
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
  End If
  i = i + 1

  Me.DtaConsulta.Recordset.MovePrevious
Loop


FechaEgreso = Me.TxtUltFechaNomina.Value

FechaContrato = Me.TxtFechaContrato.Text

'//////////SUMO 1 PARA AJUSTAR QUE SIEMPRE DA 1 DIA MENOS//////
'annos = CDbl(FechaEgreso) - CDbl(FechaContrato) + 1
MDIPrimero.DtaControles.Refresh
DiasMes = MDIPrimero.DtaControles.Recordset("DiasMes")

'annos = CalcularDiasVaca(FechaContrato, FechaEgreso)
If DateDiff("d", FechaContrato, FechaEgreso) + 1 <= 30 Then
  Dias = DateDiff("d", FechaContrato, FechaEgreso) + 1
Else
  Dias = Format(CalcularDiasAntiguedad(FechaContrato, FechaEgreso) * 12, "####0")
End If
annos = Dias
TxtAnnos.Text = Format(annos / 360, "###,##0.00")

Meses = Dias / DiasMes
Mes = Int(Meses)
Dia = Format((Meses - Mes) * DiasMes, "####0")

TxtMeses.Text = Mes & " m y " & Dia & " d"

'TxtMeses.Text = Format(Dias / DiasMes, "###,##0")
'Dias = Format(annos * 365, "###,##0.00")
    Me.CmdEfectuar.Enabled = False
    Me.CmdRenovar.Enabled = False
    
DiasReales = CalcularDiasVaca(FechaBusqueda, FechaHistorico)
MesReal = DiasReales / DiasMes
'If MesReal > 6 Then
' MesReal = 6
'End If
    
'Me.TxtDiasTrabajados.Text = Format(annos, "###,##0")
'Me.TxtDiasTrabajados.Text = CalcularDiasVaca(FechaContrato, FechaEgreso)

Me.TxtDiasTrabajados.Text = Format(CalcularDiasAntiguedad(FechaContrato, FechaEgreso) * 12, "##,##0")
Me.DTPFechaIniAgui.Value = "01/12/" & Year(FechaBusqueda)

Año = Year(FechaBusqueda)
Mes = Month(FechaBusqueda)

If CDbl(Me.TxtDiasTrabajados.Text) < 14 Then
   Me.TxtDias.Text = Me.TxtDiasTrabajados.Text
End If



If Me.AdoEmpresa.Recordset("FormatoNomina") = "Nomina Bono Produccion" Then

   
    If Me.ChkExtra.Value = xtpUnchecked Then
    
       SqlSalarios = "SELECT DISTINCT " & _
                      "TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Antiguedad) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, " & _
                      "SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, " & _
                      "SUM(DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion) AS Incentivos, " & _
                      "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo  + DetalleNomina.BonoProduccion + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.HorasExtras + DetalleNomina.IncetivoProduccion + DetalleNomina.Antiguedad) AS TotalIngresos, " & _
                      "MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, " & _
                      "SUM(DetalleNomina.BonoProduccion) AS BonoProduccion, SUM(DetalleNomina.HorasExtras) AS HorasExtras " & _
                      "FROM         DetalleNomina INNER JOIN " & _
                      "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina " & _
                      "GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                      "HAVING      (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) " & _
                      "Between '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') " & _
                      "ORDER BY Nomina.Ano, Nomina.Mes "
               
               Me.TDBGridBonos.Visible = True
               Me.TDBGridSalarios.Visible = False
'               Me.TDBGridBonos.Columns(3).Visible = True
      Else
       SqlSalarios = "SELECT DISTINCT " & _
                      "TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Antiguedad) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, " & _
                      "SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, " & _
                      "SUM(DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion) AS Incentivos, " & _
                      "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo  + DetalleNomina.BonoProduccion + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.IncetivoProduccion + DetalleNomina.Antiguedad) AS TotalIngresos, " & _
                      "MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, " & _
                      "SUM(DetalleNomina.BonoProduccion) AS BonoProduccion " & _
                      "FROM         DetalleNomina INNER JOIN " & _
                      "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina " & _
                      "GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                      "HAVING      (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) " & _
                      "Between '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') " & _
                      "ORDER BY Nomina.Ano, Nomina.Mes "
               
               Me.TDBGridBonos.Visible = True
               Me.TDBGridSalarios.Visible = False
'               Me.TDBGridBonos.Columns(3).Visible = False
      
      
      End If
Else
  
  If Me.ChkSueldoActual.Value = xtpUnchecked Then
       ' SqlSalarios = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Antiguedad) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo,SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.Antiguedad)AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO,SUM(DetalleNomina.Comisiones) As Comisiones " &
                      '"FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                     ' "HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) Between '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') " & _
                       '"ORDER BY Nomina.Ano, Nomina.Mes "
      If Me.ChkIncentivos.Value = xtpUnchecked Then
            SqlSalarios = "SELECT DISTINCT"
            SqlSalarios = SqlSalarios + "  TOP (100) PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Antiguedad) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo,"
            SqlSalarios = SqlSalarios + "    SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(0) AS Incentivos,"
            SqlSalarios = SqlSalarios + "  SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso +  DetalleNomina.Antiguedad)"
            SqlSalarios = SqlSalarios + "   AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Reembolso) AS Reembolso,"
            SqlSalarios = SqlSalarios + "    Empleado.SueldoPeriodo"
            SqlSalarios = SqlSalarios + " FROM         DetalleNomina INNER JOIN"
            SqlSalarios = SqlSalarios + "   Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN"
            SqlSalarios = SqlSalarios + "  Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado"
            SqlSalarios = SqlSalarios + " GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano, Empleado.SueldoPeriodo"
            SqlSalarios = SqlSalarios + " HAVING      (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "')"
            SqlSalarios = SqlSalarios + " ORDER BY AÑO, Nomina.Mes"
      
      
      
      
      Else
            SqlSalarios = "SELECT DISTINCT"
            SqlSalarios = SqlSalarios + "  TOP (100) PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Antiguedad) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo,"
            SqlSalarios = SqlSalarios + "    SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,"
            SqlSalarios = SqlSalarios + "  SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso + DetalleNomina.Incentivos + DetalleNomina.Antiguedad)"
            SqlSalarios = SqlSalarios + "   AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Reembolso) AS Reembolso,"
            SqlSalarios = SqlSalarios + "    Empleado.SueldoPeriodo"
            SqlSalarios = SqlSalarios + " FROM         DetalleNomina INNER JOIN"
            SqlSalarios = SqlSalarios + "   Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN"
            SqlSalarios = SqlSalarios + "  Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado"
            SqlSalarios = SqlSalarios + " GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano, Empleado.SueldoPeriodo"
            SqlSalarios = SqlSalarios + " HAVING      (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "')"
            SqlSalarios = SqlSalarios + " ORDER BY AÑO, Nomina.Mes"
      End If
        
   Else
   
      If Me.ChkIncentivos.Value = xtpUnchecked Then
   
        SqlSalarios = "SELECT DISTINCT TOP (100) PERCENT DetalleNomina.CodEmpleado, AVG(Historico.SueldoActual) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SeptimoDia * 0) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(0) AS Incentivos,  SUM(DetalleNomina.Destajo + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso + AVG(Historico.SueldoActual) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Reembolso) As Reembolso FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                      "HAVING (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') ORDER BY AÑO, Nomina.Mes"
      Else
         SqlSalarios = "SELECT DISTINCT TOP (100) PERCENT DetalleNomina.CodEmpleado, AVG(Historico.SueldoActual) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SeptimoDia * 0) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,  SUM(DetalleNomina.Destajo + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso + DetalleNomina.Incentivos) + AVG(Historico.SueldoActual) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Reembolso) As Reembolso FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                      "HAVING (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') ORDER BY AÑO, Nomina.Mes"
      
      End If
   End If
    
       Me.TDBGridBonos.Visible = False
       Me.TDBGridSalarios.Visible = True
       



End If


FechaBusqueda1 = FechaBusqueda
FechaHistorico1 = FechaHistorico


Me.AdoSalarios.RecordSource = SqlSalarios
Me.AdoSalarios.Refresh


If SueldoFijo = True Then

If Me.AdoSalarios.Recordset.EOF Then
  SueldoPeriodo = 0
  
  
  
Else
  Me.AdoSalarios.Recordset.MoveLast
  'SueldoPeriodo = Me.AdoSalarios.Recordset("TotalIngresos")
  
  If DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
    If IsNull(Me.AdoSalarios.Recordset("SueldoPeriodo")) Then
        SueldoPeriodo = 0
    Else
     SueldoPeriodo = Me.AdoSalarios.Recordset("SueldoPeriodo") * 2
    End If
  
  
   
  ElseIf DtaTipoNomina.Recordset("Periodo") = "Mensual" Then
    SueldoPeriodo = Me.AdoSalarios.Recordset("SueldoPeriodo")
  ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
    If Me.ChkSueldoActual.Value = xtpUnchecked Then
      SueldoPeriodo = (Me.AdoSalarios.Recordset("SueldoPeriodo") / 14) * DiasMes
    Else
       SueldoPeriodo = Me.AdoSalarios.Recordset("SalarioBasico")
    End If
  End If
  
  'SueldoPeriodo = 0
End If



 Me.TxtSalarios.Caption = "Empleado con Salario Fijo"
   
    '/////////////VERIFICO SI SE UTILIZA LA ANTIGUEDAD COMO BASE//////////////////////
    '/////////////PARA EL CALCULO DE LA LIQUIDACION///////////////////////////////////
       If ChkAntiguedad.Value = 1 Then
        Años = Int(Me.TxtAnnos.Text)
        Me.AdoAntiguedad.RecordSource = "SELECT años_acum, porcent From Antiguedad Where (años_acum = " & Años & ")"
        Me.AdoAntiguedad.Refresh
        If Not Me.AdoAntiguedad.Recordset.EOF Then
         PAntiguedad = 1 + Me.AdoAntiguedad.Recordset("porcent")
         Me.txtAntiguedad.Text = Me.AdoAntiguedad.Recordset("porcent")
        Else
         Me.txtAntiguedad.Text = 0
        End If
        SalarioPromedio = SueldoPeriodo * PAntiguedad
        SalarioAlto = SueldoPeriodo * PAntiguedad
         
       Else
        SalarioPromedio = SueldoPeriodo
        SalarioAlto = SueldoPeriodo
         Me.txtAntiguedad.Text = 0
       End If
 
  
Else
    Me.txtAntiguedad.Text = 0
    Me.TxtSalarios.Caption = "Empleado con Salario Variable"
    Contador = 0
    TotalSalario = 0
    Salario = 0
    SalarioAlto = 0
    Do While Not Me.AdoSalarios.Recordset.EOF
    
      If Not IsNull(Me.AdoSalarios.Recordset("TotalIngresos")) Then
        TotalSalario = TotalSalario + Me.AdoSalarios.Recordset("TotalIngresos")
        Salario = Me.AdoSalarios.Recordset("TotalIngresos")
      Else
        Salario = 0
      End If
 
        If Salario > SalarioAlto Then
            SalarioAlto = Salario
        End If
 
        Contador = Contador + 1
        Me.AdoSalarios.Recordset.MoveNext
    Loop
   
   If Not Contador = 0 Then
    If Me.Check1.Value = 0 Then
    SalarioPromedio = TotalSalario / Contador '//////esto divide por los meses enteres //////
    
    Else
    
        If Contador < 6 Then
'            MDIPrimero.DtaEmpresa.Refresh
'            If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
'             If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("SalarioMinimo")) Then
'               SalarioMinimo = MDIPrimero.DtaEmpresa.Recordset("SalarioMinimo")
'             End If
'             TarifaHoraria = (SalarioMinimo / 30) / 8
'            End If
        
        
        
    '     SalarioPromedio = TotalSalario / Contador         'MesReal  '/////ESTO LO DIVIDO POR EL TIEMPO REAL TRABAJADO /////
           SalarioPromedio = SueldoActual
          
         Else
           SalarioPromedio = TotalSalario / Contador
         
         End If
      
    End If
   End If

 End If
 
    Me.TxtSalarioPromedio.Text = Format(SalarioPromedio, "##,##0.00")
    Me.TxtSalarioAlto.Text = Format(SalarioAlto, "##,##0.00")
    
    
     If Me.ChkSueldoActual.Value = xtpChecked Then
      Me.AdoSalarios.Refresh
      If Not Me.AdoSalarios.Recordset.EOF Then
'         SalarioPromedio = Me.AdoSalarios.Recordset("TotalIngresos")
'         Me.TxtSalarioPromedio.Text = Format(SalarioPromedio, "##,##0.00")
'         Me.TxtSalarioAlto.Text = Format(SalarioPromedio, "##,##0.00")
         
      End If
     End If
    
    Dim AñoActual As Integer ', CodTipoNomina As String
    
    CodigoEmpleado = Me.TxtCodEmpleado.Text


'/////////CONSULTA EL SALARIO Y TIPO DE NOMINA DEL EMPLEADO//////////////////////////

 sql = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento," & vbLf
 sql = sql & "Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.OtrosIngresos, Empleado.DescripOtrIngre," & vbLf
 sql = sql & "Empleado.ExentoIr, Empleado.PagoInssPatronal, Empleado.Activo, Empleado.Liquidado, Empleado.Ausente, Empleado.SalarioFijo," & vbLf
 sql = sql & "Empleado.SumarSubsidio , Empleado.PorcientoIncentivo, Empleado.Gravidez, TipoNomina.Periodo" & vbLf
 sql = sql & "FROM Empleado INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina" & vbLf
 sql = sql & "WHERE     (Empleado.CodEmpleado = '" & CodigoEmpleado & "') AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
 Me.DtaConsulta.RecordSource = sql
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  TipoNomina = Me.DtaConsulta.Recordset("Periodo")
  CodTipoNomina = Me.DtaConsulta.Recordset("CodTipoNomina")
 Else
  MsgBox "Este Empleado no Existe", vbCritical, "Sistema de Nominas"
  Exit Sub
 End If
    
    
    
    
       '/////////////BUSCO EL INICIO DEL PERIODO/////////////////////
       AñoActual = Year(Me.TxtUltFechaNomina.Value)
       Me.AdoInicioAño.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (año = " & AñoActual & ") AND (Periodo = 1)"
       Me.AdoInicioAño.Refresh
       If Not Me.AdoInicioAño.Recordset.EOF Then
        FechaInicio = Me.AdoInicioAño.Recordset("Inicio")
       End If
    
    
       '//////////////////////////////////////////////////////////////////
       '//////////CALCULO CUANTOS DIAS TIENE TRABAJADOS////////////////////
       '////////////////////////////////////////////////////////////////////
    Dim Fecha1 As Date, Fecha2 As Date
       
       Fecha1 = DateSerial(Year(FechaEgreso), 5, 31)
       FechaEgreso = Me.TxtUltFechaNomina.Value
       If FechaContrato < FechaInicio Then
        FechaInicioAgui = DateSerial(Year(FechaEgreso) - 1, 12, 1)
        Me.DTPFechaIniAgui.Value = FechaInicioAgui
        If FechaEgreso > Fecha1 Then
          Me.DTPFechaIniVaca.Value = DateSerial(Year(FechaEgreso), 6, 1)
        Else
          Me.DTPFechaIniVaca.Value = DateSerial(Year(FechaEgreso) - 1, 12, 1)
        End If
       Else
       FechaInicioAgui = FechaContrato
       Me.DTPFechaIniAgui.Value = FechaInicioAgui
       Me.DTPFechaIniVaca.Value = FechaContrato
       End If
       
       
    Me.DTPFechaIniAgui.Value = "01/12/" & Year(FechaBusqueda)
       
Exit Sub
TipoErr:
ControlErrores
CmdEfectuar.Enabled = False

End Sub

Private Sub TxtCodEmpleado1_Change()
  Dim DiaMes As Double, SalarioMinimo As Double, TarifaHoraria As Double

'    SQlEmpleado = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.SueldoActualBasico, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2,Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion,Empleado.Sexo , Empleado.Activo, Empleado.Liquidado, Empleado.Ausente, Empleado.Gravidez,Empleado.TarifaHoraria, Empleado.SalarioFijo, Empleado.SueldoPeriodo FROM Departamento INNER JOIN Cargo INNER JOIN" & vbLf
'    SQlEmpleado = SQlEmpleado & "Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento" & vbLf
'    SQlEmpleado = SQlEmpleado & "WHERE  (Empleado.CodEmpleado1 = '" & Me.TxtCodEmpleado1.Text & "') AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
    
    SQlEmpleado = "SELECT Empleado.SalarioFijo, Empleado.SueldoPeriodo, Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion, Empleado.Direccion AS Expr1, Empleado.Sexo, Empleado.Activo, Empleado.TarifaHoraria, Empleado.SueldoActualBasico, Historico.SueldoActual FROM Departamento INNER JOIN Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado Where (Empleado.CodEmpleado1 = '" & Me.txtCodEmpleado1.Text & "') AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
    DtaEmpleado.RecordSource = SQlEmpleado
    DtaEmpleado.Refresh
    
     
    If Not DtaEmpleado.Recordset.EOF Then
    
    TxtNombre1 = DtaEmpleado.Recordset("Nombre1")
    TxtNombre2 = DtaEmpleado.Recordset("Nombre2")
    TxtApellido1 = DtaEmpleado.Recordset("Apellido1")
    TxtApellido2 = DtaEmpleado.Recordset("Apellido2")
    Me.CmdAcercade.Caption = Me.txtCodEmpleado1.Text + "-" + Me.TxtNombre1.Text + " " + Me.TxtNombre2.Text + " " + Me.TxtApellido1.Text + " " + Me.TxtApellido2.Text
    TxtDireccion = DtaEmpleado.Recordset("Direccion")
    txtCargo = DtaEmpleado.Recordset("Cargo")
    txtDepartamento = DtaEmpleado.Recordset("departamento")
    txtSexo = DtaEmpleado.Recordset("sexo")
    
    MDIPrimero.DtaControles.Refresh
    If Not MDIPrimero.DtaControles.Recordset.EOF Then
     DiaMes = MDIPrimero.DtaControles.Recordset("DiasMes")
    End If
    
    MDIPrimero.DtaEmpresa.Refresh
    If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
     If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("SalarioMinimo")) Then
       SalarioMinimo = MDIPrimero.DtaEmpresa.Recordset("SalarioMinimo")
     End If
     TarifaHoraria = (SalarioMinimo / 30) / 8
    End If
    
    Me.TxtCodEmpleado.Text = DtaEmpleado.Recordset("CodEmpleado")
    Me.TxtTarifa.Text = Format(TarifaHoraria, "##,##0.000000")
    Me.txtSalarioBasico.Text = Format(SalarioMinimo, "##,##0.00")
'    Me.TxtTarifa.Text = DtaEmpleado.Recordset("TarifaHoraria")
'    Me.txtSalarioBasico.Text = Format(DtaEmpleado.Recordset("TarifaHoraria") * DiaMes * 8, "##,##0.00")
    End If

End Sub

Private Sub TxtCodEmpleado1_KeyPress(KeyAscii As Integer)
Dim SqlSalarios As String, DiaMes As Double
Dim SQlEmpleado As String
If KeyAscii = 13 Then

    SQlEmpleado = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2,Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion,Empleado.Sexo , Empleado.Activo, Empleado.Liquidado, Empleado.Ausente, Empleado.Gravidez,Empleado.TarifaHoraria FROM Departamento INNER JOIN Cargo INNER JOIN" & vbLf
    SQlEmpleado = SQlEmpleado & "Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento" & vbLf
    SQlEmpleado = SQlEmpleado & "WHERE  (Empleado.CodEmpleado1 = '" & Me.txtCodEmpleado1.Text & "') AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
    DtaEmpleado.RecordSource = SQlEmpleado
    DtaEmpleado.Refresh
    
     
    If Not DtaEmpleado.Recordset.EOF Then
    
    TxtNombre1 = DtaEmpleado.Recordset("Nombre1")
    TxtNombre2 = DtaEmpleado.Recordset("Nombre2")
    TxtApellido1 = DtaEmpleado.Recordset("Apellido1")
    TxtApellido2 = DtaEmpleado.Recordset("Apellido2")
    Me.CmdAcercade.Caption = Me.txtCodEmpleado1.Text + "-" + Me.TxtNombre1.Text + " " + Me.TxtNombre2.Text + " " + Me.TxtApellido1.Text + " " + Me.TxtApellido2.Text
    TxtDireccion = DtaEmpleado.Recordset("Direccion")
    txtCargo = DtaEmpleado.Recordset("Cargo")
    txtDepartamento = DtaEmpleado.Recordset("departamento")
    txtSexo = DtaEmpleado.Recordset("sexo")

    MDIPrimero.DtaControles.Refresh
    If Not MDIPrimero.DtaControles.Recordset.EOF Then
     DiaMes = MDIPrimero.DtaControles.Recordset("DiasMes")
    End If
    
    Me.TxtCodEmpleado.Text = DtaEmpleado.Recordset("CodEmpleado")
    Me.TxtTarifa.Text = DtaEmpleado.Recordset("TarifaHoraria")
    Me.txtSalarioBasico.Text = Format(DtaEmpleado.Recordset("TarifaHoraria") * DiaMes * 8, "##,##0.00")
    End If

End If
End Sub

Private Sub TxtDescuentoDias_Change()
If Not TxtDescuentoDias.Text = "" Then
 If Not IsNumeric(TxtDescuentoDias.Text) Then
  MsgBox "El numero Digitado no es Numerico", vbCritical, "Sistema de Nominas"
  Me.TxtDescuentoDias.Text = ""
 End If
End If
End Sub

Private Sub TxtFechaHistorial_Change()
Dim FechaEgreso As Date, FechaContrato As Date, Año As Integer, Mes As Integer, DiasMes As Double, DiasReales As Double, MesReal As Double
Dim FechaBusqueda As Date, TotalSalario As Double, SalarioPromedio As Double, Contador As Integer, i As Integer
Dim SqlSalarios As String, SalarioAlto As Double, Salario As Double, FechaHistorico As Date, NumeroEmpleado As Double
    FechaEgreso = Me.TxtFechaHistorial.Value
FechaContrato = Me.TxtFechaContrato.Text
'//////////SUMO 1 PARA AJUSTAR QUE SIEMPRE DA 1 DIA MENOS//////
'annos = CDbl(FechaEgreso) - CDbl(FechaContrato) + 1
'TxtAnnos.Text = Format(annos / 365, "###,##0.00")
'TxtMeses.Text = Format(annos / 30.41, "###,##0.00")
'Me.txtDiasTrabajados.Text = Format(annos, "###,##0")
'Dias = annos



'///////////Busco la Fecha para la Busqueda////////////////////////////

NumeroEmpleado = Me.TxtCodEmpleado.Text

SqlSalarios = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.Incentivos) AS Incentivos, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.Comisiones) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes AS MES, Nomina.Ano AS AÑO,SUM(DetalleNomina.Comisiones) As Comisiones " & _
             "FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
             "HAVING      (DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") AND (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (MAX(Nomina.FechaNomina) <= CONVERT(DATETIME, '" & Format(FechaEgreso, "yyyy/mm/dd") & "', 102))"


Me.DtaConsulta.RecordSource = SqlSalarios
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
  Me.DtaConsulta.Recordset.MoveLast
Else
 FechaHistorico = Format(Now, "dd/mm/yyyy")
 FechaBusqueda = Format(Now, "dd/mm/yyyy")
End If
i = 0


Do While Not Me.DtaConsulta.Recordset.BOF
  If i = 1 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")

  ElseIf i = 5 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    Exit Do
  ElseIf i = 0 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    FechaHistorico = Me.DtaConsulta.Recordset("FechaFin")
  Else
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
  End If
  i = i + 1

  Me.DtaConsulta.Recordset.MovePrevious
Loop


FechaEgreso = Me.TxtFechaHistorial.Value
'FechaHistorico = DateSerial(Year(FechaEgreso), Month(FechaEgreso), 1 - 1)
FechaContrato = Me.TxtFechaContrato.Text
'FechaBusqueda = DateSerial(Year(FechaEgreso), Month(FechaEgreso) - 6, 1)
Año = Year(FechaBusqueda)
Mes = Month(FechaBusqueda)

MDIPrimero.DtaControles.Refresh
DiasMes = MDIPrimero.DtaControles.Recordset("DiasMes")
   
DiasReales = CalcularDiasVaca(Me.DTPFechaIniVaca.Value, FechaHistorico)
MesReal = DiasReales / DiasMes
Me.CmdEfectuar.Enabled = False

Me.DTPFechaIniAgui.Value = "01/12/" & Year(FechaBusqueda)


'SQlSalarios = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo,SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO,SUM(DetalleNomina.Comisiones) As Comisiones " & _
'             "FROM  DetalleNomina INNER JOIN  Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
'             "HAVING (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (MIN(Nomina.FechaNomina)Between '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') " & _
'             "ORDER BY Nomina.Ano, Nomina.Mes "

  If Me.ChkSueldoActual.Value = xtpUnchecked Then
        'SqlSalarios = "SELECT DISTINCT TOP 100 PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo,SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones + DetalleNomina.Incentivos)AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO,SUM(DetalleNomina.Comisiones) As Comisiones, Empleado.SueldoPeriodo " & _
                     ' "FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                     ' "HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) Between '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') " & _
                      '"ORDER BY Nomina.Ano, Nomina.Mes "
                        SqlSalarios = "SELECT DISTINCT"
        SqlSalarios = SqlSalarios + "  TOP (100) PERCENT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Antiguedad) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo,"
        SqlSalarios = SqlSalarios + "    SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos,"
        SqlSalarios = SqlSalarios + "  SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso + DetalleNomina.Incentivos + DetalleNomina.Antiguedad)"
        SqlSalarios = SqlSalarios + "   AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Reembolso) AS Reembolso,"
        SqlSalarios = SqlSalarios + "    Empleado.SueldoPeriodo"
        SqlSalarios = SqlSalarios + " FROM         DetalleNomina INNER JOIN"
        SqlSalarios = SqlSalarios + "   Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN"
        SqlSalarios = SqlSalarios + "  Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado"
        SqlSalarios = SqlSalarios + " GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano, Empleado.SueldoPeriodo"
        SqlSalarios = SqlSalarios + " HAVING      (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "')"
        SqlSalarios = SqlSalarios + " ORDER BY AÑO, Nomina.Mes"
                      
                      
   Else
'        SQlSalarios = "SELECT DISTINCT TOP (100) PERCENT DetalleNomina.CodEmpleado, AVG(Historico.SueldoActual) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(0) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos, SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones + DetalleNomina.Incentivos) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Comisiones) As Comisiones FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                      "HAVING (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') ORDER BY AÑO, Nomina.Mes"
    
         SqlSalarios = "SELECT DISTINCT TOP (100) PERCENT DetalleNomina.CodEmpleado, AVG(Historico.SueldoActual) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(0) AS Septimo,SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos, AVG(Historico.SueldoActual) + SUM(DetalleNomina.Destajo + DetalleNomina.OtrosIngresos + DetalleNomina.Reembolso + DetalleNomina.Incentivos) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Reembolso) AS Reembolso FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano  " & _
                       "HAVING (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "') ORDER BY AÑO, Nomina.Mes"
   End If

FechaBusqueda1 = FechaBusqueda
FechaHistorico1 = FechaHistorico

Me.AdoSalarios.RecordSource = SqlSalarios
Me.AdoSalarios.Refresh


If SueldoFijo = True Then
 
 If Me.AdoSalarios.Recordset.EOF Then
  SueldoPeriodo = 0
 Else
  Me.AdoSalarios.Recordset.MoveLast
'  SueldoPeriodo = Me.AdoSalarios.Recordset("TotalIngresos")
 
  If DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
    SueldoPeriodo = Me.AdoSalarios.Recordset("SueldoPeriodo") * 2
  ElseIf DtaTipoNomina.Recordset("Periodo") = "Mensual" Then
    SueldoPeriodo = Me.AdoSalarios.Recordset("SueldoPeriodo")
  End If
  

 End If
 Me.TxtSalarios.Caption = "Empleado con Salario Fijo"
   
    '/////////////VERIFICO SI SE UTILIZA LA ANTIGUEDAD COMO BASE//////////////////////
    '/////////////PARA EL CALCULO DE LA LIQUIDACION///////////////////////////////////
       If ChkAntiguedad.Value = 1 Then
        Años = Int(Me.TxtAnnos.Text)
        Me.AdoAntiguedad.RecordSource = "SELECT años_acum, porcent From Antiguedad Where (años_acum = " & Años & ")"
        Me.AdoAntiguedad.Refresh
        If Not Me.AdoAntiguedad.Recordset.EOF Then
         PAntiguedad = 1 + Me.AdoAntiguedad.Recordset("porcent")
         Me.txtAntiguedad.Text = Me.AdoAntiguedad.Recordset("porcent")
        Else
         Me.txtAntiguedad.Text = 0
        End If
        SalarioPromedio = SueldoPeriodo * PAntiguedad
        SalarioAlto = SueldoPeriodo * PAntiguedad
         
       Else
        SalarioPromedio = SueldoPeriodo
        SalarioAlto = SueldoPeriodo
         Me.txtAntiguedad.Text = 0
       End If
 
  
Else
    Me.txtAntiguedad.Text = 0
    Me.TxtSalarios.Caption = "Empleado con Salario Variable"
    Contador = 0
    TotalSalario = 0
    Salario = 0
    SalarioAlto = 0
    Do While Not Me.AdoSalarios.Recordset.EOF
        TotalSalario = TotalSalario + Me.AdoSalarios.Recordset("TotalIngresos")
        Salario = Me.AdoSalarios.Recordset("TotalIngresos")
 
        If Salario > SalarioAlto Then
            SalarioAlto = Salario
        End If
 
        Contador = Contador + 1
        Me.AdoSalarios.Recordset.MoveNext
    Loop

   If Not Contador = 0 Then
    If Me.Check1.Value = 0 Then
    SalarioPromedio = TotalSalario / Contador '//////esto divide por los meses enteres //////
    Else
     SalarioPromedio = TotalSalario / MesReal  '/////ESTO LO DIVIDO POR EL TIEMPO REAL TRABAJADO /////
    End If
   End If
   

 End If
 
 
    Me.TxtSalarioPromedio.Text = Format(SalarioPromedio, "##,##0.00")
    Me.TxtSalarioAlto.Text = Format(SalarioAlto, "##,##0.00")


    Dim AñoActual As Integer, CodTipoNomina As String
    
    CodigoEmpleado = Me.TxtCodEmpleado.Text


'/////////CONSULTA EL SALARIO Y TIPO DE NOMINA DEL EMPLEADO//////////////////////////

 sql = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento," & vbLf
 sql = sql & "Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.OtrosIngresos, Empleado.DescripOtrIngre," & vbLf
 sql = sql & "Empleado.ExentoIr, Empleado.PagoInssPatronal, Empleado.Activo, Empleado.Liquidado, Empleado.Ausente, Empleado.SalarioFijo," & vbLf
 sql = sql & "Empleado.SumarSubsidio , Empleado.PorcientoIncentivo, Empleado.Gravidez, TipoNomina.Periodo" & vbLf
 sql = sql & "FROM Empleado INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina" & vbLf
 sql = sql & "WHERE     (Empleado.CodEmpleado = '" & CodigoEmpleado & "') AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
 Me.DtaConsulta.RecordSource = sql
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  TipoNomina = Me.DtaConsulta.Recordset("Periodo")
  CodTipoNomina = Me.DtaConsulta.Recordset("CodTipoNomina")
 Else
  MsgBox "Este Empleado no Existe", vbCritical, "Sistema de Nominas"
  Exit Sub
 End If
    
    
    
    
       '/////////////BUSCO EL INICIO DEL PERIODO/////////////////////
       AñoActual = Year(Me.TxtUltFechaNomina.Value)
       Me.AdoInicioAño.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (año = " & AñoActual & ") AND (Periodo = 1)"
       Me.AdoInicioAño.Refresh
       If Not Me.AdoInicioAño.Recordset.EOF Then
        FechaInicio = Me.AdoInicioAño.Recordset("Inicio")
       End If
    
    
       '//////////////////////////////////////////////////////////////////
       '//////////CALCULO CUANTOS DIAS TIENE TRABAJADOS////////////////////
       '////////////////////////////////////////////////////////////////////
       '///SEMESTRE VACACIONES 01/12/2005-31/05/2006,  01/06/2006-31/11/2006
       
    Dim Fecha1 As Date, Fecha2 As Date
       
       Fecha1 = DateSerial(Year(FechaEgreso), 5, 31)
       FechaEgreso = Me.TxtFechaHistorial.Value
       If FechaContrato < FechaInicio Then
        FechaInicioAgui = DateSerial(Year(FechaEgreso) - 1, 12, 1)
        Me.DTPFechaIniAgui.Value = FechaInicioAgui
        If FechaEgreso > Fecha1 Then
          Me.DTPFechaIniVaca.Value = DateSerial(Year(FechaEgreso), 6, 1)
        Else
          Me.DTPFechaIniVaca.Value = DateSerial(Year(FechaEgreso) - 1, 12, 1)
        End If
       Else
       FechaInicioAgui = FechaContrato
       Me.DTPFechaIniAgui.Value = FechaInicioAgui
       Me.DTPFechaIniVaca.Value = FechaContrato
       End If
End Sub

Private Sub TxtUltFechaNomina_Change()
Dim FechaEgreso As Date, FechaContrato As Date, Año As Integer, Mes As Integer
Dim FechaBusqueda As Date, TotalSalario As Double, SalarioPromedio As Double, Contador As Integer, i As Integer
Dim SqlSalarios As String, SalarioAlto As Double, Salario As Double, FechaHistorico As Date, NumeroEmpleado As Integer
Dim DiasMes As Double, Dia As Double, Meses As Double, DiasAcumulados As Double, Fecha1 As Date, Fecha2 As Date
FechaEgreso = Me.TxtUltFechaNomina.Value
FechaContrato = Me.TxtFechaContrato.Text

'//////////SUMO 1 PARA AJUSTAR QUE SIEMPRE DA 1 DIA MENOS//////
'annos = CDbl(FechaEgreso) - CDbl(FechaContrato) + 1
MDIPrimero.DtaControles.Refresh
DiasMes = MDIPrimero.DtaControles.Recordset("DiasMes")
'annos = CalcularDiasVaca(FechaContrato, FechaEgreso)

    If DateDiff("d", FechaContrato, FechaEgreso) <= 31 Then
            If Month(FechaContrato) = Month(FechaEgreso) Then
                     If DiasMes = 30 Then
                           Dias = Day(DateSerial(Year(FechaEgreso), Month(FechaEgreso) + 1, 0))
                           If Dias = 31 Then
                             Fecha2 = "31/" & Month(FechaEgreso) & " / " & Year(FechaEgreso)
                           Else
                             Fecha2 = FechaEgreso
                           End If
                       
                           If Fecha2 = Me.TxtUltFechaNomina.Value Then
                             ' SE QUEDA EN COMENTARIO , POR QUE CUANDO QUIERO DAR DE BAJA Y SOLO TIENEN 1 DIAS DEL 03/06/2019 AL 04/06/2019  ME PODE 3O DIAS
               '               FechaEgreso = "30/ " & Month(FechaEgreso) & " / " & Year(FechaEgreso)
               '               Me.TxtUltFechaNomina.Value = FechaEgreso
                           End If
                     End If
                  Dias = DateDiff("d", FechaContrato, FechaEgreso) + 1
             Else
      
                    '////////////////////////////VERIFICO EL MES ////////////////////////////////////////
                    If Month(FechaContrato) = 2 Then
                       Fecha1 = DateSerial(Year(FechaContrato), Month(FechaContrato) + 1, 1 - 1)
                       
                        If Month(FechaContrato) = 2 Then
                          If Day(Fecha1) = 28 Then
                             Dias = Dias + 2
                          ElseIf Day(Fecha1) = 29 Then
                            Dias = Dias + 1
                          End If
                        End If
                        
                    Fecha2 = "01/ " & Month(FechaEgreso) & " / " & Year(FechaEgreso)
                    Dias = (DateDiff("d", FechaContrato, Fecha1) + 1) + (DateDiff("d", Fecha2, FechaEgreso) + 1)
                      
                    Else
                    
                        Dias = DateDiff("d", FechaContrato, FechaEgreso) + 1
'                       Fecha1 = "30/ " & Month(FechaEgreso) & " / " & Year(FechaEgreso)
                    End If
                

                    

        
        
             End If
     Else
       Dias = Format(CalcularDiasAntiguedad(FechaContrato, FechaEgreso) * 12, "####0")
     End If




annos = Dias
TxtAnnos.Text = Format(annos / 358, "###,##0.00")  ''11 meses por 30 mas 28 de febrero
Meses = Dias / DiasMes
Mes = Int(Meses)
Dia = Format((Meses - Mes) * DiasMes, "####0")

''annos = Dias
''TxtAnnos.Text = Format(annos / 360, "###,##0.00")
''
''Meses = Dias / DiasMes
''Mes = Int(Meses)
''Dia = Format((Meses - Mes) * DiasMes, "####0")



TxtMeses.Text = Mes & " m y " & Dia & " d"
'Format(Dias / DiasMes, "###,##0")
Me.TxtDiasTrabajados.Text = Format(annos, "###,##0")
'Dias = annos



'Me.TxtDiasTrabajados.Text = Format(tempDias, "###,##0")  'Dias
Me.CmdEfectuar.Enabled = False

If CDbl(Me.TxtDiasTrabajados.Text) < 14 Then
   Me.TxtDias.Text = Me.TxtDiasTrabajados.Text
End If

'Me.DTPFechaIniAgui.Value = "01/12/" & Year(Me.TxtUltFechaNomina.Value)

'///////////Busco la Fecha para la Busqueda////////////////////////////

'NumeroEmpleado = Me.txtCodEmpleado.Text
'
'SQLSalarios = "SELECT DISTINCT" & vbLf
'SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
'SQLSalarios = SQLSalarios & "AS Destajo, SUM(dbo.DetalleNomina.Incentivos) AS Incentivos, SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo)" & vbLf
'SQLSalarios = SQLSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes AS MES," & vbLf
'SQLSalarios = SQLSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
'SQLSalarios = SQLSalarios & "FROM   dbo.DetalleNomina INNER JOIN" & vbLf
'SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
'SQLSalarios = SQLSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
'SQLSalarios = SQLSalarios & "Having (dbo.DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") And (Sum(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0)"
'Me.DtaConsulta.RecordSource = SQLSalarios
'Me.DtaConsulta.Refresh
'If Not Me.DtaConsulta.Recordset.EOF Then
'  Me.DtaConsulta.Recordset.MoveLast
'Else
' FechaHistorico = Format(Now, "dd/mm/yyyy")
' FechaBusqueda = Format(Now, "dd/mm/yyyy")
'End If
'I = 0
'
'
'Do While Not Me.DtaConsulta.Recordset.BOF
'  If I = 1 Then
'    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
'
'  ElseIf I = 5 Then
'    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
'    Exit Do
'  ElseIf I = 0 Then
'    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
'    FechaHistorico = Me.DtaConsulta.Recordset("FechaFin")
'  Else
'    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
'  End If
'  I = I + 1
'
'  Me.DtaConsulta.Recordset.MovePrevious
'Loop
'
'
'FechaEgreso = Me.TxtUltFechaNomina.Value
''FechaHistorico = DateSerial(Year(FechaEgreso), Month(FechaEgreso), 1 - 1)
'FechaContrato = Me.TxtFechaContrato.Text
''FechaBusqueda = DateSerial(Year(FechaEgreso), Month(FechaEgreso) - 6, 1)
'Año = Year(FechaBusqueda)
'Mes = Month(FechaBusqueda)
''
''    SQLSalarios = "SELECT DISTINCT" & vbLf
''    SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
''    SQLSalarios = SQLSalarios & "AS Destajo, SUM(dbo.DetalleNomina.SeptimoDia) AS Septimo, SUM(dbo.DetalleNomina.OtrosIngresos) AS Otros, SUM(dbo.DetalleNomina.Incentivos)" & vbLf
''    SQLSalarios = SQLSalarios & "AS Incentivos," & vbLf
''    SQLSalarios = SQLSalarios & "SUM (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.OtrosIngresos)" & vbLf
''    SQLSalarios = SQLSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes," & vbLf
''    SQLSalarios = SQLSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
''    SQLSalarios = SQLSalarios & "FROM    dbo.DetalleNomina INNER JOIN" & vbLf
''    SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
''    SQLSalarios = SQLSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
''    SQLSalarios = SQLSalarios & "HAVING(SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0) And (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND" & vbLf
''    SQLSalarios = SQLSalarios & "'" & Format(FechaHistorico, "yyyymmdd") & "')" & vbLf
''    SQLSalarios = SQLSalarios & "ORDER BY dbo.Nomina.Ano, dbo.Nomina.Mes"
''
''Me.AdoSalarios.RecordSource = SQLSalarios
''Me.AdoSalarios.Refresh
'
'
'If SueldoFijo = True Then
'
' If Me.AdoSalarios.Recordset.EOF Then
'  SueldoPeriodo = 0
' Else
'  Me.AdoSalarios.Recordset.MoveLast
'  SueldoPeriodo = Me.AdoSalarios.Recordset("TotalIngresos")
' End If
' Me.TxtSalarios.Caption = "Empleado con Salario Fijo"
'
'    '/////////////VERIFICO SI SE UTILIZA LA ANTIGUEDAD COMO BASE//////////////////////
'    '/////////////PARA EL CALCULO DE LA LIQUIDACION///////////////////////////////////
'       If ChkAntiguedad.Value = 1 Then
'        Años = Int(Me.TxtAnnos.Text)
'        Me.adoAntiguedad.RecordSource = "SELECT años_acum, porcent From Antiguedad Where (años_acum = " & Años & ")"
'        Me.adoAntiguedad.Refresh
'        If Not Me.adoAntiguedad.Recordset.EOF Then
'         PAntiguedad = 1 + Me.adoAntiguedad.Recordset("porcent")
'         Me.txtAntiguedad.Text = Me.adoAntiguedad.Recordset("porcent")
'        Else
'         Me.txtAntiguedad.Text = 0
'        End If
'        SalarioPromedio = SueldoPeriodo * PAntiguedad
'        SalarioAlto = SueldoPeriodo * PAntiguedad
'
'       Else
'        SalarioPromedio = SueldoPeriodo
'        SalarioAlto = SueldoPeriodo
'         Me.txtAntiguedad.Text = 0
'       End If
'
'
'Else
'    Me.txtAntiguedad.Text = 0
'    Me.TxtSalarios.Caption = "Empleado con Salario Variable"
'    Contador = 0
'    TotalSalario = 0
'    Salario = 0
'    SalarioAlto = 0
'    Do While Not Me.AdoSalarios.Recordset.EOF
'        TotalSalario = TotalSalario + Me.AdoSalarios.Recordset("TotalIngresos")
'        Salario = Me.AdoSalarios.Recordset("TotalIngresos")
'
'        If Salario > SalarioAlto Then
'            SalarioAlto = Salario
'        End If
'
'        Contador = Contador + 1
'        Me.AdoSalarios.Recordset.MoveNext
'    Loop
'
'   If Contador = 0 Then
'    SalarioPromedio = 0
'   Else
'    SalarioPromedio = TotalSalario / Contador
'   End If
' End If
'    Me.TxtSalarioPromedio.Text = Format(SalarioPromedio, "##,##0.00")
'    Me.TxtSalarioAlto.Text = Format(SalarioAlto, "##,##0.00")
'
'
'    Dim AñoActual As Integer, CodTipoNomina As String
'
'    CodigoEmpleado = Me.txtCodEmpleado.Text
'
'
''/////////CONSULTA EL SALARIO Y TIPO DE NOMINA DEL EMPLEADO//////////////////////////
'
' SQL = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento," & vbLf
' SQL = SQL & "Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.OtrosIngresos, Empleado.DescripOtrIngre," & vbLf
' SQL = SQL & "Empleado.ExentoIr, Empleado.PagoInssPatronal, Empleado.Activo, Empleado.Liquidado, Empleado.Ausente, Empleado.SalarioFijo," & vbLf
' SQL = SQL & "Empleado.SumarSubsidio , Empleado.PorcientoIncentivo, Empleado.Gravidez, TipoNomina.Periodo" & vbLf
' SQL = SQL & "FROM Empleado INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina" & vbLf
' SQL = SQL & "WHERE     (Empleado.CodEmpleado = '" & CodigoEmpleado & "') AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
' Me.DtaConsulta.RecordSource = SQL
' Me.DtaConsulta.Refresh
' If Not DtaConsulta.Recordset.EOF Then
'  TipoNomina = Me.DtaConsulta.Recordset("Periodo")
'  CodTipoNomina = Me.DtaConsulta.Recordset("CodTipoNomina")
' Else
'  MsgBox "Este Empleado no Existe", vbCritical, "Sistema de Nominas"
'  Exit Sub
' End If
'
'
'
'
'       '/////////////BUSCO EL INICIO DEL PERIODO/////////////////////
'       AñoActual = Year(Me.TxtUltFechaNomina.Value)
'       Me.AdoInicioAño.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (año = " & AñoActual & ") AND (Periodo = 1)"
'       Me.AdoInicioAño.Refresh
'       If Not Me.AdoInicioAño.Recordset.EOF Then
'        FechaInicio = Me.AdoInicioAño.Recordset("Inicio")
'       End If
'
'
'       '//////////////////////////////////////////////////////////////////
'       '//////////CALCULO CUANTOS DIAS TIENE TRABAJADOS////////////////////
'       '////////////////////////////////////////////////////////////////////
'       '///SEMESTRE VACACIONES 01/12/2005-31/05/2006,  01/06/2006-31/11/2006
'
'    Dim Fecha1 As Date, Fecha2 As Date
'
'       Fecha1 = DateSerial(Year(FechaEgreso), 5, 31)
'       FechaEgreso = Me.TxtUltFechaNomina.Value
'       If FechaContrato < FechaInicio Then
'        FechaInicioAgui = DateSerial(Year(FechaEgreso) - 1, 12, 1)
'        Me.DTPFechaIniAgui.Value = FechaInicioAgui
'        If FechaEgreso > Fecha1 Then
'          Me.DTPFechaIniVaca.Value = DateSerial(Year(FechaEgreso), 6, 1)
'        Else
'          Me.DTPFechaIniVaca.Value = DateSerial(Year(FechaEgreso) - 1, 12, 1)
'        End If
'       Else
'       FechaInicioAgui = FechaContrato
'       Me.DTPFechaIniAgui.Value = FechaInicioAgui
'       Me.DTPFechaIniVaca.Value = FechaContrato
'       End If

End Sub
