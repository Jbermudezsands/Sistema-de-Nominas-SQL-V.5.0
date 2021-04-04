VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{AF8CD3F4-666F-11D1-940D-000021A73813}#5.0#0"; "osProgress.ocx"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmBajas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Despidos y Renuncias"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoDatosEmpresa 
      Height          =   495
      Left            =   4440
      Top             =   8400
      Width           =   4095
      _ExtentX        =   7223
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoDatosEmpresa"
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
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   5760
      TabIndex        =   70
      Top             =   7560
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc AdoEmpleados 
      Height          =   495
      Left            =   480
      Top             =   8640
      Width           =   3855
      _ExtentX        =   6800
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoEmpleados"
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
      Top             =   9840
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
      Top             =   9360
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
      Top             =   9600
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
      Left            =   3360
      Top             =   9840
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
   Begin Progress.osProgress osProgress 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   3735
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
   Begin VB.CommandButton CmdCalculos 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton CmdEfectuar 
      Caption         =   "Procesar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CommandButton CmdCalcular 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   10080
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc DtaNominas 
      Height          =   375
      Left            =   3120
      Top             =   9720
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
      Top             =   9360
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
      Left            =   5160
      Top             =   9840
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
      Left            =   5880
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
      Left            =   0
      Top             =   9720
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
      Top             =   9960
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
      Top             =   9360
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
      Top             =   9600
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
      Top             =   9840
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
      Top             =   9720
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
      Top             =   9360
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
      Left            =   360
      Top             =   9600
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
      Top             =   9720
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
      Top             =   9480
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
      Height          =   6855
      Left            =   120
      ScaleHeight     =   6795
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   600
      Width           =   9495
      Begin TabDlg.SSTab SSTab1 
         Height          =   6495
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   11456
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Datos Generales"
         TabPicture(0)   =   "FrmBajas.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label14"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label11"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "TxtMotivo"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "TxtDias"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Frame1"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "TDBGrid1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Frame5"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "Historial Salarial"
         TabPicture(1)   =   "FrmBajas.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TDBGridSalarios"
         Tab(1).Control(1)=   "Frame3"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Ingresos / Egresos"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame2"
         Tab(2).Control(1)=   "Frame4"
         Tab(2).ControlCount=   2
         Begin VB.Frame Frame3 
            Caption         =   "Calculos Basicos del Salario"
            Height          =   2295
            Left            =   -74880
            TabIndex        =   71
            Top             =   3000
            Width           =   9015
            Begin VB.TextBox TxtSalarioPromedio 
               Height          =   285
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   85
               Text            =   "0.00"
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox TxtSalarioAlto 
               Height          =   285
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   84
               Text            =   "0.00"
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox TxtAntiguedad 
               Height          =   285
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   82
               Text            =   "0.00"
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox TxtSalarioBasico 
               Height          =   285
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   75
               Text            =   "0.00"
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox TxtTarifa 
               Height          =   285
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   73
               Text            =   "0.00"
               Top             =   1800
               Width           =   1335
            End
            Begin SmartButtonProject.SmartButton CmdImprimirHistorial 
               Height          =   855
               Left            =   6600
               TabIndex        =   72
               Top             =   840
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   1508
               Caption         =   "Imp Historial"
               Picture         =   "FrmBajas.frx":0038
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
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmBajas.frx":0352
               TabIndex        =   74
               Top             =   1800
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
               Height          =   375
               Left            =   240
               OleObjectBlob   =   "FrmBajas.frx":03CC
               TabIndex        =   76
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
               Format          =   21364737
               CurrentDate     =   38821
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
               Height          =   255
               Left            =   3240
               OleObjectBlob   =   "FrmBajas.frx":0446
               TabIndex        =   78
               Top             =   1200
               Width           =   1815
            End
            Begin MSComCtl2.DTPicker DTPFechaIniVaca 
               Height          =   285
               Left            =   5040
               TabIndex        =   79
               Top             =   840
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               Format          =   21364737
               CurrentDate     =   38821
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
               Height          =   255
               Left            =   3240
               OleObjectBlob   =   "FrmBajas.frx":04D0
               TabIndex        =   80
               Top             =   840
               Width           =   1935
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmBajas.frx":055C
               TabIndex        =   81
               Top             =   1440
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmBajas.frx":05CE
               TabIndex        =   83
               Top             =   720
               Width           =   1335
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmBajas.frx":064C
               TabIndex        =   86
               Top             =   360
               Width           =   1455
            End
            Begin Threed.SSCommand TxtSalarios 
               Height          =   525
               Left            =   3240
               TabIndex        =   87
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
            Begin SmartButtonProject.SmartButton CmdDetalle 
               Height          =   855
               Left            =   7800
               TabIndex        =   88
               Top             =   840
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   1508
               Caption         =   "Imp Detalle"
               Picture         =   "FrmBajas.frx":06CA
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
         End
         Begin VB.Frame Frame5 
            Caption         =   "Frame5"
            Height          =   2535
            Left            =   600
            TabIndex        =   8
            Top             =   1440
            Visible         =   0   'False
            Width           =   5535
            Begin VB.TextBox TxtCodEmpleado1 
               Height          =   285
               Left            =   1680
               TabIndex        =   20
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox TxtCodTipoNomina 
               Height          =   285
               Left            =   3480
               TabIndex        =   19
               Top             =   3120
               Width           =   615
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
               Left            =   3120
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox TxtCodEmpleado 
               Enabled         =   0   'False
               Height          =   285
               Left            =   3480
               TabIndex        =   17
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox TxtSexo 
               Height          =   285
               Left            =   1680
               TabIndex        =   16
               Top             =   3120
               Width           =   1455
            End
            Begin VB.TextBox TxtCargo 
               Height          =   285
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   2760
               Width           =   2415
            End
            Begin VB.TextBox TxtDepartamento 
               Height          =   285
               Left            =   1680
               TabIndex        =   14
               Top             =   2400
               Width           =   2415
            End
            Begin VB.TextBox TxtNombre2 
               Height          =   285
               Left            =   1680
               MaxLength       =   20
               TabIndex        =   13
               Top             =   960
               Width           =   2415
            End
            Begin VB.TextBox TxtApellido1 
               Height          =   285
               Left            =   1680
               MaxLength       =   20
               TabIndex        =   12
               Top             =   1320
               Width           =   2415
            End
            Begin VB.TextBox TxtApellido2 
               Height          =   285
               Left            =   1680
               MaxLength       =   20
               TabIndex        =   11
               Top             =   1680
               Width           =   2415
            End
            Begin VB.TextBox TxtDireccion 
               Height          =   285
               Left            =   1680
               MaxLength       =   200
               TabIndex        =   10
               Top             =   2040
               Width           =   2415
            End
            Begin VB.TextBox TxtNombre1 
               Height          =   285
               Left            =   1680
               MaxLength       =   20
               TabIndex        =   9
               Top             =   600
               Width           =   2415
            End
            Begin VB.Label Label56 
               Caption         =   "Segundo Apellido:"
               Height          =   255
               Left            =   240
               TabIndex        =   29
               Top             =   1680
               Width           =   1335
            End
            Begin VB.Label Label55 
               Caption         =   "Primer Apellido:"
               Height          =   375
               Left            =   480
               TabIndex        =   28
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label Label54 
               Caption         =   "Segundo Nombre:"
               Height          =   255
               Left            =   240
               TabIndex        =   27
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label Label12 
               Caption         =   "Cargo:"
               Height          =   255
               Left            =   840
               TabIndex        =   26
               Top             =   2760
               Width           =   735
            End
            Begin VB.Label Label10 
               Caption         =   "Depto:"
               Height          =   255
               Left            =   840
               TabIndex        =   25
               Top             =   2400
               Width           =   735
            End
            Begin VB.Label Label6 
               Caption         =   "Sexo:"
               Height          =   255
               Left            =   1080
               TabIndex        =   24
               Top             =   3120
               Width           =   495
            End
            Begin VB.Label Label3 
               Caption         =   "Direccion:"
               Height          =   255
               Left            =   840
               TabIndex        =   23
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label Label2 
               Caption         =   "Primer Nombre:"
               Height          =   255
               Left            =   480
               TabIndex        =   22
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "CodEmpleado:"
               Height          =   255
               Left            =   240
               TabIndex        =   21
               Top             =   240
               Width           =   1695
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Bindings        =   "FrmBajas.frx":09E4
            Height          =   3855
            Left            =   120
            TabIndex        =   69
            Top             =   600
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   6800
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "CodEmpleado1"
            Columns(0).DataField=   "CodEmpleado1"
            Columns(0).DataWidth=   50
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "CodEmpleado"
            Columns(1).DataField=   "CodEmpleado"
            Columns(1).DataWidth=   23
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nombres"
            Columns(2).DataField=   "Nombres"
            Columns(2).DataWidth=   83
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Sexo"
            Columns(3).DataField=   "Sexo"
            Columns(3).DataWidth=   15
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "NumCedula"
            Columns(4).DataField=   "NumCedula"
            Columns(4).DataWidth=   30
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Sindicalista"
            Columns(5).DataField=   "Sindicalista"
            Columns(5).DataWidth=   2
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Activo"
            Columns(6).DataField=   "Activo"
            Columns(6).DataWidth=   10
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2302"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2223"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(10)=   "Column(1)._AlignLeft=0"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=5292"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=5212"
            Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(15)=   "Column(3).Width=2064"
            Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=1984"
            Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(19)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(23)=   "Column(5).Width=1614"
            Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=1535"
            Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(27)=   "Column(6).Width=1402"
            Splits(0)._ColumnProps(28)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(29)=   "Column(6)._WidthInPix=1323"
            Splits(0)._ColumnProps(30)=   "Column(6).Visible=0"
            Splits(0)._ColumnProps(31)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(32)=   "Column(6)._AlignLeft=0"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=244,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(64)  =   "Named:id=33:Normal"
            _StyleDefs(65)  =   ":id=33,.parent=0"
            _StyleDefs(66)  =   "Named:id=34:Heading"
            _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   ":id=34,.wraptext=-1"
            _StyleDefs(69)  =   "Named:id=35:Footing"
            _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   "Named:id=36:Selected"
            _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=37:Caption"
            _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(75)  =   "Named:id=38:HighlightRow"
            _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=39:EvenRow"
            _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(79)  =   "Named:id=40:OddRow"
            _StyleDefs(80)  =   ":id=40,.parent=33"
            _StyleDefs(81)  =   "Named:id=41:RecordSelector"
            _StyleDefs(82)  =   ":id=41,.parent=34"
            _StyleDefs(83)  =   "Named:id=42:FilterBar"
            _StyleDefs(84)  =   ":id=42,.parent=33"
         End
         Begin VB.Frame Frame1 
            Height          =   1815
            Left            =   240
            TabIndex        =   49
            Top             =   4560
            Width           =   8775
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "FrmBajas.frx":09FF
               Left            =   6600
               List            =   "FrmBajas.frx":0A09
               TabIndex        =   89
               Text            =   "Produccion"
               Top             =   1320
               Width           =   1935
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
               Height          =   255
               Left            =   3960
               OleObjectBlob   =   "FrmBajas.frx":0A29
               TabIndex        =   68
               Top             =   840
               Width           =   1335
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
               Height          =   255
               Left            =   3960
               OleObjectBlob   =   "FrmBajas.frx":0A9F
               TabIndex        =   67
               Top             =   360
               Width           =   1335
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmBajas.frx":0B1D
               TabIndex        =   66
               Top             =   1320
               Width           =   1335
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmBajas.frx":0B99
               TabIndex        =   65
               Top             =   840
               Width           =   1455
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
               Height          =   255
               Left            =   2400
               OleObjectBlob   =   "FrmBajas.frx":0C11
               TabIndex        =   64
               Top             =   360
               Width           =   375
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmBajas.frx":0C77
               TabIndex        =   63
               Top             =   360
               Width           =   1215
            End
            Begin VB.CheckBox ChkAntiguedad 
               Caption         =   "Antiguedad Base para Calculo"
               Height          =   255
               Left            =   3720
               TabIndex        =   59
               Top             =   1320
               Width           =   2775
            End
            Begin VB.TextBox TxtDiasTrabajados 
               Height          =   285
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   58
               Top             =   360
               Width           =   975
            End
            Begin VB.Frame Frame6 
               Caption         =   "Tipo de Baja"
               Height          =   615
               Left            =   2040
               TabIndex        =   54
               Top             =   2280
               Visible         =   0   'False
               Width           =   3855
               Begin VB.OptionButton OptDespido 
                  Caption         =   "Despido"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   57
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.OptionButton OptRenuncia 
                  Caption         =   "Renuncia"
                  Height          =   255
                  Left            =   1080
                  TabIndex        =   56
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.OptionButton OptFinContrato 
                  Caption         =   "Final. Contrato"
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   55
                  Top             =   240
                  Width           =   1335
               End
            End
            Begin VB.TextBox TxtMeses 
               Height          =   285
               Left            =   5400
               Locked          =   -1  'True
               TabIndex        =   53
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox TxtAnnos 
               Height          =   285
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   52
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox TxtFechaContrato 
               Height          =   300
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   51
               Top             =   840
               Width           =   1815
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmBajas.frx":0CF3
               TabIndex        =   50
               Top             =   360
               Width           =   1215
            End
            Begin MSComCtl2.DTPicker TxtUltFechaNomina 
               Height          =   300
               Left            =   5280
               TabIndex        =   60
               Top             =   840
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   529
               _Version        =   393216
               Format          =   21364737
               CurrentDate     =   38802
            End
            Begin MSComCtl2.DTPicker TxtFechaHistorial 
               Height          =   300
               Left            =   1680
               TabIndex        =   61
               Top             =   1320
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   529
               _Version        =   393216
               Format          =   21364737
               CurrentDate     =   38802
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
               Left            =   6240
               TabIndex        =   62
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.TextBox TxtDias 
            Height          =   285
            Left            =   2280
            TabIndex        =   44
            Top             =   6600
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox TxtMotivo 
            Height          =   615
            Left            =   5280
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   6720
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.Frame Frame4 
            Caption         =   "Deducciones"
            Height          =   3375
            Left            =   -70440
            TabIndex        =   40
            Top             =   480
            Width           =   4335
            Begin VB.CheckBox ChkDeducciones 
               Caption         =   "Deducciones"
               Height          =   255
               Left            =   360
               TabIndex        =   42
               Top             =   840
               Width           =   2055
            End
            Begin VB.CheckBox ChkPrestamo 
               Caption         =   "Prestamos"
               Height          =   255
               Left            =   360
               TabIndex        =   41
               Top             =   480
               Width           =   2295
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Prestaciones"
            Height          =   3375
            Left            =   -74880
            TabIndex        =   30
            Top             =   480
            Width           =   4455
            Begin VB.TextBox TxtMontoOtrPrestacion 
               Height          =   285
               Left            =   1920
               TabIndex        =   39
               Top             =   2040
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.TextBox TxtOtrPrestacion 
               Height          =   285
               Left            =   960
               TabIndex        =   38
               Top             =   1680
               Visible         =   0   'False
               Width           =   2535
            End
            Begin VB.TextBox TxtDescuentoDias 
               Height          =   285
               Left            =   2400
               TabIndex        =   37
               Top             =   1320
               Width           =   855
            End
            Begin VB.CheckBox ChkOtro 
               Caption         =   "Otros Ingresos"
               Height          =   255
               Left            =   1800
               TabIndex        =   36
               Top             =   840
               Width           =   1335
            End
            Begin VB.CheckBox ChkCargo 
               Caption         =   "Viaticos"
               Height          =   195
               Left            =   240
               TabIndex        =   35
               Top             =   840
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.CheckBox ChkExtra 
               Caption         =   "Horas Extra"
               Height          =   255
               Left            =   1800
               TabIndex        =   34
               Top             =   600
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.CheckBox ChkAntigue 
               Caption         =   "Antiguedad"
               Height          =   195
               Left            =   240
               TabIndex        =   33
               Top             =   600
               Value           =   1  'Checked
               Width           =   1455
            End
            Begin VB.CheckBox ChkVaca 
               Caption         =   "Vacaciones"
               Height          =   255
               Left            =   1800
               TabIndex        =   32
               Top             =   360
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.CheckBox Chk13mes 
               Caption         =   "13vo Mes"
               Height          =   255
               Left            =   240
               TabIndex        =   31
               Top             =   360
               Value           =   1  'Checked
               Width           =   1455
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGridSalarios 
            Bindings        =   "FrmBajas.frx":0D51
            Height          =   2175
            Left            =   -74880
            TabIndex        =   45
            Top             =   480
            Width           =   9015
            _ExtentX        =   15901
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
            Columns(1).Caption=   "Produccion"
            Columns(1).DataField=   "Destajo"
            Columns(1).NumberFormat=   "###,###0.00"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Septimo Dias"
            Columns(2).DataField=   "Septimo"
            Columns(2).NumberFormat=   "###,###0.00"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Otros Ingresos"
            Columns(3).DataField=   "Otros"
            Columns(3).NumberFormat=   "###,###0.00"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Total Ingresos"
            Columns(4).DataField=   "TotalIngresos"
            Columns(4).NumberFormat=   "###,###0.00"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "MES"
            Columns(5).DataField=   "MES"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "AÑO"
            Columns(6).DataField=   "AÑO"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2302"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2223"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=2"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2302"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2223"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=2"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2302"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2223"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=2"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2302"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2223"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=2"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=2"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=1402"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1323"
            Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=1"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=1402"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1323"
            Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=1"
            Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=1"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(64)  =   "Named:id=33:Normal"
            _StyleDefs(65)  =   ":id=33,.parent=0"
            _StyleDefs(66)  =   "Named:id=34:Heading"
            _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   ":id=34,.wraptext=-1"
            _StyleDefs(69)  =   "Named:id=35:Footing"
            _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   "Named:id=36:Selected"
            _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=37:Caption"
            _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(75)  =   "Named:id=38:HighlightRow"
            _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=39:EvenRow"
            _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(79)  =   "Named:id=40:OddRow"
            _StyleDefs(80)  =   ":id=40,.parent=33"
            _StyleDefs(81)  =   "Named:id=41:RecordSelector"
            _StyleDefs(82)  =   ":id=41,.parent=34"
            _StyleDefs(83)  =   "Named:id=42:FilterBar"
            _StyleDefs(84)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label11 
            Caption         =   "Días Trabajados"
            Height          =   255
            Left            =   840
            TabIndex        =   48
            Top             =   6600
            Visible         =   0   'False
            Width           =   1335
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
            Left            =   2760
            TabIndex        =   47
            Top             =   6720
            Visible         =   0   'False
            Width           =   2055
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
            Left            =   5280
            TabIndex        =   46
            Top             =   6960
            Visible         =   0   'False
            Width           =   3375
         End
      End
   End
   Begin Threed.SSCommand CmdAcercade 
      Height          =   525
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
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
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FrmBajas.frx":0D6B
      Top             =   0
   End
End
Attribute VB_Name = "FrmBajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChkConfi_Click()

End Sub

Private Sub ChkAntiguedad_Click()
Dim FechaEgreso As Date, FechaContrato As Date, Año As Integer, Mes As Integer, I As Integer
Dim FechaBusqueda As Date, TotalSalario As Double, SalarioPromedio As Double, Contador As Integer
Dim SQLSalarios As String, SalarioAlto As Double, Salario As Double, FechaHistorico As Date, NumeroEmpleado As Integer
FechaEgreso = Me.TxtUltFechaNomina.Value
FechaContrato = Me.TxtFechaContrato.Text
'//////////SUMO 1 PARA AJUSTAR QUE SIEMPRE DA 1 DIA MENOS//////
annos = CDbl(FechaEgreso) - CDbl(FechaContrato) + 1
TxtAnnos.Text = Format(annos / 365, "###,##0.00")
TxtMeses.Text = Format(annos / 30.41, "###,##0.00")
Me.TxtDiasTrabajados.Text = Format(annos, "###,##0")
Dias = annos
Me.CmdEfectuar.Enabled = False

'///////////Busco la Fecha para la Busqueda////////////////////////////

NumeroEmpleado = Me.TxtCodEmpleado.Text

SQLSalarios = "SELECT DISTINCT" & vbLf
SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
SQLSalarios = SQLSalarios & "AS Destajo, SUM(dbo.DetalleNomina.Incentivos) AS Incentivos, SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo)" & vbLf
SQLSalarios = SQLSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes AS MES," & vbLf
SQLSalarios = SQLSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
SQLSalarios = SQLSalarios & "FROM   dbo.DetalleNomina INNER JOIN" & vbLf
SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
SQLSalarios = SQLSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
SQLSalarios = SQLSalarios & "Having (dbo.DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") And (Sum(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0)"
Me.DtaConsulta.RecordSource = SQLSalarios
Me.DtaConsulta.Refresh
Me.DtaConsulta.Recordset.MoveLast
I = 0
Do While Not Me.DtaConsulta.Recordset.BOF
  If I = 1 Then
    FechaHistorico = Me.DtaConsulta.Recordset("FechaFin")

  ElseIf I = 6 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    Exit Do
  ElseIf I = 0 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    FechaHistorico = Me.DtaConsulta.Recordset("FechaFin")
  Else
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
  End If
  I = I + 1

  Me.DtaConsulta.Recordset.MovePrevious
Loop


FechaEgreso = Me.TxtUltFechaNomina.Value
'FechaHistorico = DateSerial(Year(FechaEgreso), Month(FechaEgreso), 1 - 1)
FechaContrato = Me.TxtFechaContrato.Text
'FechaBusqueda = DateSerial(Year(FechaEgreso), Month(FechaEgreso) - 6, 1)
Año = Year(FechaBusqueda)
Mes = Month(FechaBusqueda)

    SQLSalarios = "SELECT DISTINCT" & vbLf
    SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
    SQLSalarios = SQLSalarios & "AS Destajo, SUM(dbo.DetalleNomina.SeptimoDia) AS Septimo, SUM(dbo.DetalleNomina.OtrosIngresos) AS Otros, SUM(dbo.DetalleNomina.Incentivos)" & vbLf
    SQLSalarios = SQLSalarios & "AS Incentivos," & vbLf
    SQLSalarios = SQLSalarios & "SUM (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.OtrosIngresos)" & vbLf
    SQLSalarios = SQLSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes," & vbLf
    SQLSalarios = SQLSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
    SQLSalarios = SQLSalarios & "FROM    dbo.DetalleNomina INNER JOIN" & vbLf
    SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
    SQLSalarios = SQLSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
    SQLSalarios = SQLSalarios & "HAVING(SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0) And (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND" & vbLf
    SQLSalarios = SQLSalarios & "'" & Format(FechaHistorico, "yyyymmdd") & "')" & vbLf
    SQLSalarios = SQLSalarios & "ORDER BY dbo.Nomina.Ano, dbo.Nomina.Mes"

Me.AdoSalarios.RecordSource = SQLSalarios
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
         Me.TxtAntiguedad.Text = Me.AdoAntiguedad.Recordset("porcent")
        Else
         Me.TxtAntiguedad.Text = 0
        End If
        SalarioPromedio = SueldoPeriodo * PAntiguedad
        SalarioAlto = SueldoPeriodo * PAntiguedad
         
       Else
        SalarioPromedio = SueldoPeriodo
        SalarioAlto = SueldoPeriodo
         Me.TxtAntiguedad.Text = 0
       End If
 
  
Else
    Me.TxtAntiguedad.Text = 0
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

 SQL = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento," & vbLf
 SQL = SQL & "Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.OtrosIngresos, Empleado.DescripOtrIngre," & vbLf
 SQL = SQL & "Empleado.ExentoIr, Empleado.PagoInssPatronal, Empleado.Activo, Empleado.Liquidado, Empleado.Ausente, Empleado.SalarioFijo," & vbLf
 SQL = SQL & "Empleado.SumarSubsidio , Empleado.PorcientoIncentivo, Empleado.Gravidez, TipoNomina.Periodo" & vbLf
 SQL = SQL & "FROM Empleado INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina" & vbLf
 SQL = SQL & "WHERE     (Empleado.CodEmpleado = '" & CodigoEmpleado & "') AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
 Me.DtaConsulta.RecordSource = SQL
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

Private Sub CmdCalcular_Click()
Dim Sueldo As Double
Dim Fecha As String
Dim I As Integer, Espacio As String
Dim SalMayor As Double, Año As Integer
Dim SalTemp As Double, Meses As Integer
Dim SalBrutoTemp As Double, j As Integer
Dim SalBrutoMayor As Double, H As Integer
Dim Mes As Byte, SqlHrsExtras As String
Dim DiaMes As Double, HE As Integer
Dim DiaSemana As Double, DiasAntiguedad As Integer
Dim Mes13 As Double
Dim Vacaciones As Double, SalarioMensual As Double
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
For I = 0 To 5
  
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
            
  Next I




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
   DiasDescuento = Val(Me.TxtDescuentoDias.Text)
   Vacaciones = SalMayor * (((CantRegistros * 1.25) - DiasDescuento) / DiasMes)
Else
   'Dias = 0
   Vacaciones = 0
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
       Otro = Val(TxtMontoOtrPrestacion.Text)
    End If
Else
    TextOtro = "Ninguna"
    Otro = 0
End If

'////////////////////////////////////////////////////////////////////////////
'////////Hago el Calculo Proporcional de los dias trabajados////////////////
'/////////////////////////////////////////////////////////////////////////////

If Val(TxtDias.Text) > 0 Then
MontoNomPropor = (SalarioBasico / DiasMes) * Val(TxtDias.Text)
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
        DtaPrestamo.RecordSource = SQlPrestamo
        DtaPrestamo.Refresh
Prestamo = 0
Do While Not DtaPrestamo.Recordset.EOF
   Prestamo = DtaPrestamo.Recordset("CuotaIgual") + Prestamo
 Me.DtaPrestamo.Recordset.MoveNext
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
         Deducciones = Deducciones + Val(Me.DtaDeducciones.Recordset("valor"))
     
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
SalarioMensual = MontoNomPropor + Vacaciones + MontoHRSExtras + Otro

         
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

MontoIR = 0
MontoIRPatronal = 0

'Hago el Calcul del nuevo Techo para el Ir
MontoBrutoMensual = SalarioMensual - MontoInss

        'agregar IR laboral y patronal
        MontoIR = 0
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
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIR = Format(MontoIR / CantSabados / 12, "###,##0.00")
               MontoIRPatronal = MontoIR
               Exit Do
            End If
            End If
            
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIR = Format(MontoIR / CantSabados / 12, "###,##0.00")
               MontoIRPatronal = MontoIR
               Exit Do
                       
            End If
            End If
            
        ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Viernes" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
  '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
 
                MontoIrMensual = Format(MontoIR / 1 / 12, "###,##0.00")
                MontoIR = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior

            End If
            Else
               MontoIrMensual = 0
               MontoIR = MontoIrMensual - MontoIrAnterior
               MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
            End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
  '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
 
                MontoIrMensual = Format(MontoIR / 1 / 12, "###,##0.00")
                MontoIR = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior

            End If
            Else
               MontoIrMensual = 0
                MontoIR = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
            End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'///////Verfico si el la Ultima Quincena para hacer ajustes////////////

                MontoIrMensual = Format(MontoIR / 1 / 12, "###,##0.00")
                MontoIR = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior

            End If
            Else
               MontoIrMensual = 0
                MontoIR = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
            End If
         
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Mensual" Then
           If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIR = Format(MontoIR / 12, "###,##0.00")
               MontoIRPatronal = MontoIR
               Exit Do
            End If
         End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
           If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIR = Format(MontoIR / 4, "###,##0.00")
               MontoIRPatronal = MontoIR
               Exit Do
            End If
           End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
             If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIR = Format(MontoIR / 2, "###,##0.00")
               MontoIRPatronal = MontoIR
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
 ID = 0
Else
 Me.DtaBajas.Recordset.MoveLast
 ID = Me.DtaBajas.Recordset("id") + 1
End If

Me.DtaBajas.RecordSource = "SELECT Id,  AnnosTrabajados,CodEmpleado, FechaBaja, MesesTrabajados, DiasTrabajados, MontoNomPropor, MontoVaca, Monto13Mes, MontoAnosTrab, MontoCargoConfianza,MontoAntiguedad, TipoBaja, MotivoBaja, Otro, MontoOtro, Prestamo, Deducciones, SalarioMensual, MontoINSS, MontoIR, FechaIniAgui, FechaFinAgui,DiasAguinaldo , DiasVacaciones, DiasMenosVaca, FechaIniVaca, FechaFinVaca, HorasExtra, Viaticos, MontoHorasExtra From Bajas WHERE     (CodEmpleado = '" & CodEmpleado & "')"
Me.DtaBajas.Refresh

'////////////////////////Sumo el total de los ingresos y Egresoso//////////////////////////
TotalIngresos = MontoNomPropor + Vacaciones + Mes13 + MontoAntiguedad + MontoHRSExtras + Otro
TotalEgresos = Deducciones + Prestamo + MontoInss + MontoIR
If Me.DtaBajas.Recordset.EOF Then
'////////Agrego un nuevo Registro////////////////
 DtaBajas.Recordset.AddNew
    DtaBajas.Recordset("id") = ID
    DtaBajas.Recordset("CodEmpleado") = CodEmpleado
    DtaBajas.Recordset("fechabaja") = Format(Now, "dd/mm/yyyy")
    DtaBajas.Recordset("DiasTrabajados") = Val(TxtDias.Text)
    DtaBajas.Recordset("MontoNomPropor") = MontoNomPropor
    DtaBajas.Recordset("annostrabajados") = Val(Me.TxtAnnos.Text)
    DtaBajas.Recordset("mesestrabajados") = Val(Me.TxtMeses.Text)
    DtaBajas.Recordset("DiasTrabajados") = Val(TxtDias.Text)
    DtaBajas.Recordset("montovaca") = Format(Vacaciones, "##,##0.00")
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
    DtaBajas.Recordset("MontoIR") = MontoIR
DtaBajas.Recordset.Update

Else
'/////////////Edito el registro existente.///////////////
 'Me.DtaBajas.Recordset.Edit
    DtaBajas.Recordset("CodEmpleado") = CodEmpleado
    DtaBajas.Recordset("fechabaja") = Format(Now, "dd/mm/yyyy")
    DtaBajas.Recordset("DiasTrabajados") = Val(TxtDias.Text)
    DtaBajas.Recordset("MontoNomPropor") = MontoNomPropor
    DtaBajas.Recordset("annostrabajados") = Val(Me.TxtAnnos.Text)
    DtaBajas.Recordset("mesestrabajados") = Val(Me.TxtMeses.Text)
    DtaBajas.Recordset("DiasTrabajados") = Val(TxtDias.Text)
    DtaBajas.Recordset("montovaca") = Format(Vacaciones, "##,##0.00")
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
    DtaBajas.Recordset("MontoIR") = MontoIR

 Me.DtaBajas.Recordset.Update
End If







'///////Imprimo la Liquidacion//////////////////////////////////////
k% = MsgBox("Desea Imprimir las bajas?", vbYesNo)
If k% = 6 Then

    CodEmpleado = TxtCodEmpleado.Text
    Me.DtaConsulta.RecordSource = "SELECT [Bajas].[MontoNomPropor]+[Bajas].[MontoVaca]+[Bajas].[Monto13Mes]+[Bajas].[MontoAnosTrab]+[Bajas].[MontoCargoConfianza]+[Bajas].[MontoAntiguedad]+[Bajas].[MontoOtro] AS Ingresos, [Bajas].[Prestamo]+[Bajas].[Deducciones]+[Bajas].[MontoINSS]+[Bajas].[MontoIR] AS Egresos, Bajas.CodEmpleado From Bajas Where (((Bajas.CodEmpleado) = '" & CodEmpleado & "'))"
    Me.DtaConsulta.Refresh
    
    ArepBajas.DataControl1.ConnectionString = ConexionReporte
    ArepBajas.LblTitulo.Caption = Titulo
    ArepBajas.LblSubtitulo.Caption = Subtitulo
    ArepBajas.ImgLogo.Picture = LoadPicture(RutaLogo)
    cadena = "SELECT Empleado.CodEmpleado, [Empleado].[Nombre1]+'" & Espacio & "'+[Empleado].[Nombre2]+'" & Espacio & "'+[Empleado].[Apellido1]+'" & Espacio & "'+[Empleado].[Apellido2] AS Nombres, Bajas.FechaBaja, Bajas.AnnosTrabajados, Bajas.MesesTrabajados, Bajas.DiasTrabajados, Bajas.MontoNomPropor, Bajas.MontoVaca, Bajas.Monto13Mes, Bajas.MontoAnosTrab, Bajas.MontoCargoConfianza, Bajas.MontoAntiguedad, Bajas.MotivoBaja, Bajas.TipoBaja, Bajas.Otro, Bajas.MontoOtro, Bajas.Prestamo, Bajas.Deducciones, Bajas.MontoINSS, Bajas.MontoIR, Cargo.Cargo, Departamento.Departamento, Historico.FechaContrato, Bajas.SalarioMensual,Bajas.HorasExtra, Bajas.Viaticos " & vbLf
    cadena = cadena & ",Bajas.MontoHorasExtra FROM ((Departamento INNER JOIN (Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) ON Departamento.CodDepartamento = Empleado.CodDepartamento) INNER JOIN Bajas ON Empleado.CodEmpleado = Bajas.CodEmpleado) LEFT JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado Where (((Empleado.CodEmpleado) = '" & CodEmpleado & "'))"
    ArepBajas.DataControl1.Source = cadena
'    ArepBajas.LblFechaFinVaca = FechaFinVaca
'    ArepBajas.LblFechaIniVaca = FechaIniVaca
'    ArepBajas.LblFechaIniAguinaldo = FechaIniAgui
'    ArepBajas.LblFechaFinAguinaldo = FechaFinAgui
'    ArepBajas.LblDiasAguinaldo = Format(Dias, "##,##0.00")
'    ArepBajas.LblDiasBruto = CantRegistros * 1.25
'    ArepBajas.LblDiasMenos = DiasDescuento
'    ArepBajas.LblDiasNetos = (CantRegistros * 1.25) - DiasDescuento
    
  
     NetoPagar = TotalIngresos - TotalEgresos
'     ArepBajas.LblTotalEgresos.Caption = Format(TotalEgresos, "##,##0.00")
'     ArepBajas.LblTotalIngresos.Caption = Format(TotalIngresos, "##,##0.00")
'     ArepBajas.LblNetoPagar.Caption = Format(NetoPagar, "##,##0.00")
    

    ArepBajas.Show 1
    Me.CmdEfectuar.Enabled = True
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
Quien = "Despido"
FrmBuscaEmpleado.Show 1
End Sub

Private Sub CmdCalculos_Click()
Dim SalarioHoraAgui As Double, SalarioDiarioAgui As Double, SalarioMesAgui As Double, MontoIR As Double, MontoIRPatronal As Double
Dim TipoNomina As String, SQL As String, CodigoEmpleado As String, CodTipoNomina As String, MontoBrutoMensual
Dim Años As Double, PAntiguedad As Double, AñoActual As Integer, FechaInicio As Date, Prestamo As Double
Dim DiasTrabajados As Double, FechaEgreso As Date, TotalAguinaldo As Double, TotalVacaciones As Double
Dim DiasAntiguedad As Integer, TotalAntiguedad As Double, TotalInss As Double, TotalIr As Double
Dim MontoHRSExtras As Double, HE As Double, Adelanto13vo As Double, AdelantoVaca As Double, TotalOtrosSalarios As Double
Dim SalarioHoraVaca As Double, SalarioDiarioVaca As Double, SalarioMesVaca As Double
Dim DiasTrabajadosVaca As Double, DiasTrabajadosAgui As Double, FechaInicioVaca As Date, FechaInicioAgui As Date
Dim DiasMes As Double, I As Integer, CantidadEmpleados As Double, CodEmpleado As Double, CodTiposNomina As String
Dim FechaInicioContrato As Date

If Me.Combo1.Text = "Administracion" Then
  CodTiposNomina = "01"
 Me.TxtFechaHistorial.Value = "31/08/2007"
 Me.TxtUltFechaNomina.Value = "05/09/2007"
Else
  CodTiposNomina = "02"
 Me.TxtFechaHistorial.Value = "20/08/2007"
 Me.TxtUltFechaNomina.Value = "05/09/2007"
End If


Me.AdoEmpleados.RecordSource = "SELECT CodEmpleado1, CodEmpleado, Nombre1 + N' ' + Nombre2 + N' ' + Apellido1 + N' ' + Apellido2 AS Nombres, Sexo, NumCedula, Sindicalista, Activo From Empleado Where (Activo = 1) AND (CodTipoNomina = '" & CodTiposNomina & "') ORDER BY CodEmpleado1"
Me.AdoEmpleados.Refresh
Me.AdoEmpleados.Recordset.MoveLast
CantidadEmpleados = Me.AdoEmpleados.Recordset.RecordCount
Me.osProgress.Min = 1
Me.osProgress.Value = 0
Me.osProgress.Max = CantidadEmpleados
Me.osProgress.Visible = True
I = 1
Me.AdoEmpleados.Recordset.MoveFirst
Do While Not Me.AdoEmpleados.Recordset.EOF
  Me.osProgress.Value = I
  I = I + 1
  DoEvents
  

  
  
  
  
If Me.Combo1.Text = "Administracion" Then
  CodTiposNomina = "01"
 Me.TxtFechaHistorial.Value = "31/08/2007"
 Me.TxtUltFechaNomina.Value = "05/09/2007"
Else
  CodTiposNomina = "02"
 Me.TxtFechaHistorial.Value = "20/08/2007"
 Me.TxtUltFechaNomina.Value = "05/09/2007"
End If

CodigoEmpleado = Me.TDBGrid1.Columns(1).Text
CodEmpleado = Me.TDBGrid1.Columns(1).Text
  Me.TxtCodEmpleado1.Text = Me.TDBGrid1.Columns(0).Text
'  Me.TxtCodEmpleado.Text = Me.TDBGrid1.Columns(1).Text

 If CodigoEmpleado = "4900" Then
  cod = 1
 End If


'/////////CONSULTA EL SALARIO Y TIPO DE NOMINA DEL EMPLEADO//////////////////////////

 SQL = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento," & vbLf
 SQL = SQL & "Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.OtrosIngresos, Empleado.DescripOtrIngre," & vbLf
 SQL = SQL & "Empleado.ExentoIr, Empleado.PagoInssPatronal, Empleado.Activo, Empleado.Liquidado, Empleado.Ausente, Empleado.SalarioFijo," & vbLf
 SQL = SQL & "Empleado.SumarSubsidio , Empleado.PorcientoIncentivo, Empleado.Gravidez, TipoNomina.Periodo" & vbLf
 SQL = SQL & "FROM Empleado INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina" & vbLf
 SQL = SQL & "WHERE     (Empleado.CodEmpleado = '" & CodigoEmpleado & "') AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
 Me.DtaConsulta.RecordSource = SQL
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
  TipoNomina = Me.DtaConsulta.Recordset("Periodo")
  CodTipoNomina = Me.DtaConsulta.Recordset("CodTipoNomina")
 Else
  MsgBox "Este Empleado no Existe", vbCritical, "Sistema de Nominas"
  Exit Sub
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
        End If
'        SalarioMes = (Me.DtaConsulta.Recordset("SueldoPeriodo") * 2) * PAntiguedad

        Me.DtaControles.Refresh
        If Not Me.DtaControles.Recordset.EOF Then
         DiasMes = Me.DtaControles.Recordset("DiasMes")
        End If
        
        If Val(Me.TxtSalarioAlto.Text) > Val(Me.TxtSalarioBasico.Text) Then
          SalarioMesAgui = Me.TxtSalarioAlto.Text
          SalarioHoraAgui = (SalarioMesAgui / DiasMes) / 8
          SalarioDiarioAgui = SalarioMesAgui / DiasMes
        Else
          SalarioMesAgui = Me.TxtSalarioBasico.Text
          SalarioHoraAgui = (SalarioMesAgui / DiasMes) / 8
          SalarioDiarioAgui = SalarioMesAgui / DiasMes
        End If
        If Val(Me.TxtSalarioPromedio.Text) > Val(Me.TxtSalarioBasico.Text) Then
         SalarioMesVaca = Me.TxtSalarioPromedio.Text
        Else
         SalarioMesVaca = Me.TxtSalarioBasico.Text
        End If
        SalarioHoraVaca = (SalarioMesVaca / DiasMes) / 8
        SalarioDiarioVaca = SalarioMesVaca / DiasMes
       
       Else
       
        Me.DtaControles.Refresh
        If Not Me.DtaControles.Recordset.EOF Then
         DiasMes = Me.DtaControles.Recordset("DiasMes")
        End If
'        SalarioMes = Me.DtaConsulta.Recordset("SueldoPeriodo") * 2
        If Val(Me.TxtSalarioAlto.Text) > Val(Me.TxtSalarioBasico.Text) Then
          SalarioMesAgui = Me.TxtSalarioAlto.Text
          SalarioHoraAgui = (SalarioMesAgui / DiasMes) / 8
          SalarioDiarioAgui = SalarioMesAgui / DiasMes
        Else
          SalarioMesAgui = Me.TxtSalarioBasico.Text
          SalarioHoraAgui = (SalarioMesAgui / DiasMes) / 8
          SalarioDiarioAgui = SalarioMesAgui / DiasMes
        End If
        
        If Val(Me.TxtSalarioPromedio.Text) > Val(Me.TxtSalarioBasico.Text) Then
         SalarioMesVaca = Me.TxtSalarioPromedio.Text
        Else
         SalarioMesVaca = Me.TxtSalarioBasico.Text
        End If
        SalarioHoraVaca = (SalarioMesVaca / DiasMes) / 8
        SalarioDiarioVaca = SalarioMesVaca / DiasMes
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
       FechaInicioAgui = Me.DTPFechaIniAgui
       FechaInicioVaca = Me.DTPFechaIniVaca
       
       
       FechaInicioContrato = Me.TxtFechaContrato.Text
       
       If CDbl(FechaInicioAgui) < CDbl(FechaInicioContrato) Then
       DiasTrabajados = CDbl(FechaEgreso) - CDbl(FechaInicio) + Val(Me.TxtDias.Text) + 1
       DiasTrabajadosAgui = CDbl(FechaEgreso) - CDbl(FechaInicioContrato) + 1
       Else
       DiasTrabajados = CDbl(FechaEgreso) - CDbl(FechaInicio) + Val(Me.TxtDias.Text) + 1
       DiasTrabajadosAgui = CDbl(FechaEgreso) - CDbl(FechaInicioAgui) + 1
       End If
       
       If CDbl(FechaInicioVaca) < CDbl(FechaInicioContrato) Then
       DiasTrabajados = CDbl(FechaEgreso) - CDbl(FechaInicio) + Val(Me.TxtDias.Text) + 1
'       DiasTrabajadosVaca = CDbl(FechaEgreso) - CDbl(FechaInicioVaca) + 1
       DiasTrabajadosVaca = CDbl(FechaEgreso) - CDbl(FechaInicioContrato) + 1
       Else
       DiasTrabajados = CDbl(FechaEgreso) - CDbl(FechaInicio) + Val(Me.TxtDias.Text) + 1
'       DiasTrabajadosVaca = CDbl(FechaEgreso) - CDbl(FechaInicioVaca) + 1
       DiasTrabajadosVaca = CDbl(FechaEgreso) - CDbl(FechaInicioVaca) + 1
       End If
       

       '////////////////////////////////////////////////////////////////////////
       '////////////////CALCULO AGUINALDO Y VACACIONES//////////////////////////
       '////////////////////////////////////////////////////////////////////////
       'aguinaldo= Diastrabajado*1/12*salariodiario
       
       
       If Me.Chk13mes.Value = 1 Then
         TotalAguinaldo = DiasTrabajadosAgui * 0.083333 * SalarioDiarioAgui
         
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
        DiasTrabajadosVaca = DiasTrabajadosVaca - Val(Me.TxtDescuentoDias.Text)
        TotalVacaciones = DiasTrabajadosVaca * 0.083333 * SalarioDiarioVaca
       
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
         
       
       '////////////////////////////////////////////////////////////////////////////////
       '//////////////CALCULO ANTIGUEDAD/////////////////////////////////////////////////
       '/////////////////////////////////////////////////////////////////////////////////
       DiasAntiguedad = Me.TxtDiasTrabajados.Text
       '/////////////VERIFICO SI SE UTILIZA LA ANTIGUEDAD COMO BASE//////////////////////
       '/////////////PARA EL CALCULO DE LA LIQUIDACION///////////////////////////////////
       
        If DiasAntiguedad >= 365 Then
           TotalAntiguedad = DiasAntiguedad * 0.083333 * SalarioDiarioVaca
        Else
            TotalAntiguedad = 0
        End If
 
       
      '//////////////////////////////////////////////////////////////////////////
      '/////////CALCULO LOS OTROS INGRESOS//////////////////////////////////////
      If Me.ChkOtro.Value = 1 Then
        TotalOtrosSalarios = Val(Me.TxtMontoOtrPrestacion.Text)
      Else
        TotalOtrosSalarios = 0
      
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
        DtaPrestamo.RecordSource = SQlPrestamo
        DtaPrestamo.Refresh
        Prestamo = 0
        Do While Not DtaPrestamo.Recordset.EOF
        Prestamo = DtaPrestamo.Recordset("CuotaIgual") + Prestamo
        Me.DtaPrestamo.Recordset.MoveNext
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
            DtaDeduccion.Recordset.MoveNext
            Loop
        End If
      
      
      
             
       '//////////////CALCULO DEL INSS///////////////////////////////////////
       TotalInss = (TotalVacaciones + MontoHRSExtras + TotalOtrosSalarios) * 0.0625
       
       '////////////////////////////////////////////////////////////////////////////////////
       '////////////////////////CALCULO DEL IR//////////////////////////////////////////////
       '////////////////////////////////////////////////////////////////////////////////////

        MontoIR = 0
        MontoIRPatronal = 0

        'Hago el Calcul del nuevo Techo para el Ir
         MontoBrutoMensual = (TotalVacaciones + TotalAntiguedad + MontoHRSExtras + TotalOtrosSalarios) - TotalInss

        'agregar IR laboral y patronal
        MontoIR = 0
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
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIR = Format(MontoIR / CantSabados / 12, "###,##0.00")
               MontoIRPatronal = MontoIR
               Exit Do
            End If
            End If
            
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIR = Format(MontoIR / CantSabados / 12, "###,##0.00")
               MontoIRPatronal = MontoIR
               Exit Do
                       
            End If
            End If
            
        ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Viernes" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
  '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
 
                MontoIrMensual = Format(MontoIR / 1 / 12, "###,##0.00")
                MontoIR = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior

            End If
            Else
               MontoIrMensual = 0
               MontoIR = MontoIrMensual - MontoIrAnterior
               MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
            End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
  '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
 
                MontoIrMensual = Format(MontoIR / 1 / 12, "###,##0.00")
                MontoIR = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior

            End If
            Else
               MontoIrMensual = 0
                MontoIR = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
            End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
            If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
'///////Verfico si el la Ultima Quincena para hacer ajustes////////////

                MontoIrMensual = Format(MontoIR / 1 / 12, "###,##0.00")
                MontoIR = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior

            End If
            Else
               MontoIrMensual = 0
                MontoIR = MontoIrMensual - MontoIrAnterior
                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
            End If
         
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Mensual" Then
           If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIR = Format(MontoIR / 12, "###,##0.00")
               MontoIRPatronal = MontoIR
               Exit Do
            End If
         End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
           If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIR = Format(MontoIR / 4, "###,##0.00")
               MontoIRPatronal = MontoIR
               Exit Do
            End If
           End If
         ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
             If (MontoBrutoMensual) >= MinIR Then
            If DtaIR.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIR.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
               MontoIR = ((MontoBrutoMensual * 12) - DtaIR.Recordset("SobreExceso")) * (DtaIR.Recordset("PorcientoImpuesto") / 100) + DtaIR.Recordset("ImpuestoBase")
               MontoIR = Format(MontoIR / 2, "###,##0.00")
               MontoIRPatronal = MontoIR
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
            ID = 0
        Else
            Me.DtaBajas.Recordset.MoveLast
            ID = Me.DtaBajas.Recordset("id") + 1
        End If

        Me.DtaBajas.RecordSource = "SELECT Id,  AnnosTrabajados,CodEmpleado, FechaBaja, MesesTrabajados, DiasTrabajados, MontoNomPropor, MontoVaca, Monto13Mes, MontoAnosTrab, MontoCargoConfianza,MontoAntiguedad, TipoBaja, MotivoBaja, Otro, MontoOtro, Prestamo, Deducciones, SalarioMensual, MontoINSS, MontoIR, FechaIniAgui, FechaFinAgui,DiasAguinaldo , DiasVacaciones, DiasMenosVaca, FechaIniVaca, FechaFinVaca, HorasExtra, Viaticos, MontoHorasExtra,SalarioPromedio, SalarioBasico, TotalIngresos, TotalEgresos, TotalPagar,DiasTrabajadosAgui, DiasTrabajadosVaca  From Bajas WHERE     (CodEmpleado = '" & CodEmpleado & "')"
        Me.DtaBajas.Refresh

        '////////////////////////Sumo el total de los ingresos y Egresoso//////////////////////////
        TotalIngresos = TotalVacaciones + TotalAguinaldo + TotalAntiguedad
        TotalEgresos = Deducciones + Prestamo + TotalInss + MontoIR
     If Me.DtaBajas.Recordset.EOF Then
        '////////Agrego un nuevo Registro////////////////
        DtaBajas.Recordset.AddNew
        DtaBajas.Recordset("id") = ID
        DtaBajas.Recordset("CodEmpleado") = CodEmpleado
        DtaBajas.Recordset("fechabaja") = Me.TxtUltFechaNomina.Value
        DtaBajas.Recordset("MontoNomPropor") = MontoNomPropor
        DtaBajas.Recordset("annostrabajados") = Val(Me.TxtAnnos.Text)
        DtaBajas.Recordset("mesestrabajados") = Val(Me.TxtMeses.Text)
        DtaBajas.Recordset("DiasTrabajados") = Val(Me.TxtDiasTrabajados.Text)
        DtaBajas.Recordset("montovaca") = Format(TotalVacaciones, "##,##0.00")
        DtaBajas.Recordset("monto13mes") = Format(TotalAguinaldo, "##,##0.00")
        DtaBajas.Recordset("MontoAntiguedad") = Format(TotalAntiguedad, "##,##0.00")
        DtaBajas.Recordset("HorasExtra") = HE
        DtaBajas.Recordset("MontoHorasExtra") = MontoHRSExtras
        DtaBajas.Recordset("SalarioBasico") = Me.TxtSalarioBasico.Text
        DtaBajas.Recordset("SalarioPromedio") = Me.TxtSalarioPromedio.Text
        DtaBajas.Recordset("TotalIngresos") = TotalIngresos
        DtaBajas.Recordset("TotalEgresos") = TotalEgresos
        DtaBajas.Recordset("TotalPagar") = TotalIngresos - TotalEgresos
        DtaBajas.Recordset("Otro") = TextOtro
        DtaBajas.Recordset("montootro") = Otro
        
            DtaBajas.Recordset("FechaIniVaca") = CDate(FechaInicioVaca)
            DtaBajas.Recordset("FechaFinVaca") = CDate(Me.TxtUltFechaNomina.Value)


            DtaBajas.Recordset("FechaIniAgui") = FechaInicioAgui


            DtaBajas.Recordset("FechaFinAgui") = Me.TxtUltFechaNomina.Value

            DtaBajas.Recordset("DiasAguinaldo") = Format(DiasTrabajadosAgui * (1 / 12), "##,##0.00")

            DtaBajas.Recordset("DiasVacaciones") = Format(DiasTrabajadosVaca * (1 / 12), "##,##0.00")
            
            DtaBajas.Recordset("DiasTrabajadosAgui") = DiasTrabajadosAgui
                
            DtaBajas.Recordset("DiasTrabajadosVaca") = DiasTrabajadosVaca


            DtaBajas.Recordset("DiasMenosVaca") = Val(Me.TxtDescuentoDias.Text)

            DtaBajas.Recordset("SalarioMensual") = SalarioMesAgui
    
    
    
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
        DtaBajas.Recordset("MontoIR") = MontoIR
        DtaBajas.Recordset.Update

    Else
        '/////////////Edito el registro existente.///////////////
 'Me.DtaBajas.Recordset.Edit
        DtaBajas.Recordset("MontoNomPropor") = MontoNomPropor
        DtaBajas.Recordset("annostrabajados") = Val(Me.TxtAnnos.Text)
        DtaBajas.Recordset("mesestrabajados") = Val(Me.TxtMeses.Text)
        DtaBajas.Recordset("DiasTrabajados") = Me.TxtDiasTrabajados.Text
        DtaBajas.Recordset("montovaca") = Format(TotalVacaciones, "##,##0.00")
        DtaBajas.Recordset("monto13mes") = Format(TotalAguinaldo, "##,##0.00")
        DtaBajas.Recordset("MontoAntiguedad") = Format(TotalAntiguedad, "##,##0.00")
        
        DtaBajas.Recordset("SalarioBasico") = Me.TxtSalarioBasico.Text
        DtaBajas.Recordset("SalarioPromedio") = Me.TxtSalarioPromedio.Text
        DtaBajas.Recordset("TotalIngresos") = TotalIngresos
        DtaBajas.Recordset("TotalEgresos") = TotalEgresos
        DtaBajas.Recordset("TotalPagar") = TotalIngresos - TotalEgresos
        DtaBajas.Recordset("fechabaja") = Me.TxtUltFechaNomina.Value
    'DtaBajas.Recordset.montoCargoConfianza = CargoConfianza
    DtaBajas.Recordset("Otro") = TextOtro
    DtaBajas.Recordset("montootro") = Otro

     DtaBajas.Recordset("FechaIniVaca") = CDate(FechaInicioVaca)

    DtaBajas.Recordset("FechaFinVaca") = CDate(Me.TxtUltFechaNomina.Value)

    DtaBajas.Recordset("FechaIniAgui") = FechaInicioAgui

     DtaBajas.Recordset("FechaFinAgui") = Me.TxtUltFechaNomina.Value

    DtaBajas.Recordset("DiasAguinaldo") = Format(DiasTrabajadosAgui * (1 / 12), "####0.00")

    DtaBajas.Recordset("DiasVacaciones") = Format(DiasTrabajadosVaca * (1 / 12), "####0.00")
    
    DtaBajas.Recordset("DiasTrabajadosAgui") = DiasTrabajadosAgui
    DtaBajas.Recordset("DiasTrabajadosVaca") = DiasTrabajadosVaca

    DtaBajas.Recordset("DiasMenosVaca") = Val(Me.TxtDescuentoDias.Text)

    DtaBajas.Recordset("SalarioMensual") = SalarioMesAgui
    
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
    DtaBajas.Recordset("MontoIR") = MontoIR

 Me.DtaBajas.Recordset.Update
End If


  Me.AdoEmpleados.Recordset.MoveNext
Loop






   
       
  

End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdDetalle_Click()
Dim SQLSalarios As String, FechaEgreso As Date, FechaContrato As Date, FechaHistorico As Date, FechaBusqueda As Date
Dim CodEmpleado As String, NumeroEmpleado As Integer, I As Integer
    CodEmpleado = TxtCodEmpleado.Text
  

  
  '///////////Busco la Fecha para la Busqueda////////////////////////////
  
FechaEgreso = Me.TxtFechaHistorial.Value
FechaContrato = Me.TxtFechaContrato.Text


NumeroEmpleado = Me.TxtCodEmpleado.Text

SQLSalarios = "SELECT DISTINCT" & vbLf
SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
SQLSalarios = SQLSalarios & "AS Destajo, SUM(dbo.DetalleNomina.Incentivos) AS Incentivos, SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo)" & vbLf
SQLSalarios = SQLSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes AS MES," & vbLf
SQLSalarios = SQLSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
SQLSalarios = SQLSalarios & "FROM   dbo.DetalleNomina INNER JOIN" & vbLf
SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
SQLSalarios = SQLSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
SQLSalarios = SQLSalarios & "Having (dbo.DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") And (Sum(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0) AND (MAX(Nomina.FechaNomina) <= CONVERT(DATETIME, '" & Format(FechaEgreso, "yyyy/mm/dd") & "', 102))"
Me.DtaConsulta.RecordSource = SQLSalarios
Me.DtaConsulta.Refresh

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
Me.DtaConsulta.Recordset.MoveLast
I = 0
Do While Not Me.DtaConsulta.Recordset.BOF
  If I = 1 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")

  ElseIf I = 5 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    Exit Do
  ElseIf I = 0 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    FechaHistorico = Me.DtaConsulta.Recordset("FechaFin")
  Else
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
  End If
  I = I + 1

  Me.DtaConsulta.Recordset.MovePrevious
Loop

  
  
  
    ArepHistorialLiquida.DataControl1.ConnectionString = ConexionReporte
    ArepHistorialLiquida.LblTitulo.Caption = Titulo
    ArepHistorialLiquida.LblSubtitulo.Caption = Subtitulo
    ArepHistorialLiquida.ImgLogo.Picture = LoadPicture(RutaLogo)
    FechaEgreso = Me.TxtUltFechaNomina.Value

    FechaContrato = Me.TxtFechaContrato.Text


    Año = Year(FechaBusqueda)
    Mes = Month(FechaBusqueda)
    
    SQLSalarios = "SELECT DISTINCT" & vbLf
    SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, dbo.DetalleNomina.SalarioBasico AS SalarioBasico, dbo.DetalleNomina.Destajo AS Destajo," & vbLf
    SQLSalarios = SQLSalarios & "dbo.DetalleNomina.SeptimoDia AS Septimo, dbo.DetalleNomina.OtrosIngresos AS Otros, dbo.DetalleNomina.Incentivos AS Incentivos," & vbLf
    SQLSalarios = SQLSalarios & "dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.OtrosIngresos AS TotalIngresos," & vbLf
    SQLSalarios = SQLSalarios & "dbo.Nomina.FechaNominaINI AS FechaInicio, dbo.Nomina.FechaNomina AS FechaFin, dbo.Nomina.Mes, dbo.Nomina.Ano AS AÑO" & vbLf
    SQLSalarios = SQLSalarios & "FROM         dbo.DetalleNomina INNER JOIN" & vbLf
    SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
    SQLSalarios = SQLSalarios & "WHERE     (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo <> 0) AND (dbo.DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND" & vbLf
    SQLSalarios = SQLSalarios & "(dbo.Nomina.FechaNomina BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND '" & Format(FechaHistorico, "yyyymmdd") & "')" & vbLf
    SQLSalarios = SQLSalarios & "ORDER BY dbo.Nomina.Ano, dbo.Nomina.Mes,dbo.Nomina.FechaNomina"
'
    
    ArepDetalleLiquida.LblTitulo.Caption = Titulo
    ArepDetalleLiquida.LblSubtitulo.Caption = Subtitulo
    ArepDetalleLiquida.ImgLogo.Picture = LoadPicture(RutaLogo)
    ArepDetalleLiquida.DataControl1.ConnectionString = ConexionReporte
    ArepDetalleLiquida.DataControl1.Source = SQLSalarios
    
    ArepDetalleLiquida.LblCodEmpleado.Caption = Me.TxtCodEmpleado1.Text
    ArepDetalleLiquida.LblNombreEmpleado.Caption = Me.TxtNombre1.Text + " " + Me.TxtNombre2.Text + " " + Me.TxtApellido1.Text + " " + Me.TxtApellido2.Text
    ArepDetalleLiquida.LblDepartamento.Caption = Me.TxtDepartamento.Text
    ArepDetalleLiquida.LblCargo.Caption = Me.TxtCargo.Text
    ArepDetalleLiquida.LblAños.Caption = Me.TxtAnnos.Text
    ArepDetalleLiquida.LblDias.Caption = Me.TxtDiasTrabajados.Text
    ArepDetalleLiquida.LblMeses.Caption = Me.TxtMeses.Text
    
    ArepDetalleLiquida.Show 1

End Sub

Private Sub CmdEfectuar_Click()
On Error GoTo TipoErr
Dim I As Integer, NumNomina As Integer
Dim SalMayor As Double, SQL As String
Dim SalTemp As Double
Dim SalBrutoTemp As Double
Dim SalBrutoMayor As Double
Dim Mes As Byte
Dim DiaMes As Double
Dim DiaSemana As Double
Dim Mes13 As Double
Dim Vacaciones As Double
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
Me.DtaPrestamo.RecordSource = "SELECT Prestamo.CodEmpleado, Prestamo.Cancelado From Prestamo Where (((Prestamo.CodEmpleado) = " & CodigoEmpleado & "))"
DtaPrestamo.Refresh
    Prestamo = 0
Do While Not DtaPrestamo.Recordset.EOF

   'DtaPrestamo.Recordset.Edit
   DtaPrestamo.Recordset("cancelado") = True
   DtaPrestamo.Recordset.Update


 Me.DtaPrestamo.Recordset.MoveNext
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
DtaDeduccion.Recordset.MoveNext
Loop

End If

'/////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////BUSCO SI TIENE NOMINAS ACTIVAS PARA BORRAR EL REGISTRO//////
'////////////////////////////////////////////////////////////////////////////////////////

SQL = "SELECT   DetalleNomina.NumNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HE, " & vbLf
SQL = SQL & "DetalleNomina.DD, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre," & vbLf
SQL = SQL & "DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
SQL = SQL & "DetalleNomina.MontoINSS, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.INATEC, DetalleNomina.Mes13," & vbLf
SQL = SQL & "DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.TotalSubsidio, DetalleNomina.VacacionesPagadas," & vbLf
SQL = SQL & "DetalleNomina.DiasVacaciones, DetalleNomina.AdelantosVacaciones, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia," & vbLf
SQL = SQL & "DetalleNomina.IncetivoProduccion , DetalleNomina.TarifaHoraria, Nomina.Activa" & vbLf
SQL = SQL & "FROM         DetalleNomina INNER JOIN" & vbLf
SQL = SQL & "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina" & vbLf
SQL = SQL & "Where (DetalleNomina.CodEmpleado = " & CodEmpleado & ") And (Nomina.Activa = 1)"
Me.DtaConsulta.RecordSource = SQL
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
SQL = "SELECT DetalleNomSubsidio.NumNominaSubsidio, DetalleNomSubsidio.CodEmpleado, DetalleNomSubsidio.Subsidio, NomSubsidio.Activa" & vbLf
SQL = SQL & "FROM DetalleNomSubsidio INNER JOIN NomSubsidio ON DetalleNomSubsidio.NumNominaSubsidio = NomSubsidio.NumNomina" & vbLf
SQL = SQL & "Where (NomSubsidio.Activa = 1) And (DetalleNomSubsidio.CodEmpleado = " & CodEmpleado & ")"
Me.DtaConsulta.RecordSource = SQL
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
SQL = "SELECT     DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, " & vbLf
SQL = SQL & "                      DetalleNomVaca.DiasDescuento , DetalleNomVaca.AdelantoVacaciones, DetalleNomVaca.Inss, DetalleNomVaca.TarifaHoraria, NomVaca.Activa" & vbLf
SQL = SQL & "FROM         DetalleNomVaca INNER JOIN" & vbLf
SQL = SQL & "                      NomVaca ON DetalleNomVaca.NumNomVaca = NomVaca.NumNomVaca" & vbLf
SQL = SQL & "Where (DetalleNomVaca.CodEmpleado = " & CodEmpleado & ") And (NomVaca.Activa = 1)"
Me.DtaConsulta.RecordSource = SQL
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
SQL = "SELECT     DetalleNom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.SalarioAPagar, " & vbLf
SQL = SQL & "DetalleNom13Mes.DiasAPagar , DetalleNom13Mes.Adelanto13vo, Nom13Mes.Activa" & vbLf
SQL = SQL & "FROM         DetalleNom13Mes INNER JOIN" & vbLf
SQL = SQL & "Nom13Mes ON DetalleNom13Mes.NumNom13Mes = Nom13Mes.NumNom13Mes" & vbLf
SQL = SQL & "Where (Nom13Mes.Activa = 1) And (DetalleNom13Mes.CodEmpleado = " & CodEmpleado & ")"
Me.DtaConsulta.RecordSource = SQL
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
  NumNomina = Me.DtaConsulta.Recordset("NumNom13Mes")
Else
  NumNomina = -1
End If

rs.Open "DELETE FROM DetalleNom13Mes Where (NumNom13Mes = " & NumNomina & ") And (CodEmpleado = " & CodEmpleado & ")", Conexion



    

   
    Me.CmdEfectuar.Enabled = False

'CmdCancelar.Caption = "Cerrar"
Exit Sub
TipoErr:

End Sub

Private Sub ChkVaca_Click()
If ChkVaca.Value = 1 Then
Me.TxtDescuentoDias.Visible = True
'Me.TxtDiasDescuento.Visible = True
Else
    

'Me.TxtDiasDescuento.Visible = False
'Me.TxtDescuentoDias.Visible = False
End If

End Sub

Private Sub Command1_Click()
'///////Imprimo la Liquidacion//////////////////////////////////////
k% = MsgBox("Desea Imprimir las bajas?", vbYesNo)
If k% = 6 Then

'  Me.AdoEmpleados.RecordSource = "SELECT CodEmpleado1, CodEmpleado, Nombre1 + N' ' + Nombre2 + N' ' + Apellido1 + N' ' + Apellido2 AS Nombres, Sexo, NumCedula, Sindicalista, Activo From Empleado Where (Activo = 1) ORDER BY CodEmpleado1"
'  Me.AdoEmpleados.Refresh
  

'    CodEmpleado = TxtCodEmpleado.Text
'    Me.DtaConsulta.RecordSource = "SELECT [Bajas].[MontoNomPropor]+[Bajas].[MontoVaca]+[Bajas].[Monto13Mes]+[Bajas].[MontoAnosTrab]+[Bajas].[MontoCargoConfianza]+[Bajas].[MontoAntiguedad]+[Bajas].[MontoOtro] AS Ingresos, [Bajas].[Prestamo]+[Bajas].[Deducciones]+[Bajas].[MontoINSS]+[Bajas].[MontoIR] AS Egresos, Bajas.CodEmpleado From Bajas Where (((Bajas.CodEmpleado) = '" & CodEmpleado & "'))"
'    Me.DtaConsulta.Refresh
    
    ArepBajas.DataControl1.ConnectionString = ConexionReporte
    ArepBajas.LblTitulo.Caption = Titulo
    ArepBajas.LblSubtitulo.Caption = Subtitulo
    ArepBajas.ImgLogo.Picture = LoadPicture(RutaLogo)
'    ArepBajas.LblSalarioAlto.Caption = Me.TxtSalarioPromedio.Text
    cadena = "SELECT     Empleado.CodEmpleado1,Empleado.CodEmpleado, Empleado.Nombre1 + Empleado.Nombre2 + Empleado.Apellido1 + Empleado.Apellido2 AS Nombres, Bajas.FechaBaja, " & vbLf
    cadena = cadena & "                  Bajas.AnnosTrabajados, Bajas.MesesTrabajados, Bajas.DiasTrabajados, Bajas.MontoNomPropor, Bajas.MontoVaca, Bajas.Monto13Mes," & vbLf
    cadena = cadena & "                  Bajas.MontoAnosTrab, Bajas.MontoCargoConfianza, Bajas.MontoAntiguedad, Bajas.MotivoBaja, Bajas.TipoBaja, Bajas.Otro, Bajas.MontoOtro," & vbLf
    cadena = cadena & "                  Bajas.Prestamo, Bajas.Deducciones, Bajas.MontoINSS, Bajas.MontoIR, Cargo.Cargo, Departamento.Departamento, Historico.FechaContrato," & vbLf
    cadena = cadena & "                  Bajas.SalarioMensual, Bajas.HorasExtra, Bajas.Viaticos, Bajas.MontoHorasExtra, Bajas.FechaIniAgui, Bajas.FechaFinAgui, Bajas.DiasAguinaldo," & vbLf
    cadena = cadena & "                  Bajas.DiasVacaciones , Bajas.DiasMenosVaca, Bajas.FechaIniVaca, Bajas.FechaFinVaca, SalarioPromedio, SalarioBasico, TotalIngresos, TotalEgresos, TotalPagar,DiasTrabajadosAgui, DiasTrabajadosVaca" & vbLf
    cadena = cadena & "FROM         Departamento INNER JOIN" & vbLf
    cadena = cadena & "                  Cargo INNER JOIN" & vbLf
    cadena = cadena & "                  Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento INNER JOIN" & vbLf
    cadena = cadena & "                  Bajas ON Empleado.CodEmpleado = Bajas.CodEmpleado LEFT OUTER JOIN" & vbLf
    cadena = cadena & "                  Historico ON Empleado.CodEmpleado = Historico.Codempleado" & vbLf
    cadena = cadena & "Where (Activo = 1) ORDER BY CodEmpleado1"
    
    ArepBajas.DataControl1.Source = cadena
'    ArepBajas.LblDiasNeto = Int(Format(DiasTrabajadosVaca * (1 / 12), "####00")) - Val(Me.TxtDescuentoDias.Text)
'    TotalIngresos = (TotalVacaciones + TotalAguinaldo + TotalAntiguedad + MontoHRSExtras + TotalOtrosSalarios)
'    TotalEgresos = TotalInss + MontoIR + Prestamo + Deducciones
'     NetoPagar = TotalIngresos - TotalEgresos
'     ArepBajas.LblTotalEgresos.Caption = Format(TotalEgresos, "##,##0.00")
'     ArepBajas.LblTotalIngresos.Caption = Format(TotalIngresos, "##,##0.00")
'     ArepBajas.LblNetoPagar.Caption = Format(NetoPagar, "##,##0.00")
'     ArepBajas.LblSalarioBasico.Caption = Me.TxtSalarioBasico.Text
'     ArepBajas.LblDiasAgui.Caption = DiasTrabajadosAgui
'     ArepBajas.LblDiasVaca.Caption = DiasTrabajadosVaca
    ArepBajas.Show 1
    Me.CmdEfectuar.Enabled = True
End If

End Sub

Private Sub CmdImprimirHistorial_Click()
Dim SQLSalarios As String, FechaEgreso As Date, FechaContrato As Date, FechaHistorico As Date, FechaBusqueda As Date
Dim CodEmpleado As String, NumeroEmpleado As Integer, I As Integer
    CodEmpleado = TxtCodEmpleado.Text
  
  '///////////Busco la Fecha para la Busqueda////////////////////////////
  
FechaEgreso = Me.TxtFechaHistorial.Value
'FechaContrato = Me.TxtFechaContrato.Text


NumeroEmpleado = Me.TxtCodEmpleado.Text

SQLSalarios = "SELECT DISTINCT" & vbLf
SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
SQLSalarios = SQLSalarios & "AS Destajo, SUM(dbo.DetalleNomina.Incentivos) AS Incentivos, SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo)" & vbLf
SQLSalarios = SQLSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes AS MES," & vbLf
SQLSalarios = SQLSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
SQLSalarios = SQLSalarios & "FROM   dbo.DetalleNomina INNER JOIN" & vbLf
SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
SQLSalarios = SQLSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
SQLSalarios = SQLSalarios & "Having (dbo.DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") And (Sum(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0) AND (MAX(Nomina.FechaNomina) <= CONVERT(DATETIME, '" & Format(FechaEgreso, "yyyy/mm/dd") & "', 102))"
Me.DtaConsulta.RecordSource = SQLSalarios
Me.DtaConsulta.Refresh

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
Me.DtaConsulta.Recordset.MoveLast
I = 0
Do While Not Me.DtaConsulta.Recordset.BOF
  If I = 1 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")

  ElseIf I = 5 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    Exit Do
  ElseIf I = 0 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    FechaHistorico = Me.DtaConsulta.Recordset("FechaFin")
  Else
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
  End If
  I = I + 1

  Me.DtaConsulta.Recordset.MovePrevious
Loop

  
  
  
    ArepHistorialLiquida.DataControl1.ConnectionString = ConexionReporte
    ArepHistorialLiquida.LblTitulo.Caption = Titulo
    ArepHistorialLiquida.LblSubtitulo.Caption = Subtitulo
    ArepHistorialLiquida.ImgLogo.Picture = LoadPicture(RutaLogo)
    FechaEgreso = Me.TxtUltFechaNomina.Value
'    FechaHistorico = DateSerial(Year(FechaEgreso), Month(FechaEgreso), 1 - 1)
    FechaContrato = Me.TxtFechaContrato.Text
'    FechaBusqueda = DateSerial(Year(FechaEgreso), Month(FechaEgreso) - 6, 1)

    Año = Year(FechaBusqueda)
    Mes = Month(FechaBusqueda)
    
    
    SQLSalarios = "SELECT DISTINCT" & vbLf
    SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
    SQLSalarios = SQLSalarios & "AS Destajo, SUM(dbo.DetalleNomina.SeptimoDia) AS Septimo, SUM(dbo.DetalleNomina.OtrosIngresos) AS Otros, SUM(dbo.DetalleNomina.Incentivos)" & vbLf
    SQLSalarios = SQLSalarios & "AS Incentivos," & vbLf
    SQLSalarios = SQLSalarios & "SUM (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.OtrosIngresos)" & vbLf
    SQLSalarios = SQLSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes," & vbLf
    SQLSalarios = SQLSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
    SQLSalarios = SQLSalarios & "FROM    dbo.DetalleNomina INNER JOIN" & vbLf
    SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
    SQLSalarios = SQLSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
    SQLSalarios = SQLSalarios & "HAVING(SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0) And (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND" & vbLf
    SQLSalarios = SQLSalarios & "'" & Format(FechaHistorico, "yyyymmdd") & "')" & vbLf
    SQLSalarios = SQLSalarios & "ORDER BY dbo.Nomina.Ano, dbo.Nomina.Mes"
    
    ArepHistorialLiquida.LblCodEmpleado.Caption = Me.TxtCodEmpleado1.Text
    ArepHistorialLiquida.LblNombreEmpleado.Caption = Me.TxtNombre1.Text + " " + Me.TxtNombre2.Text + " " + Me.TxtApellido1.Text + " " + Me.TxtApellido2.Text
    ArepHistorialLiquida.LblDepartamento.Caption = Me.TxtDepartamento.Text
    ArepHistorialLiquida.LblCargo.Caption = Me.TxtCargo.Text
    ArepHistorialLiquida.LblAños.Caption = Me.TxtAnnos.Text
    ArepHistorialLiquida.LblDias.Caption = Me.TxtDiasTrabajados.Text
    ArepHistorialLiquida.LblMeses.Caption = Me.TxtMeses.Text
    
    ArepHistorialLiquida.LblSalarioAlto.Caption = Me.TxtSalarioAlto.Text
    ArepHistorialLiquida.LblSalarioBasico.Caption = Me.TxtSalarioBasico.Text
    ArepHistorialLiquida.LblSalarioPromedio.Caption = Me.TxtSalarioPromedio.Text
    ArepHistorialLiquida.LblTarifaHoraria.Caption = Me.TxtTarifa.Text
    ArepHistorialLiquida.lblVacaciones.Caption = Me.DTPFechaIniVaca.Value
    ArepHistorialLiquida.LblAguinaldo.Caption = Me.DTPFechaIniAgui.Value

    ArepHistorialLiquida.DataControl1.ConnectionString = ConexionReporte
    ArepHistorialLiquida.DataControl1.Source = SQLSalarios
    ArepHistorialLiquida.Show 1
    
End Sub


Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Activate()
'txtCodEmpleado.Text = CodEmpleado
End Sub

Private Sub Form_Load()
 Dim ConexionSTR1 As String
'Me.TxtUltFechaNomina.Value = Format(Now, "dd/mm/yyyy")
'Me.TxtFechaHistorial.Value = Format(Now, "dd/mm/yyyy")
Me.CmdDetalle.BackColor = RGB(219, 226, 242)
Me.CmdImprimirHistorial.BackColor = RGB(219, 226, 242)
Me.Skin1.ApplySkin hWnd
 Me.TDBGridSalarios.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.TDBGridSalarios.OddRowStyle.BackColor = &H80000005
 Me.TDBGridSalarios.AlternatingRowStyle = True
 
 Me.DTPFechaIniAgui.Value = "01/12/2006"
 Me.DTPFechaIniVaca.Value = "01/06/2007"
 
 Open App.Path + "\SysInfo.dll" For Input As #1
  Do Until EOF(1)
   Line Input #1, NextLine
        ConexionSTR1 = Trim(NextLine)
   Loop
 Close #1
  
  
 Me.TxtFechaHistorial.Value = "20/08/2007"
 Me.TxtUltFechaNomina.Value = "05/09/2007"
  
  Conexion = ConexionSTR1
  ConexionReporte = ConexionSTR1
  
With Me.AdoDatosEmpresa
  .ConnectionString = Conexion
End With
 
With Me.AdoEmpleados
  .ConnectionString = Conexion
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

With Me.DtaHistorico
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaNominas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaPrestamo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With


With Me.DtaTipoNomina
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "TipoNomina"
   .Refresh
End With

Me.AdoEmpleados.RecordSource = "SELECT CodEmpleado1, CodEmpleado, Nombre1 + N' ' + Nombre2 + N' ' + Apellido1 + N' ' + Apellido2 AS Nombres, Sexo, NumCedula, Sindicalista, Activo From Empleado Where (Activo = 1)  ORDER BY CodEmpleado1"
Me.AdoEmpleados.Refresh

End Sub

Private Sub TDBGrid1_DblClick()
  Dim DiaMes As Double, CodEmpleado As Double
  
  CodEmpleado = Me.TDBGrid1.Columns(1).Text
  Me.TxtCodEmpleado1.Text = Me.TDBGrid1.Columns(0).Text
  Me.TxtCodEmpleado.Text = Me.TDBGrid1.Columns(1).Text


'    SQlEmpleado = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2,Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion,Empleado.Sexo , Empleado.Activo, Empleado.Liquidado, Empleado.Ausente, Empleado.Gravidez,Empleado.TarifaHoraria FROM Departamento INNER JOIN Cargo INNER JOIN" & vbLf
'    SQlEmpleado = SQlEmpleado & "Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento" & vbLf
'    SQlEmpleado = SQlEmpleado & "WHERE  (Empleado.CodEmpleado = " & CodEmpleado & ") AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
'    DtaEmpleado.RecordSource = SQlEmpleado
'    DtaEmpleado.Refresh
'
'
'    If Not DtaEmpleado.Recordset.EOF Then
'
'    TxtNombre1 = DtaEmpleado.Recordset("Nombre1")
'    TxtNombre2 = DtaEmpleado.Recordset("Nombre2")
'    TxtApellido1 = DtaEmpleado.Recordset("Apellido1")
'    TxtApellido2 = DtaEmpleado.Recordset("Apellido2")
'    Me.CmdAcercade.Caption = Me.TxtCodEmpleado1.Text + "-" + Me.TxtNombre1.Text + " " + Me.TxtNombre2.Text + " " + Me.TxtApellido1.Text + " " + Me.TxtApellido2.Text
'    TxtDireccion = DtaEmpleado.Recordset("Direccion")
'    TxtCargo = DtaEmpleado.Recordset("Cargo")
'    TxtDepartamento = DtaEmpleado.Recordset("departamento")
'    TxtSexo = DtaEmpleado.Recordset("sexo")
'
'    Me.DtaControles.Refresh
'    If Not Me.DtaControles.Recordset.EOF Then
'     DiaMes = Me.DtaControles.Recordset("DiasMes")
'    End If
'
'    Me.TxtCodEmpleado.Text = DtaEmpleado.Recordset("CodEmpleado")
'    Me.TxtTarifa.Text = DtaEmpleado.Recordset("TarifaHoraria")
'    Me.TxtSalarioBasico.Text = Format(DtaEmpleado.Recordset("TarifaHoraria") * DiaMes * 8, "##,##0.00")
'    End If
End Sub

Private Sub TxtCodEmpleado_Change()
On Error GoTo TipoErr
Dim FechaContrato As Date, FechaInicio As Date, FechaEgreso As Date
Dim FechaHoy As Date, SueldoPeriodo As Double, FechaHistorico As Date
Dim FechaUltNomina As Date, I As Integer, NumeroEmpleado As Integer
Dim annos As Date
Dim SQlEmpleado As String
Dim SQlPrestamo As String
Dim SQlDeducciones As String
Dim SqlNominas As String
Dim FechaBusqueda As Date, Año As Integer, Mes As Integer
Dim Contador As Integer, TotalSalario As Double, Salario As Double, SalarioAlto As Double, SalarioPromedio As Double

SQlEmpleado = "SELECT Empleado.SalarioFijo,Empleado.SueldoPeriodo,Empleado.CodEmpleado,Empleado.CodEmpleado1,Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion, Empleado.Direccion, Empleado.Sexo, Empleado.Activo,Empleado.TarifaHoraria FROM Departamento INNER JOIN (Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) ON Departamento.CodDepartamento = Empleado.CodDepartamento WHERE Empleado.CodEmpleado='" & TxtCodEmpleado.Text & "'"
DtaEmpleado.RecordSource = SQlEmpleado
DtaEmpleado.Refresh

SQlPrestamo = "SELECT Prestamo.NumPrestamo, Prestamo.CodEmpleado, Prestamo.Monto, Prestamo.CantCuotas, Prestamo.Interes, Prestamo.Saldo, Prestamo.FechaInicial, Prestamo.Cancelado From Prestamo WHERE Prestamo.Cancelado=0 AND Prestamo.CodEmpleado='" & TxtCodEmpleado.Text & "'"
DtaPrestamo.RecordSource = SQlPrestamo
DtaPrestamo.Refresh

SQlDeducciones = "SELECT Deduccion.NumDeduccion, Deduccion.CodEmpleado, Deduccion.CodTipoDeduccion, DetalleDeduccion.NumDeduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.pagado  FROM Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion WHERE DetalleDeduccion.pagado=0 AND Deduccion.CodEmpleado='" & TxtCodEmpleado.Text & "'"
DtaDeducciones.RecordSource = SQlDeducciones
DtaDeducciones.Refresh

'If Me.Combo1.Text = "Administracion" Then
'  CodTiposNomina = "01"
' Me.TxtFechaHistorial.Value = "31/08/2007"
' Me.TxtUltFechaNomina.Value = "05/09/2007"
'Else
'  CodTiposNomina = "02"
' Me.TxtFechaHistorial.Value = "20/08/2007"
' Me.TxtUltFechaNomina.Value = "05/09/2007"
'End If



If Not DtaEmpleado.Recordset.EOF Then
       If DtaEmpleado.Recordset("activo") = False Then
           MsgBox "Este empleado ya fue dado de Baja"
           Exit Sub
        End If
    Me.TxtCodEmpleado1.Text = DtaEmpleado.Recordset("CodEmpleado1")
    CodEmpleado = DtaEmpleado.Recordset("CodEmpleado")
    TxtNombre1 = DtaEmpleado.Recordset("Nombre1")
    TxtNombre2 = DtaEmpleado.Recordset("Nombre2")
    TxtApellido1 = DtaEmpleado.Recordset("Apellido1")
    TxtApellido2 = DtaEmpleado.Recordset("Apellido2")
    TxtDireccion = DtaEmpleado.Recordset("Direccion")
    TxtCargo = DtaEmpleado.Recordset("Cargo")
    TxtDepartamento = DtaEmpleado.Recordset("departamento")
    TxtSexo = DtaEmpleado.Recordset("sexo")
    SalarioBasico = DtaEmpleado.Recordset("SueldoPeriodo")
    If DtaEmpleado.Recordset("SalarioFijo") = "S" Then
     SueldoFijo = True
    Else
     SueldoFijo = False
    End If
Else
    TxtNombre1 = ""
    TxtNombre2 = ""
    TxtApellido1 = ""
    TxtApellido2 = ""
    TxtDireccion = ""
    TxtCargo = ""
    TxtDepartamento = ""
    TxtSexo = ""
    TxtAnnos.Text = ""
    TxtMeses.Text = ""
    TxtFechaContrato.Text = ""
    TxtUltFechaNomina.Value = Now
    TxtMotivo.Text = ""
   Exit Sub
End If


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
'   TxtUltFechaNomina.Value = FechaUltNomina
   NumFecha1 = FechaUltNomina
Else
'   TxtUltFechaNomina = TxtFechaContrato
   'MsgBox "No ha sido Grabada Ninguna nómina a este empleado, Se le realizará la baja desde su contrato"
End If







'///////////Busco la Fecha para la Busqueda////////////////////////////

NumeroEmpleado = Me.TxtCodEmpleado.Text

FechaEgreso = Me.TxtFechaHistorial.Value
FechaContrato = Me.TxtFechaContrato.Text

SQLSalarios = "SELECT DISTINCT" & vbLf
SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
SQLSalarios = SQLSalarios & "AS Destajo, SUM(dbo.DetalleNomina.Incentivos) AS Incentivos, SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo)" & vbLf
SQLSalarios = SQLSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes AS MES," & vbLf
SQLSalarios = SQLSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
SQLSalarios = SQLSalarios & "FROM   dbo.DetalleNomina INNER JOIN" & vbLf
SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
SQLSalarios = SQLSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
SQLSalarios = SQLSalarios & "Having (dbo.DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") And (Sum(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0) AND (MAX(Nomina.FechaNomina) <= CONVERT(DATETIME, '" & Format(FechaEgreso, "yyyy/mm/dd") & "', 102))"


'SQLSalarios = "SELECT DISTINCT" & vbLf
'SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
'SQLSalarios = SQLSalarios & "AS Destajo, SUM(dbo.DetalleNomina.SeptimoDia) AS Septimo, SUM(dbo.DetalleNomina.OtrosIngresos) AS Otros, SUM(dbo.DetalleNomina.Incentivos)" & vbLf
'SQLSalarios = SQLSalarios & "AS Incentivos," & vbLf
'SQLSalarios = SQLSalarios & "Sum (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.OtrosIngresos)" & vbLf
'SQLSalarios = SQLSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes," & vbLf
'SQLSalarios = SQLSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
'SQLSalarios = SQLSalarios & "FROM         dbo.DetalleNomina INNER JOIN" & vbLf
'SQLSalarios = SQLSalarios & "                      dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
'SQLSalarios = SQLSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
'SQLSalarios = SQLSalarios & "Having (Sum(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0) And (dbo.DetalleNomina.CodEmpleado = " & NumeroEmpleado & ")" & vbLf
'SQLSalarios = SQLSalarios & "ORDER BY dbo.Nomina.Ano, dbo.Nomina.Mes"


Me.DtaConsulta.RecordSource = SQLSalarios
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
 Me.DtaConsulta.Recordset.MoveLast
Else
 FechaHistorico = Format(Now, "dd/mm/yyyy")
 FechaBusqueda = Format(Now, "dd/mm/yyyy")
End If
I = 0
Do While Not Me.DtaConsulta.Recordset.BOF
  If I = 1 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")

  ElseIf I = 5 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    Exit Do
  ElseIf I = 0 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    FechaHistorico = Me.DtaConsulta.Recordset("FechaFin")
  Else
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
  End If
  I = I + 1

  Me.DtaConsulta.Recordset.MovePrevious
Loop

FechaEgreso = Me.TxtUltFechaNomina.Value

FechaContrato = Me.TxtFechaContrato.Text

'//////////SUMO 1 PARA AJUSTAR QUE SIEMPRE DA 1 DIA MENOS//////
annos = CDbl(FechaEgreso) - CDbl(FechaContrato) + 1
TxtAnnos.Text = Format(annos / 365, "###,##0.00")
TxtMeses.Text = Format(annos / 30.41, "###,##0.00")
Dias = Format(annos * 365, "###,##0.00")
    Me.CmdEfectuar.Enabled = False
Me.TxtDiasTrabajados.Text = Format(annos, "###,##0")


Año = Year(FechaBusqueda)
Mes = Month(FechaBusqueda)


    SQLSalarios = "SELECT DISTINCT" & vbLf
    SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
    SQLSalarios = SQLSalarios & "AS Destajo, SUM(dbo.DetalleNomina.SeptimoDia) AS Septimo, SUM(dbo.DetalleNomina.OtrosIngresos) AS Otros, SUM(dbo.DetalleNomina.Incentivos)" & vbLf
    SQLSalarios = SQLSalarios & "AS Incentivos," & vbLf
    SQLSalarios = SQLSalarios & "SUM (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.OtrosIngresos)" & vbLf
    SQLSalarios = SQLSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes," & vbLf
    SQLSalarios = SQLSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
    SQLSalarios = SQLSalarios & "FROM    dbo.DetalleNomina INNER JOIN" & vbLf
    SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
    SQLSalarios = SQLSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
    SQLSalarios = SQLSalarios & "HAVING(SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0) And (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND" & vbLf
    SQLSalarios = SQLSalarios & "'" & Format(FechaHistorico, "yyyymmdd") & "')" & vbLf
    SQLSalarios = SQLSalarios & "ORDER BY dbo.Nomina.Ano, dbo.Nomina.Mes"



Me.AdoSalarios.RecordSource = SQLSalarios
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
         Me.TxtAntiguedad.Text = Me.AdoAntiguedad.Recordset("porcent")
        Else
         Me.TxtAntiguedad.Text = 0
        End If
        SalarioPromedio = SueldoPeriodo * PAntiguedad
        SalarioAlto = SueldoPeriodo * PAntiguedad
         
       Else
        SalarioPromedio = SueldoPeriodo
        SalarioAlto = SueldoPeriodo
         Me.TxtAntiguedad.Text = 0
       End If
 
  
Else
    Me.TxtAntiguedad.Text = 0
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
    SalarioPromedio = TotalSalario / Contador
   End If

 End If
    Me.TxtSalarioPromedio.Text = Format(SalarioPromedio, "##,##0.00")
    Me.TxtSalarioAlto.Text = Format(SalarioAlto, "##,##0.00")
    
    Dim AñoActual As Integer, CodTipoNomina As String
    
    CodigoEmpleado = Me.TxtCodEmpleado.Text


'/////////CONSULTA EL SALARIO Y TIPO DE NOMINA DEL EMPLEADO//////////////////////////

 SQL = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento," & vbLf
 SQL = SQL & "Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.OtrosIngresos, Empleado.DescripOtrIngre," & vbLf
 SQL = SQL & "Empleado.ExentoIr, Empleado.PagoInssPatronal, Empleado.Activo, Empleado.Liquidado, Empleado.Ausente, Empleado.SalarioFijo," & vbLf
 SQL = SQL & "Empleado.SumarSubsidio , Empleado.PorcientoIncentivo, Empleado.Gravidez, TipoNomina.Periodo" & vbLf
 SQL = SQL & "FROM Empleado INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina" & vbLf
 SQL = SQL & "WHERE     (Empleado.CodEmpleado = '" & CodigoEmpleado & "') AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
 Me.DtaConsulta.RecordSource = SQL
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
'          Me.DTPFechaIniVaca.Value = DateSerial(Year(FechaEgreso), 6, 1)
        Else
'          Me.DTPFechaIniVaca.Value = DateSerial(Year(FechaEgreso) - 1, 12, 1)
        End If
       Else
       FechaInicioAgui = FechaContrato
'       Me.DTPFechaIniAgui.Value = FechaInicioAgui
'       Me.DTPFechaIniVaca.Value = FechaContrato
       End If
Exit Sub
TipoErr:
'ControlErrores
CmdEfectuar.Enabled = False

End Sub

Private Sub TxtCodEmpleado1_Change()
  Dim DiaMes As Double

    SQlEmpleado = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2,Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion,Empleado.Sexo , Empleado.Activo, Empleado.Liquidado, Empleado.Ausente, Empleado.Gravidez,Empleado.TarifaHoraria FROM Departamento INNER JOIN Cargo INNER JOIN" & vbLf
    SQlEmpleado = SQlEmpleado & "Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento" & vbLf
    SQlEmpleado = SQlEmpleado & "WHERE  (Empleado.CodEmpleado1 = '" & Me.TxtCodEmpleado1.Text & "') AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
    DtaEmpleado.RecordSource = SQlEmpleado
    DtaEmpleado.Refresh
    
     
    If Not DtaEmpleado.Recordset.EOF Then
    
    TxtNombre1 = DtaEmpleado.Recordset("Nombre1")
    TxtNombre2 = DtaEmpleado.Recordset("Nombre2")
    TxtApellido1 = DtaEmpleado.Recordset("Apellido1")
    TxtApellido2 = DtaEmpleado.Recordset("Apellido2")
    Me.CmdAcercade.Caption = Me.TxtCodEmpleado1.Text + "-" + Me.TxtNombre1.Text + " " + Me.TxtNombre2.Text + " " + Me.TxtApellido1.Text + " " + Me.TxtApellido2.Text
    TxtDireccion = DtaEmpleado.Recordset("Direccion")
    TxtCargo = DtaEmpleado.Recordset("Cargo")
    TxtDepartamento = DtaEmpleado.Recordset("departamento")
    TxtSexo = DtaEmpleado.Recordset("sexo")
    
    Me.DtaControles.Refresh
    If Not Me.DtaControles.Recordset.EOF Then
     DiaMes = Me.DtaControles.Recordset("DiasMes")
    End If
    
    Me.TxtCodEmpleado.Text = DtaEmpleado.Recordset("CodEmpleado")
    Me.TxtTarifa.Text = DtaEmpleado.Recordset("TarifaHoraria")
    Me.TxtSalarioBasico.Text = Format(DtaEmpleado.Recordset("TarifaHoraria") * DiaMes * 8, "##,##0.00")
    End If

End Sub

Private Sub TxtCodEmpleado1_KeyPress(KeyAscii As Integer)
Dim SQLSalarios As String, DiaMes As Double
Dim SQlEmpleado As String
If KeyAscii = 13 Then

    SQlEmpleado = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2,Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion,Empleado.Sexo , Empleado.Activo, Empleado.Liquidado, Empleado.Ausente, Empleado.Gravidez,Empleado.TarifaHoraria FROM Departamento INNER JOIN Cargo INNER JOIN" & vbLf
    SQlEmpleado = SQlEmpleado & "Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento" & vbLf
    SQlEmpleado = SQlEmpleado & "WHERE  (Empleado.CodEmpleado1 = '" & Me.TxtCodEmpleado1.Text & "') AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
    DtaEmpleado.RecordSource = SQlEmpleado
    DtaEmpleado.Refresh
    
     
    If Not DtaEmpleado.Recordset.EOF Then
    
    TxtNombre1 = DtaEmpleado.Recordset("Nombre1")
    TxtNombre2 = DtaEmpleado.Recordset("Nombre2")
    TxtApellido1 = DtaEmpleado.Recordset("Apellido1")
    TxtApellido2 = DtaEmpleado.Recordset("Apellido2")
    Me.CmdAcercade.Caption = Me.TxtCodEmpleado1.Text + "-" + Me.TxtNombre1.Text + " " + Me.TxtNombre2.Text + " " + Me.TxtApellido1.Text + " " + Me.TxtApellido2.Text
    TxtDireccion = DtaEmpleado.Recordset("Direccion")
    TxtCargo = DtaEmpleado.Recordset("Cargo")
    TxtDepartamento = DtaEmpleado.Recordset("departamento")
    TxtSexo = DtaEmpleado.Recordset("sexo")

    Me.DtaControles.Refresh
    If Not Me.DtaControles.Recordset.EOF Then
     DiaMes = Me.DtaControles.Recordset("DiasMes")
    End If
    
    Me.TxtCodEmpleado.Text = DtaEmpleado.Recordset("CodEmpleado")
    Me.TxtTarifa.Text = DtaEmpleado.Recordset("TarifaHoraria")
    Me.TxtSalarioBasico.Text = Format(DtaEmpleado.Recordset("TarifaHoraria") * DiaMes * 8, "##,##0.00")
    End If

End If
End Sub

Private Sub TxtDescuentoDias_Change()
If Not IsNumeric(TxtDescuentoDias.Text) And Not TxtDescuentoDias.Text = "" Then
 MsgBox "El numero Digitado no es Numerico", vbCritical, "Sistema de Nominas"
 Me.TxtDescuentoDias.Text = ""
End If
End Sub

Private Sub TxtFechaHistorial_Change()
Dim FechaEgreso As Date, FechaContrato As Date, Año As Integer, Mes As Integer
Dim FechaBusqueda As Date, TotalSalario As Double, SalarioPromedio As Double, Contador As Integer, I As Integer
Dim SQLSalarios As String, SalarioAlto As Double, Salario As Double, FechaHistorico As Date, NumeroEmpleado As Integer
FechaEgreso = Me.TxtFechaHistorial.Value
FechaContrato = Me.TxtFechaContrato.Text
'//////////SUMO 1 PARA AJUSTAR QUE SIEMPRE DA 1 DIA MENOS//////
'annos = CDbl(FechaEgreso) - CDbl(FechaContrato) + 1
'TxtAnnos.Text = Format(annos / 365, "###,##0.00")
'TxtMeses.Text = Format(annos / 30.41, "###,##0.00")
'Me.txtDiasTrabajados.Text = Format(annos, "###,##0")
'Dias = annos
Me.CmdEfectuar.Enabled = False

'///////////Busco la Fecha para la Busqueda////////////////////////////

NumeroEmpleado = Me.TDBGrid1.Columns(1).Text

SQLSalarios = "SELECT DISTINCT" & vbLf
SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
SQLSalarios = SQLSalarios & "AS Destajo, SUM(dbo.DetalleNomina.Incentivos) AS Incentivos, SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo)" & vbLf
SQLSalarios = SQLSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes AS MES," & vbLf
SQLSalarios = SQLSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
SQLSalarios = SQLSalarios & "FROM   dbo.DetalleNomina INNER JOIN" & vbLf
SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
SQLSalarios = SQLSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
SQLSalarios = SQLSalarios & "Having (dbo.DetalleNomina.CodEmpleado = " & NumeroEmpleado & ") And (Sum(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0) AND (MAX(Nomina.FechaNomina) <= CONVERT(DATETIME, '" & Format(FechaEgreso, "yyyy/mm/dd") & "', 102))"
Me.DtaConsulta.RecordSource = SQLSalarios
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
  Me.DtaConsulta.Recordset.MoveLast
Else
 FechaHistorico = Format(Now, "dd/mm/yyyy")
 FechaBusqueda = Format(Now, "dd/mm/yyyy")
End If
I = 0


Do While Not Me.DtaConsulta.Recordset.BOF
  If I = 1 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")

  ElseIf I = 5 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    Exit Do
  ElseIf I = 0 Then
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
    FechaHistorico = Me.DtaConsulta.Recordset("FechaFin")
  Else
    FechaBusqueda = Me.DtaConsulta.Recordset("FechaInicio")
  End If
  I = I + 1

  Me.DtaConsulta.Recordset.MovePrevious
Loop


FechaEgreso = Me.TxtFechaHistorial.Value
'FechaHistorico = DateSerial(Year(FechaEgreso), Month(FechaEgreso), 1 - 1)
FechaContrato = Me.TxtFechaContrato.Text
'FechaBusqueda = DateSerial(Year(FechaEgreso), Month(FechaEgreso) - 6, 1)
Año = Year(FechaBusqueda)
Mes = Month(FechaBusqueda)

    SQLSalarios = "SELECT DISTINCT" & vbLf
    SQLSalarios = SQLSalarios & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
    SQLSalarios = SQLSalarios & "AS Destajo, SUM(dbo.DetalleNomina.SeptimoDia) AS Septimo, SUM(dbo.DetalleNomina.OtrosIngresos) AS Otros, SUM(dbo.DetalleNomina.Incentivos)" & vbLf
    SQLSalarios = SQLSalarios & "AS Incentivos," & vbLf
    SQLSalarios = SQLSalarios & "SUM (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.OtrosIngresos)" & vbLf
    SQLSalarios = SQLSalarios & "AS TotalIngresos, MIN(dbo.Nomina.FechaNominaINI) AS FechaInicio, MAX(dbo.Nomina.FechaNomina) AS FechaFin, dbo.Nomina.Mes," & vbLf
    SQLSalarios = SQLSalarios & "dbo.Nomina.Ano AS AÑO" & vbLf
    SQLSalarios = SQLSalarios & "FROM    dbo.DetalleNomina INNER JOIN" & vbLf
    SQLSalarios = SQLSalarios & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
    SQLSalarios = SQLSalarios & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
    SQLSalarios = SQLSalarios & "HAVING(SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo) <> 0) And (DetalleNomina.CodEmpleado = '" & Me.TxtCodEmpleado.Text & "') AND (MIN(Nomina.FechaNomina) BETWEEN '" & Format(FechaBusqueda, "yyyymmdd") & "' AND" & vbLf
    SQLSalarios = SQLSalarios & "'" & Format(FechaHistorico, "yyyymmdd") & "')" & vbLf
    SQLSalarios = SQLSalarios & "ORDER BY dbo.Nomina.Ano, dbo.Nomina.Mes"

Me.AdoSalarios.RecordSource = SQLSalarios
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
         Me.TxtAntiguedad.Text = Me.AdoAntiguedad.Recordset("porcent")
        Else
         Me.TxtAntiguedad.Text = 0
        End If
        SalarioPromedio = SueldoPeriodo * PAntiguedad
        SalarioAlto = SueldoPeriodo * PAntiguedad
         
       Else
        SalarioPromedio = SueldoPeriodo
        SalarioAlto = SueldoPeriodo
         Me.TxtAntiguedad.Text = 0
       End If
 
  
Else
    Me.TxtAntiguedad.Text = 0
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

   If Contador = 0 Then
    SalarioPromedio = 0
   Else
    SalarioPromedio = TotalSalario / Contador
   End If
 End If
    Me.TxtSalarioPromedio.Text = Format(SalarioPromedio, "##,##0.00")
    Me.TxtSalarioAlto.Text = Format(SalarioAlto, "##,##0.00")


    Dim AñoActual As Integer, CodTipoNomina As String
    
    CodigoEmpleado = Me.TxtCodEmpleado.Text


'/////////CONSULTA EL SALARIO Y TIPO DE NOMINA DEL EMPLEADO//////////////////////////

 SQL = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento," & vbLf
 SQL = SQL & "Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.OtrosIngresos, Empleado.DescripOtrIngre," & vbLf
 SQL = SQL & "Empleado.ExentoIr, Empleado.PagoInssPatronal, Empleado.Activo, Empleado.Liquidado, Empleado.Ausente, Empleado.SalarioFijo," & vbLf
 SQL = SQL & "Empleado.SumarSubsidio , Empleado.PorcientoIncentivo, Empleado.Gravidez, TipoNomina.Periodo" & vbLf
 SQL = SQL & "FROM Empleado INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina" & vbLf
 SQL = SQL & "WHERE     (Empleado.CodEmpleado = '" & CodigoEmpleado & "') AND (Empleado.Activo = 1) AND (Empleado.Liquidado = 0)"
 Me.DtaConsulta.RecordSource = SQL
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
Dim FechaBusqueda As Date, TotalSalario As Double, SalarioPromedio As Double, Contador As Integer, I As Integer
Dim SQLSalarios As String, SalarioAlto As Double, Salario As Double, FechaHistorico As Date, NumeroEmpleado As Integer
FechaEgreso = Me.TxtUltFechaNomina.Value
FechaContrato = Me.TxtFechaContrato.Text
'//////////SUMO 1 PARA AJUSTAR QUE SIEMPRE DA 1 DIA MENOS//////
annos = CDbl(FechaEgreso) - CDbl(FechaContrato) + 1
TxtAnnos.Text = Format(annos / 365, "###,##0.00")
TxtMeses.Text = Format(annos / 30.41, "###,##0.00")
Me.TxtDiasTrabajados.Text = Format(annos, "###,##0")
Dias = annos
Me.CmdEfectuar.Enabled = False

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
