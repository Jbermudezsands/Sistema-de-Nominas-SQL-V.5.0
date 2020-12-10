VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form Frm13VacaMes 
   Caption         =   "Calculo del 13vo mes y Vacaciones"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   ScaleHeight     =   568
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   872
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   375
      Left            =   10920
      Top             =   8880
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   13095
      TabIndex        =   53
      Top             =   0
      Width           =   13095
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "NOMINAS MES"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   5040
         TabIndex        =   54
         Top             =   360
         Width           =   4200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   13080
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image2 
         Height          =   1020
         Left            =   0
         Picture         =   "Frm13VacaMes.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1290
      End
   End
   Begin MSAdodcLib.Adodc AdoHistorialSalarial 
      Height          =   375
      Left            =   7680
      Top             =   9000
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
      Caption         =   "AdoHistorialSalarial"
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
   Begin MSAdodcLib.Adodc Dta13voMes 
      Height          =   375
      Left            =   7560
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
      Caption         =   "Dta13voMes"
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
   Begin MSAdodcLib.Adodc DtaDetalleNom13Mes 
      Height          =   375
      Left            =   7560
      Top             =   9360
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
      Caption         =   "DtaDetalleNom13Mes"
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
   Begin MSAdodcLib.Adodc DtaNom13Mes 
      Height          =   375
      Left            =   7560
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
      Caption         =   "DtaNom13Mes"
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
      Left            =   3480
      Top             =   9960
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
   Begin MSAdodcLib.Adodc DtaVacaciones 
      Height          =   375
      Left            =   4680
      Top             =   9720
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
      Caption         =   "DtaVacaciones"
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
   Begin MSAdodcLib.Adodc DtaDetalleNomVaca 
      Height          =   375
      Left            =   7200
      Top             =   10080
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
      Caption         =   "DtaDetalleNomVaca"
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
   Begin MSAdodcLib.Adodc DtaNomVaca 
      Height          =   375
      Left            =   9000
      Top             =   9720
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
      Caption         =   "DtaNomVaca"
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
      Left            =   120
      Top             =   9840
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
   Begin MSAdodcLib.Adodc DtaEmpleados 
      Height          =   375
      Left            =   4080
      Top             =   9360
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
      Caption         =   "DtaEmpleados"
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
   Begin MSAdodcLib.Adodc DtaDetalleNominas 
      Height          =   375
      Left            =   960
      Top             =   8880
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
      Caption         =   "DtaDetalleNominas"
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
   Begin MSAdodcLib.Adodc DtaConsecutivos 
      Height          =   375
      Left            =   480
      Top             =   8880
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
      Caption         =   "DtaConsecutivos"
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
      Left            =   120
      Top             =   9240
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
      Left            =   480
      Top             =   8880
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
   Begin MSAdodcLib.Adodc DtaAdelanto 
      Height          =   375
      Left            =   480
      Top             =   9240
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
   Begin MSAdodcLib.Adodc DtaNominas 
      Height          =   375
      Left            =   960
      Top             =   8880
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
   Begin VB.PictureBox Picture1 
      Height          =   7095
      Left            =   120
      ScaleHeight     =   7035
      ScaleWidth      =   12795
      TabIndex        =   0
      Top             =   1200
      Width           =   12855
      Begin VB.TextBox txtAntiguedad 
         Height          =   285
         Left            =   9720
         TabIndex        =   43
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   600
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo DBCNominas 
         Bindings        =   "Frm13VacaMes.frx":0AFE
         Height          =   315
         Left            =   1920
         TabIndex        =   29
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nomina"
         Text            =   "Listado de Nominas"
      End
      Begin VB.CommandButton CmdSalir 
         DownPicture     =   "Frm13VacaMes.frx":0B1A
         Height          =   375
         Left            =   10920
         Picture         =   "Frm13VacaMes.frx":25FC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6600
         Width           =   1455
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5415
         Left            =   0
         TabIndex        =   2
         Top             =   1080
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   9551
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   12632319
         TabCaption(0)   =   "Vacaciones"
         TabPicture(0)   =   "Frm13VacaMes.frx":40DE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label3"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label7"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label13"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "PushButton1"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "PBVacaciones"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "TxtFechaAplica"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "dtpFPInicio"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "dtpFPFinal"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "SmartButton2"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "CmdCalVaca"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "TxtNumNomVaca"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "CmdCerrarVacaciones"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "TxtDiasDescuento"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "CmdPRVaca"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "TxtFINIVaca"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "TxtFFinVaca"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Command2"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "DbgrVacaciones"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "CmdExportar"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "CmdMonedasvaca"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "CmdColillaVaca"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "CmdNominaVaca"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "CHKTranferir"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "ChkRestar"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "CmdCalcularVacaciones"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "ChkExtraVaca"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "ChkEliminar"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "ChkImprimirDptoVaca"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).ControlCount=   30
         TabCaption(1)   =   "Trecavo Mes"
         TabPicture(1)   =   "Frm13VacaMes.frx":40FA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label6"
         Tab(1).Control(1)=   "Label5"
         Tab(1).Control(2)=   "Label2"
         Tab(1).Control(3)=   "CmdExportaBanpro"
         Tab(1).Control(4)=   "CmdExportaBac"
         Tab(1).Control(5)=   "PB13Mes"
         Tab(1).Control(6)=   "CmdprNomina"
         Tab(1).Control(7)=   "CmdPrnNomina"
         Tab(1).Control(8)=   "CmdCal13"
         Tab(1).Control(9)=   "TxtNumNom13"
         Tab(1).Control(10)=   "CmdCerrar13"
         Tab(1).Control(11)=   "TxtFINI13"
         Tab(1).Control(12)=   "TxtFFIN13"
         Tab(1).Control(13)=   "Command1"
         Tab(1).Control(14)=   "CmdExporta2"
         Tab(1).Control(15)=   "Dbgr13Mes"
         Tab(1).Control(16)=   "SmartButton1"
         Tab(1).Control(17)=   "CmdDenominacion"
         Tab(1).Control(18)=   "Frame1"
         Tab(1).Control(19)=   "ChkColillaDpto"
         Tab(1).ControlCount=   20
         Begin VB.CheckBox ChkImprimirDptoVaca 
            Caption         =   "Imprimir por Dpto"
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Top             =   2640
            Width           =   1695
         End
         Begin VB.CheckBox ChkColillaDpto 
            Caption         =   "Colilla por Dpto"
            Height          =   255
            Left            =   -74760
            TabIndex        =   62
            Top             =   3240
            Width           =   1695
         End
         Begin VB.CheckBox ChkEliminar 
            Caption         =   "Eliminar el Calculo Anterior"
            Height          =   255
            Left            =   8640
            TabIndex        =   61
            Top             =   600
            Width           =   2175
         End
         Begin VB.CheckBox ChkExtraVaca 
            Caption         =   "Calcular Horas Extra"
            Height          =   255
            Left            =   8640
            TabIndex        =   60
            Top             =   960
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CommandButton CmdCalcularVacaciones 
            DownPicture     =   "Frm13VacaMes.frx":4116
            Height          =   375
            Left            =   -120
            Picture         =   "Frm13VacaMes.frx":5BF8
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   4920
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox ChkRestar 
            Caption         =   "Restar Inss Nomina"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   3315
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.Frame Frame1 
            Caption         =   "Fechas Segun Periodos"
            Height          =   855
            Left            =   -70680
            TabIndex        =   45
            Top             =   480
            Width           =   7935
            Begin MSComCtl2.DTPicker DtpFin13vo 
               Height          =   315
               Left            =   5280
               TabIndex        =   46
               Top             =   360
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               Format          =   17104897
               CurrentDate     =   38305
            End
            Begin MSComCtl2.DTPicker DtpInicio13vo 
               Height          =   315
               Left            =   1800
               TabIndex        =   47
               Top             =   360
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               Format          =   17104897
               CurrentDate     =   38305
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               Caption         =   "Fecha Final"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   49
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               Caption         =   "Fecha Inicial"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   48
               Top             =   360
               Width           =   1815
            End
         End
         Begin VB.CheckBox CHKTranferir 
            Caption         =   "Cerrar sin Tranferir"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   3000
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CommandButton CmdNominaVaca 
            DownPicture     =   "Frm13VacaMes.frx":773A
            Enabled         =   0   'False
            Height          =   375
            Left            =   9360
            Picture         =   "Frm13VacaMes.frx":921C
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton CmdColillaVaca 
            DownPicture     =   "Frm13VacaMes.frx":ACFE
            Enabled         =   0   'False
            Height          =   375
            Left            =   7920
            Picture         =   "Frm13VacaMes.frx":C7E0
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton CmdMonedasvaca 
            DownPicture     =   "Frm13VacaMes.frx":E2C2
            Enabled         =   0   'False
            Height          =   375
            Left            =   6480
            Picture         =   "Frm13VacaMes.frx":FDA4
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton CmdExportar 
            DownPicture     =   "Frm13VacaMes.frx":116A6
            Enabled         =   0   'False
            Height          =   375
            Left            =   10800
            Picture         =   "Frm13VacaMes.frx":13188
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton CmdDenominacion 
            DownPicture     =   "Frm13VacaMes.frx":14A32
            Enabled         =   0   'False
            Height          =   375
            Left            =   -68520
            Picture         =   "Frm13VacaMes.frx":16514
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   4920
            Width           =   1455
         End
         Begin SmartButtonProject.SmartButton SmartButton1 
            Height          =   975
            Left            =   -74520
            TabIndex        =   32
            Top             =   3600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1720
            Caption         =   "Historial Manual"
            Picture         =   "Frm13VacaMes.frx":17E16
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
         Begin TrueOleDBGrid70.TDBGrid Dbgr13Mes 
            Bindings        =   "Frm13VacaMes.frx":186F0
            Height          =   3015
            Left            =   -72960
            TabIndex        =   31
            Top             =   1440
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   5318
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid DbgrVacaciones 
            Bindings        =   "Frm13VacaMes.frx":18709
            Height          =   2535
            Left            =   2160
            TabIndex        =   30
            Top             =   1680
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   4471
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=255,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin VB.CommandButton CmdExporta2 
            DownPicture     =   "Frm13VacaMes.frx":18725
            Enabled         =   0   'False
            Height          =   375
            Left            =   -64200
            Picture         =   "Frm13VacaMes.frx":1A207
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            DownPicture     =   "Frm13VacaMes.frx":1BAB1
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -64200
            Picture         =   "Frm13VacaMes.frx":1D593
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   4560
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            DownPicture     =   "Frm13VacaMes.frx":1F075
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   10800
            Picture         =   "Frm13VacaMes.frx":20B57
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   4440
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker TxtFFIN13 
            Height          =   315
            Left            =   -72720
            TabIndex        =   25
            Top             =   840
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   17104897
            CurrentDate     =   38305
         End
         Begin MSComCtl2.DTPicker TxtFINI13 
            Height          =   315
            Left            =   -74760
            TabIndex        =   24
            Top             =   840
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   17104897
            CurrentDate     =   38305
         End
         Begin MSComCtl2.DTPicker TxtFFinVaca 
            Height          =   315
            Left            =   2400
            TabIndex        =   23
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   17104897
            CurrentDate     =   38305
         End
         Begin MSComCtl2.DTPicker TxtFINIVaca 
            Height          =   315
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   17104897
            CurrentDate     =   38305
         End
         Begin VB.CommandButton CmdPRVaca 
            DownPicture     =   "Frm13VacaMes.frx":22639
            Enabled         =   0   'False
            Height          =   375
            Left            =   -120
            Picture         =   "Frm13VacaMes.frx":2411B
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   4920
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtDiasDescuento 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1560
            TabIndex        =   11
            Text            =   "0"
            Top             =   2280
            Width           =   375
         End
         Begin VB.CommandButton CmdCerrar13 
            DownPicture     =   "Frm13VacaMes.frx":25BFD
            Height          =   375
            Left            =   -65640
            Picture         =   "Frm13VacaMes.frx":276DF
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   4560
            Width           =   1455
         End
         Begin VB.CommandButton CmdCerrarVacaciones 
            DownPicture     =   "Frm13VacaMes.frx":28FE1
            Height          =   375
            Left            =   9360
            Picture         =   "Frm13VacaMes.frx":2AAC3
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   4440
            Width           =   1455
         End
         Begin VB.TextBox TxtNumNom13 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   420
            Left            =   -74520
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox TxtNumNomVaca 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   420
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1800
            Width           =   1455
         End
         Begin VB.CommandButton CmdCal13 
            DownPicture     =   "Frm13VacaMes.frx":2C3C5
            Height          =   375
            Left            =   -67080
            Picture         =   "Frm13VacaMes.frx":2DEA7
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   4560
            Width           =   1455
         End
         Begin VB.CommandButton CmdCalVaca 
            DownPicture     =   "Frm13VacaMes.frx":2F9E9
            Height          =   375
            Left            =   7920
            Picture         =   "Frm13VacaMes.frx":314CB
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   4440
            UseMaskColor    =   -1  'True
            Width           =   1455
         End
         Begin VB.CommandButton CmdPrnNomina 
            DownPicture     =   "Frm13VacaMes.frx":3300D
            Enabled         =   0   'False
            Height          =   375
            Left            =   -67080
            Picture         =   "Frm13VacaMes.frx":34AEF
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton CmdprNomina 
            DownPicture     =   "Frm13VacaMes.frx":365D1
            Enabled         =   0   'False
            Height          =   375
            Left            =   -65640
            Picture         =   "Frm13VacaMes.frx":380B3
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   4920
            Width           =   1455
         End
         Begin SmartButtonProject.SmartButton SmartButton2 
            Height          =   975
            Left            =   360
            TabIndex        =   33
            Top             =   3720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1720
            Caption         =   "Historial Manual"
            Picture         =   "Frm13VacaMes.frx":39B95
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
         Begin MSComCtl2.DTPicker dtpFPFinal 
            Height          =   315
            Left            =   2400
            TabIndex        =   40
            Top             =   1200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   17104897
            CurrentDate     =   38305
         End
         Begin MSComCtl2.DTPicker dtpFPInicio 
            Height          =   315
            Left            =   240
            TabIndex        =   41
            Top             =   1200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   17104897
            CurrentDate     =   38305
         End
         Begin MSComCtl2.DTPicker TxtFechaAplica 
            Height          =   315
            Left            =   5880
            TabIndex        =   52
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   17104897
            CurrentDate     =   38305
         End
         Begin XtremeSuiteControls.ProgressBar PBVacaciones 
            Height          =   375
            Left            =   120
            TabIndex        =   55
            Top             =   4920
            Width           =   6135
            _Version        =   786432
            _ExtentX        =   10821
            _ExtentY        =   661
            _StockProps     =   93
            BackColor       =   14737632
            Scrolling       =   1
            Appearance      =   6
         End
         Begin XtremeSuiteControls.ProgressBar PB13Mes 
            Height          =   375
            Left            =   -74880
            TabIndex        =   56
            Top             =   4800
            Width           =   6255
            _Version        =   786432
            _ExtentX        =   11033
            _ExtentY        =   661
            _StockProps     =   93
            BackColor       =   14737632
            Scrolling       =   1
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton CmdExportaBac 
            Height          =   375
            Left            =   -74760
            TabIndex        =   57
            Top             =   1560
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Exportar"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "Frm13VacaMes.frx":3A46F
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton CmdExportaBanpro 
            Height          =   375
            Left            =   -74760
            TabIndex        =   58
            Top             =   2040
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Exportar"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "Frm13VacaMes.frx":3ABF3
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   375
            Left            =   2160
            TabIndex        =   64
            Top             =   4320
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   " Exportar"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "Frm13VacaMes.frx":3B4CF
            ImageAlignment  =   0
         End
         Begin VB.Label Label13 
            Caption         =   "Fecha Aplica"
            Height          =   255
            Left            =   4800
            TabIndex        =   51
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Días de Descuento"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   18
            Top             =   2640
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Fecha Inicial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -74880
            TabIndex        =   16
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Fecha Final"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -72840
            TabIndex        =   15
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Fecha Inicial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Fecha Final"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   13
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Label Label10 
         Caption         =   "dias"
         Height          =   255
         Left            =   10560
         TabIndex        =   44
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Antiguedad mayor a:"
         Height          =   255
         Left            =   8040
         TabIndex        =   42
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Listado de Nominas"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label LblTotal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   6240
         Width           =   6855
      End
   End
   Begin MSAdodcLib.Adodc AdoBusca 
      Height          =   375
      Left            =   5040
      Top             =   8880
      Width           =   3615
      _ExtentX        =   6376
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
      Caption         =   "AdoBusca"
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
Attribute VB_Name = "Frm13VacaMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public Sub LlenarHistorialAguinaldo(CodEmpleado As Double)
 Dim Meses As Double, CalculaHE As Boolean, SalTempHE As Double, SalTemp As Double, Ingresos As Double
 Dim HorasExtra As Double
 
 
        Anno = Year(Me.TxtFFIN13)
        MesActual = Month(Me.TxtFFIN13)

        If MesActual > 6 Then
            Meses = MesActual - 5
        Else
           Meses = 1
        End If
        

        
   CalculaHE = True
   SalTempHE = 0
   For Mes = Meses To MesActual
        'SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo] + [DetalleNomina].[SeptimoDia] + [DetalleNomina].[OtrosIngresos] + [DetalleNomina].[Comisiones] + [DetalleNomina].[IncetivoProduccion] + [DetalleNomina].[Antiguedad] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where (Nomina.Mes =" & Mes & ") And (Nomina.Ano =" & Anno & ") and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
        'SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where ((Month([Nomina].[FechaNomina])) =  '" & Mes & "' ) And ((Year([Nomina].[FechaNomina])) ='" & Anno & "') and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
        
       If CodEmpleado = "10709" Or CodEmpleado = "10796" Then
             SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.OtrosIngresos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo] + [DetalleNomina].[Antiguedad] + DetalleNomina.Incentivos + DetalleNomina.OtrosIngresos  AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where (Nomina.Mes =" & Mes & ") And (Nomina.Ano =" & Anno & ") and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
       Else
              SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.OtrosIngresos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo] + [DetalleNomina].[Antiguedad] + DetalleNomina.Incentivos + DetalleNomina.Comisiones + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos   AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where (Nomina.Mes =" & Mes & ") And (Nomina.Ano =" & Anno & ") and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
       End If
            DtaNominas.RecordSource = SqlNominas
            DtaNominas.Refresh
               SalTemp = 0
               Ingresos = 0
               HorasExtra = 0
               SalTempHE = 0
         Do While Not DtaNominas.Recordset.EOF
            
              
               HorasExtra = HorasExtra + DtaNominas.Recordset("HorasExtras")
               SalTemp = SalTemp + DtaNominas.Recordset("Total")
               SalTempHE = SalTempHE + DtaNominas.Recordset("Total") + DtaNominas.Recordset("HorasExtras")
               SalTemp = SalTempHE  '''PARA QUE CALCULE CON HORAS EXTRAS
               Ingresos = Ingresos + (DtaNominas.Recordset("Destajo") + DtaNominas.Recordset("OtrosIngresos") + DtaNominas.Recordset("Incentivos"))
               
              
              '/////////////////////////////////////////////////////////////////////////////////////
              '//////////////////////////////RECORRO LA TABLA DE PERIODO PARA SELECCIONAR LOS MESES /
              '///////////////////////////////////////////////////////////////////////////////////////
              If Me.DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
                Me.DtaConsulta.RecordSource = "SELECT año, CodTipoNomina, COUNT(mes) AS Cont, mes AS Mes From Fecha_Planilla GROUP BY año, CodTipoNomina, mes HAVING (año = " & Anno & ") AND (CodTipoNomina = '" & CodTipoNomina & "')  AND (mes = '" & Format(Mes, "0#") & "')" 'AND (COUNT(mes) = 3)
                Me.DtaConsulta.Refresh
                Do While Not Me.DtaConsulta.Recordset.EOF
                   If Me.DtaConsulta.Recordset("Cont") = 3 Then
                     SalTemp = (SalTemp / 42) * 30
                   Else
                     SalTemp = (SalTemp / 28) * 30
                 
                   End If
                 
                  
                  Me.DtaConsulta.Recordset.MoveNext
                Loop
              End If


               
               CantRegistros = CantRegistros + 1
               DtaNominas.Recordset.MoveNext
              

            Loop
              
              If HorasExtra <= 0 Then
                CalculaHE = False
              End If
              
              
              '///////////////////////////////////////////////////////////////////////////////////
              '//////////////////////////SI EL SALARIO ES BASICO NO SE TOMA EN CUENTA PROMEDIO/////////////
              '///////////////////////////////////////////////////////////////////////////////////
              MDIPrimero.DtaConsulta.RecordSource = "SELECT  Historico.*, Empleado.SueldoActualBasico FROM Historico INNER JOIN Empleado ON Historico.Codempleado = Empleado.CodEmpleado  Where (Historico.CodEmpleado = " & CodEmpleado & ")"
              MDIPrimero.DtaConsulta.Refresh
              If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
                If MDIPrimero.DtaConsulta.Recordset("SueldoActualBasico") = True Then
                 If Not IsNull(MDIPrimero.DtaConsulta.Recordset("SueldoActual")) Then
                    SalarioBasico = MDIPrimero.DtaConsulta.Recordset("SueldoActual")
                 Else
                    SalarioBasico = 0
                 End If
                    SalTemp = SalarioBasico + Ingresos
                End If
              End If
              
                '///////////////////////Busco si el Empleado tiene Subsidio para Sumarlo////////////////////
  
                 'Me.DtaConsulta.RecordSource = "SELECT NomSubsidio.NumNomina, NomSubsidio.FechaPago, DetalleNomSubsidio.CodEmpleado, DetalleNomSubsidio.Subsidio FROM NomSubsidio INNER JOIN DetalleNomSubsidio ON NomSubsidio.NumNomina = DetalleNomSubsidio.NumNominaSubsidio WHERE (((NomSubsidio.FechaPago) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((DetalleNomSubsidio.CodEmpleado)='" & CodEmpleado & "'))"
                 Me.DtaConsulta.RecordSource = "SELECT NomSubsidio.NumNomina, DetalleNomSubsidio.CodEmpleado, DetalleNomSubsidio.Subsidio, Empleado.SumarSubsidio, Year([NomSubsidio].[FechaPago]) AS Anno, Month([NomSubsidio].[FechaPago]) AS Mes FROM NomSubsidio INNER JOIN (Empleado INNER JOIN DetalleNomSubsidio ON Empleado.CodEmpleado = DetalleNomSubsidio.CodEmpleado) ON NomSubsidio.NumNomina = DetalleNomSubsidio.NumNominaSubsidio Where (((DetalleNomSubsidio.CodEmpleado) = '" & CodEmpleado & "') And ((Empleado.SumarSubsidio) = 1) And ((Year([NomSubsidio].[FechaPago])) = '" & Anno & "') And ((Month([NomSubsidio].[FechaPago])) = '" & Mes & "')) ORDER BY NomSubsidio.NumNomina"
                 Me.DtaConsulta.Refresh
                 MontoSubsidio = 0
                 Do While Not Me.DtaConsulta.Recordset.EOF
                  MontoSubsidio = MontoSubsidio + Me.DtaConsulta.Recordset("Subsidio")
                  Me.DtaConsulta.Recordset.MoveNext
                 Loop
              
                 SalTemp = SalTemp + MontoSubsidio
                 Fecha1 = Year(Me.TxtFINI13.Value) & "-" & Month(Me.TxtFINI13.Value) & "-" & Day(Me.TxtFINI13.Value)
                 Fecha2 = Year(Me.TxtFFIN13.Value) & "-" & Month(Me.TxtFFIN13.Value) & "-" & Day(Me.TxtFFIN13.Value)
                 
               
                 
               If SalTemp <> 0 Then
                 Me.AdoHistorialSalarial.RecordSource = "SELECT NumNomina, Tipo, CodEmpleado, FechaIni, FechaFin, Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, Agosto, Septiembre, Octubre, Noviembre, Diciembre From HistorialSalarioMes WHERE     (CodEmpleado = '" & CodEmpleado & "') AND (FechaIni = CONVERT(DATETIME, '" & Fecha1 & "', 102)) AND (FechaFin = CONVERT(DATETIME, '" & Fecha2 & "',102))"
                 Me.AdoHistorialSalarial.Refresh
                  If Me.AdoHistorialSalarial.Recordset.EOF Then
                        Me.AdoHistorialSalarial.Recordset.AddNew
                        Me.AdoHistorialSalarial.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
                        Me.AdoHistorialSalarial.Recordset("FechaIni") = CDate(Me.TxtFINI13.Value)
                        Me.AdoHistorialSalarial.Recordset("FechaFin") = CDate(Me.TxtFFIN13.Value)
                        Me.AdoHistorialSalarial.Recordset("NumNomina") = val(Me.TxtNumNom13.Text)
                        Me.AdoHistorialSalarial.Recordset("Tipo") = "Aguinaldo"
                        Select Case Mes
                          Case 1
                            Me.AdoHistorialSalarial.Recordset("Enero") = SalTemp
                          Case 2
                            Me.AdoHistorialSalarial.Recordset("Febrero") = SalTemp
                          Case 3
                            Me.AdoHistorialSalarial.Recordset("Marzo") = SalTemp
                          Case 4
                            Me.AdoHistorialSalarial.Recordset("Abril") = SalTemp
                          Case 5
                            Me.AdoHistorialSalarial.Recordset("Mayo") = SalTemp
                          Case 6
                            Me.AdoHistorialSalarial.Recordset("Junio") = SalTemp
                          Case 7
                            Me.AdoHistorialSalarial.Recordset("Julio") = SalTemp
                          Case 8
                            Me.AdoHistorialSalarial.Recordset("Agosto") = SalTemp
                          Case 9
                            Me.AdoHistorialSalarial.Recordset("Septiembre") = SalTemp
                          Case 10
                            Me.AdoHistorialSalarial.Recordset("Octubre") = SalTemp
                          Case 11
                            Me.AdoHistorialSalarial.Recordset("Noviembre") = SalTemp
                          Case 12
                            Me.AdoHistorialSalarial.Recordset("Diciembre") = SalTemp
                         End Select
                        Me.AdoHistorialSalarial.Recordset.Update
                  Else
                         Me.AdoHistorialSalarial.Recordset("FechaIni") = CDate(Me.TxtFINI13.Value)
                         Me.AdoHistorialSalarial.Recordset("FechaFin") = CDate(Me.TxtFFIN13.Value)
                         Me.AdoHistorialSalarial.Recordset("NumNomina") = val(Me.TxtNumNom13.Text)
                         Me.AdoHistorialSalarial.Recordset("Tipo") = "Aguinaldo"
                        Select Case Mes
                          Case 1
                            Me.AdoHistorialSalarial.Recordset("Enero") = SalTemp
                          Case 2
                            Me.AdoHistorialSalarial.Recordset("Febrero") = SalTemp
                          Case 3
                            Me.AdoHistorialSalarial.Recordset("Marzo") = SalTemp
                          Case 4
                            Me.AdoHistorialSalarial.Recordset("Abril") = SalTemp
                          Case 5
                            Me.AdoHistorialSalarial.Recordset("Mayo") = SalTemp
                          Case 6
                            Me.AdoHistorialSalarial.Recordset("Junio") = SalTemp
                          Case 7
                            Me.AdoHistorialSalarial.Recordset("Julio") = SalTemp
                          Case 8
                            Me.AdoHistorialSalarial.Recordset("Agosto") = SalTemp
                          Case 9
                            Me.AdoHistorialSalarial.Recordset("Septiembre") = SalTemp
                          Case 10
                            Me.AdoHistorialSalarial.Recordset("Octubre") = SalTemp
                          Case 11
                            Me.AdoHistorialSalarial.Recordset("Noviembre") = SalTemp
                          Case 12
                            Me.AdoHistorialSalarial.Recordset("Diciembre") = SalTemp
                        End Select
                        Me.AdoHistorialSalarial.Recordset.Update
       
                  End If
            
            Else
'                 Me.AdoHistorialSalarial.RecordSource = "SELECT NumNomina, CodEmpleado, FechaIni, FechaFin, Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, Agosto, Septiembre, Octubre, Noviembre, Diciembre From HistorialSalarioMes WHERE  (CodEmpleado = '" & CodEmpleado & "')"
'                 Me.AdoHistorialSalarial.Refresh
'                  If Not Me.AdoHistorialSalarial.Recordset.EOF Then
'
'                        Select Case Mes
'                          Case 1
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Enero")
'                          Case 2
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Febrero")
'                          Case 3
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Marzo")
'                          Case 4
'                           SalTemp = Me.AdoHistorialSalarial.Recordset("Abril")
'                          Case 5
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Mayo")
'                          Case 6
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Junio")
'                          Case 7
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Julio")
'                          Case 8
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Agosto")
'                          Case 9
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Septiembre")
'                          Case 10
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Octubre")
'                          Case 11
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Noviembre")
'                          Case 12
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Diciembre")
'                         End Select
'
'                 End If
            
            End If
            
        If SalMayor < SalTemp Then
           SalMayor = SalTemp
        End If

        If SalMayorCHE < SalTempHE Then
            SalMayorCHE = SalTempHE
        End If
        
        
        
  
   Next



End Sub

 
 
 
Private Sub CmdCal13_Click()
Dim M1 As Integer, M2 As Integer
Dim SqlEmpleados As String, Meses As Double
Dim CodTipoNomina As String
Dim SqlNominas As String, FechaContrato As Date
Dim SalMayor As Double, NumFecha1 As Long
Dim SalTemp As Double, NumFecha2 As Long
Dim CodEmpleado As String
Dim Edicion As Boolean, CantMeses As Double
Dim Anno As Integer, TarifaHoraria As Double
Dim Mes As Integer, MesActual As Integer
Dim CantEmpleados As Long
Dim HorasExtra As Double, CalculaHE As Boolean, SalTempHE As Double
Dim i As Integer, CantRegistros As Integer
Dim Dias As Double, annos As Double, Ingresos As Double, PorcentajePension As Double, MontoPension As Double
Dim Adelanto13vo As Double, MontoSubsidio As Double, SalPorciento As Double, SalMayorCHE As Double
Anno = Year(Me.TxtFFIN13)
MesActual = Month(Me.TxtFFIN13)

NumFecha1 = CDate(Me.TxtFINI13)
NumFecha2 = CDate(Me.TxtFFIN13)

Edicion = False

MousePointer = 11
If DBCNominas.Text = "Listado de Nominas" Then
   MsgBox "No ha seleccionado el tipo de nomina al cual le desea calcular las Vacaciones"
   MousePointer = 1
   DBCNominas.SetFocus
   Exit Sub
End If




DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop

NumNom13Mes = val(Me.TxtNumNom13.Text)
'pregunto si ya existe una nomina de 13vo mes elaborada


Me.DtaNom13Mes.RecordSource = "SELECT CodTipoNomina, NumNom13Mes, FechaAplica, MontoPagado, FechaIni, FechaFin, Activa From Nom13Mes Where (NumNom13Mes = " & NumNom13Mes & ") AND (Activa = 1)"
DtaNom13Mes.Refresh
If Not Me.DtaNom13Mes.Recordset.EOF Then

    If DtaNom13Mes.Recordset("NumNom13Mes") = NumNom13Mes And DtaNom13Mes.Recordset("Activa") = True Then
       'DtaNom13Mes.Recordset.Edit
       DtaNom13Mes.Recordset("fechaaplica") = Format(Now, "DD/MM/YYYY")
       DtaNom13Mes.Recordset("montopagado") = 0
       DtaNom13Mes.Recordset("Fechaini") = CDate(Me.TxtFINI13.Value)
       DtaNom13Mes.Recordset("Fechafin") = CDate(Me.TxtFFIN13.Value)
       DtaNom13Mes.Recordset.Update
       Edicion = True
    
    End If
End If




If Not Edicion Then 'es primera vez que se crearan los 13vo mes

       DtaNom13Mes.Recordset.AddNew
       DtaNom13Mes.Recordset("NumNom13Mes") = val(TxtNumNom13.Text)
       DtaNom13Mes.Recordset("fechaaplica") = Format(Now, "DD/MM/YYYY")
       DtaNom13Mes.Recordset("montopagado") = 0
       DtaNom13Mes.Recordset("Fechaini") = CDate(Me.TxtFINI13.Value)
       DtaNom13Mes.Recordset("Fechafin") = CDate(Me.TxtFFIN13.Value)
       DtaNom13Mes.Recordset("Activa") = 1
       DtaNom13Mes.Recordset("CodTipoNomina") = CodTipoNomina
       DtaNom13Mes.Recordset.Update
       
Else 'se borran los movimientos anteriores
     Me.DtaDetalleNom13Mes.RecordSource = "SELECT Id, NumNom13Mes, CodEmpleado, SalarioMensual, SalarioAPagar, DiasAPagar, Adelanto13vo From DetalleNom13Mes Where (NumNom13Mes = " & NumNom13Mes & ")"
     DtaDetalleNom13Mes.Refresh
     Do While Not DtaDetalleNom13Mes.Recordset.EOF
     If DtaDetalleNom13Mes.Recordset("NumNom13Mes") = val(TxtNumNom13) Then
        DtaDetalleNom13Mes.Recordset.Delete
     End If
     
     DtaDetalleNom13Mes.Recordset.MoveNext
     Loop
     
     
End If

'hago el sql de los empleados que pertenezcan solo a la nomina seleccionada

'SqlEmpleados = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.CodTipoNomina From Empleado WHERE Empleado.CodTipoNomina=  '" & CodTipoNomina & "'"
SqlEmpleados = "SELECT Empleado.SalarioFijo, Empleado.CodEmpleado1, Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, Empleado.ExentoIr, Empleado.OtrosIngresos,  Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, Empleado.Ausente, Empleado.PorcentajePension, Empleado.MontoPension From Empleado WHERE Empleado.CodTipoNomina= '" & CodTipoNomina & "' AND Empleado.Activo=1 "
DtaEmpleados.RecordSource = SqlEmpleados
DtaEmpleados.Refresh

DtaEmpleados.Recordset.MoveLast
CantEmpleados = DtaEmpleados.Recordset.RecordCount
Dim CodEmpleado1 As String
With PB13Mes
.Min = 0
.Value = 0
.Max = CantEmpleados

i = 1
DtaEmpleados.Refresh
'recorro la BD empleados y a cada uno le busco su salario mayor (si es destajo) si no solo extraigo su salario
Do While Not DtaEmpleados.Recordset.EOF
MontoPension = 0
PorcentajePension = 0

Dias = 0
    
  CodEmpleado = DtaEmpleados.Recordset("CodEmpleado")
  If Not IsNull(DtaEmpleados.Recordset("MontoPension")) Then
    MontoPension = DtaEmpleados.Recordset("MontoPension")
  Else
   MontoPension = 0
  End If
  
  If Not IsNull(DtaEmpleados.Recordset("PorcentajePension")) Then
    PorcentajePension = DtaEmpleados.Recordset("PorcentajePension")
  End If
  CodEmpleado1 = DtaEmpleados.Recordset("CodEmpleado1")
  
  If CodEmpleado1 = "S120110101" Then
   CodEmpleado1 = "S120110101"
  End If

  DtaHistorico.RecordSource = "SELECT Historico.Codempleado, Historico.FechaBaja, Historico.FechaContrato From Historico Where (((Historico.CodEmpleado) = '" & CodEmpleado & "'))"
  DtaHistorico.Refresh
  NumFecha1 = CDate(Me.TxtFINI13)
  NumFecha2 = CDate(Me.TxtFFIN13)
    '/////////////Busco el Adelanto de 13vo mes Registrados//////////////
    '////////////////////////////////////////////////////////////////////////
        Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo WHERE (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Adelanto13vo.FechaAdelanto) Between  " & NumFecha1 & " And " & NumFecha2 & ") AND ((Adelanto13vo.TipoAdelanto)='13vo Mes'))"
        'Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo WHERE (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Adelanto13vo.FechaAdelanto) Between " & NumFecha1 & " And " & NumFecha2 & "))"
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
'     NumFecha2 = FechaContrato
'     NumFecha1 = Me.TxtFFIN13.Value
'     annos = (CDbl(NumFecha1) - CDbl(NumFecha2) + 1) / 365
'     CantMeses = annos * 12
'
'     If CantMeses < 0 Then
'      Dias = 0
'     ElseIf CantMeses <= 12 Then
''       Dias = Format(CantMeses * (DiasMes / 12), "##,##0.0000")
'       Dias = (CalcularDiasVaca(CDate(FechaContrato), Me.TxtFFIN13.Value)) * 0.083333333333
'     Else
'
''      Dias = DiasMes
'       Dias = 30
'
'     End If
'
'     If Dias > 30 Then
'        Dias = 30
'     End If
'
     

     
     If FechaContrato < Me.TxtFINI13.Value Then
         Dias = CalcularDiasAguinaldo(CodEmpleado1, Me.TxtFINI13.Value, Me.TxtFFIN13.Value)
     ElseIf DateDiff("d", FechaContrato, Me.TxtFFIN13.Value) <= 2 Then
        Dias = DateDiff("d", FechaContrato, Me.TxtFFIN13.Value) + 1 * 0.08333
     Else
       Dias = CalcularDiasAguinaldo(CodEmpleado1, FechaContrato, Me.TxtFFIN13.Value)
     End If
     
    End If
  End If
  
  
    

      '   MsgBox "01"
       'End If
   '  .Value = I
     'Me.xp_canvas1.Caption = "Procesando:  " & i & " de " & CantEmpleados & " Empleados "
     Me.Caption = "Procesando:  " & i & " de " & CantEmpleados & " Empleados "
      Me.LblTotal.Caption = "Procesando:  " & i & " de " & CantEmpleados & " Empleados "
     DoEvents
    'tengo que hacer un SQL de Solo los que esten en el rango de fechas
    'solo se veran las nominas de cada mes
    'se deben de hacer ciclos por cada mes, seis ciclos por los seis meses.
     CantMeses = 0
     CantRegistros = 0
     TarifaHoraria = DtaEmpleados.Recordset("TarifaHoraria")
      
     If Me.DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
    
        If Me.DtaTipoNomina.Recordset("TipoPago") = "Salario Fijo" Then
           TarifaHoraria = Format(DtaEmpleados.Recordset("SueldoPeriodo") / 112, "###,##0.000000")
        Else
          TarifaHoraria = DtaEmpleados.Recordset("TarifaHoraria")
        End If
     End If
     
'     TarifaHoraria = DtaEmpleados.Recordset("TarifaHoraria")
     SalMayor = Format(DiasMes * 8 * TarifaHoraria, "##,##0.00")
      
'///Verifico si tiene salario fijo///////////
  If Not DtaEmpleados.Recordset.EOF Then
    If DtaEmpleados.Recordset("SalarioFijo") = "N" Then
        'extraigo el salario mayor de Diciembre del año pasado
            'SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where ((Month([Nomina].[FechaNomina])) =  '" & 12 & "' ) And ((Year([Nomina].[FechaNomina])) ='" & Anno - 1 & "') and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
             '   DtaNominas.RecordSource = SqlNominas
              '  DtaNominas.Refresh
               '    SalTemp = 0
                  'If Not DtaNominas.Recordset.EOF Then
                   '   CantMeses = CantMeses + 1
                  'End If
                '  Do While Not DtaNominas.Recordset.EOF
                 '  SalTemp = SalTemp + DtaNominas.Recordset("Total")
                   
                  ' DtaNominas.Recordset.MoveNext
                  'Loop
           ' If SalMayor < SalTemp Then
            '   SalMayor = SalTemp
            'End If
   End If
 End If
'/////////////extraigo el salario mayor de enero  los
'////////////////Ultimos 6 meses

'///Verifico si tiene salario fijo///////////

      If CodEmpleado = "41133" Then
         CodEmpleado = "41133"
      End If

If Dias >= CDbl(Me.txtAntiguedad.Text) Then

  
 If DtaEmpleados.Recordset("SalarioFijo") = "N" Then
        If MesActual > 6 Then
            Meses = MesActual - 5
        Else
           Meses = 1
        End If
        

        
       CalculaHE = True
       SalTempHE = 0
        For Mes = Meses To MesActual
        'SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo] + [DetalleNomina].[SeptimoDia] + [DetalleNomina].[OtrosIngresos] + [DetalleNomina].[Comisiones] + [DetalleNomina].[IncetivoProduccion] + [DetalleNomina].[Antiguedad] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where (Nomina.Mes =" & Mes & ") And (Nomina.Ano =" & Anno & ") and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
        'SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where ((Month([Nomina].[FechaNomina])) =  '" & Mes & "' ) And ((Year([Nomina].[FechaNomina])) ='" & Anno & "') and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
        
       If CodEmpleado = "10709" Or CodEmpleado = "10796" Then
             SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.OtrosIngresos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo] + [DetalleNomina].[Antiguedad] + DetalleNomina.Incentivos + DetalleNomina.OtrosIngresos  AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where (Nomina.Mes =" & Mes & ") And (Nomina.Ano =" & Anno & ") and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
       Else
              SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.OtrosIngresos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo] + [DetalleNomina].[Antiguedad] + DetalleNomina.Incentivos + DetalleNomina.Comisiones + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos   AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where (Nomina.Mes =" & Mes & ") And (Nomina.Ano =" & Anno & ") and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
       End If
            DtaNominas.RecordSource = SqlNominas
            DtaNominas.Refresh
               SalTemp = 0
               Ingresos = 0
               HorasExtra = 0
               SalTempHE = 0
              Do While Not DtaNominas.Recordset.EOF
            
              
               HorasExtra = HorasExtra + DtaNominas.Recordset("HorasExtras")
               SalTemp = SalTemp + DtaNominas.Recordset("Total")
               SalTempHE = SalTempHE + DtaNominas.Recordset("Total") + DtaNominas.Recordset("HorasExtras")
               SalTemp = SalTempHE  '''PARA QUE CALCULE CON HORAS EXTRAS
               Ingresos = Ingresos + (DtaNominas.Recordset("Destajo") + DtaNominas.Recordset("OtrosIngresos") + DtaNominas.Recordset("Incentivos"))
               
'               '--------------------------BUSCO LOS INCENTIVOS SIN INCLUIR EL INCENTIVO EXCENTO 14 -------------------------
'               Me.DtaConsulta.RecordSource = "SELECT DetalleIncentivo.Id, DetalleIncentivo.NumIncentivo, DetalleIncentivo.Valor, DetalleIncentivo.NumVez, DetalleIncentivo.Pagado, DetalleIncentivo.NumNomina, Incentivo.CodTipoIncentivo FROM DetalleIncentivo INNER JOIN Incentivo ON DetalleIncentivo.NumIncentivo = Incentivo.NumIncentivo  " & _
'                                             "WHERE  (DetalleIncentivo.NumNomina = " & DtaNominas.Recordset("NumNomina") & ") AND (Incentivo.CodTipoIncentivo <> '14') AND (Incentivo.CodEmpleado = " & CodEmpleado & ")"
'
'               Me.DtaConsulta.Refresh
'               Do While Not DtaConsulta.Recordset.EOF
'                   SalTemp = SalTemp + Me.DtaConsulta.Recordset("Valor")
'                   Me.DtaConsulta.Recordset.MoveNext
'               Loop

              '/////////////////////////////////////////////////////////////////////////////////////
              '//////////////////////////////RECORRO LA TABLA DE PERIODO PARA SELECCIONAR LOS MESES /
              '///////////////////////////////////////////////////////////////////////////////////////
              If Me.DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
                Me.DtaConsulta.RecordSource = "SELECT año, CodTipoNomina, COUNT(mes) AS Cont, mes AS Mes From Fecha_Planilla GROUP BY año, CodTipoNomina, mes HAVING (año = " & Anno & ") AND (CodTipoNomina = '" & CodTipoNomina & "')  AND (mes = '" & Format(Mes, "0#") & "')" 'AND (COUNT(mes) = 3)
                Me.DtaConsulta.Refresh
                Do While Not Me.DtaConsulta.Recordset.EOF
                   If Me.DtaConsulta.Recordset("Cont") = 3 Then
                     SalTemp = (SalTemp / 42) * 30
                   Else
                     SalTemp = (SalTemp / 28) * 30
                 
                   End If
                 
                  
                  Me.DtaConsulta.Recordset.MoveNext
                Loop
              End If
'
'                If Mes = 5 Then
'                   SalTemp = (SalTemp / 42) * 30
'                 ElseIf Mes = 10 Then
'                   SalTemp = (SalTemp / 42) * 30
'                 Else
'                   SalTemp = (SalTemp / 28) * 30
'
'                 End If

               
               CantRegistros = CantRegistros + 1
               DtaNominas.Recordset.MoveNext
              

              Loop
              
              If HorasExtra <= 0 Then
                CalculaHE = False
              End If
              
              
              '///////////////////////////////////////////////////////////////////////////////////
              '//////////////////////////SI EL SALARIO ES BASICO NO SE TOMA EN CUENTA PROMEDIO/////////////
              '///////////////////////////////////////////////////////////////////////////////////
              MDIPrimero.DtaConsulta.RecordSource = "SELECT  Historico.*, Empleado.SueldoActualBasico FROM Historico INNER JOIN Empleado ON Historico.Codempleado = Empleado.CodEmpleado  Where (Historico.CodEmpleado = " & CodEmpleado & ")"
              MDIPrimero.DtaConsulta.Refresh
              If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
                If MDIPrimero.DtaConsulta.Recordset("SueldoActualBasico") = True Then
                 If Not IsNull(MDIPrimero.DtaConsulta.Recordset("SueldoActual")) Then
                    SalarioBasico = MDIPrimero.DtaConsulta.Recordset("SueldoActual")
                 Else
                    SalarioBasico = 0
                 End If
                    SalTemp = SalarioBasico + Ingresos
                End If
              End If
              
                '///////////////////////Busco si el Empleado tiene Subsidio para Sumarlo////////////////////
  
                 'Me.DtaConsulta.RecordSource = "SELECT NomSubsidio.NumNomina, NomSubsidio.FechaPago, DetalleNomSubsidio.CodEmpleado, DetalleNomSubsidio.Subsidio FROM NomSubsidio INNER JOIN DetalleNomSubsidio ON NomSubsidio.NumNomina = DetalleNomSubsidio.NumNominaSubsidio WHERE (((NomSubsidio.FechaPago) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((DetalleNomSubsidio.CodEmpleado)='" & CodEmpleado & "'))"
                 Me.DtaConsulta.RecordSource = "SELECT NomSubsidio.NumNomina, DetalleNomSubsidio.CodEmpleado, DetalleNomSubsidio.Subsidio, Empleado.SumarSubsidio, Year([NomSubsidio].[FechaPago]) AS Anno, Month([NomSubsidio].[FechaPago]) AS Mes FROM NomSubsidio INNER JOIN (Empleado INNER JOIN DetalleNomSubsidio ON Empleado.CodEmpleado = DetalleNomSubsidio.CodEmpleado) ON NomSubsidio.NumNomina = DetalleNomSubsidio.NumNominaSubsidio Where (((DetalleNomSubsidio.CodEmpleado) = '" & CodEmpleado & "') And ((Empleado.SumarSubsidio) = 1) And ((Year([NomSubsidio].[FechaPago])) = '" & Anno & "') And ((Month([NomSubsidio].[FechaPago])) = '" & Mes & "')) ORDER BY NomSubsidio.NumNomina"
                 Me.DtaConsulta.Refresh
                 MontoSubsidio = 0
                 Do While Not Me.DtaConsulta.Recordset.EOF
                  MontoSubsidio = MontoSubsidio + Me.DtaConsulta.Recordset("Subsidio")
                  Me.DtaConsulta.Recordset.MoveNext
                 Loop
              
              SalTemp = SalTemp + MontoSubsidio
                 Fecha1 = Year(Me.TxtFINI13.Value) & "-" & Month(Me.TxtFINI13.Value) & "-" & Day(Me.TxtFINI13.Value)
                 Fecha2 = Year(Me.TxtFFIN13.Value) & "-" & Month(Me.TxtFFIN13.Value) & "-" & Day(Me.TxtFFIN13.Value)
                 
               
                 
               If SalTemp <> 0 Then
                 Me.AdoHistorialSalarial.RecordSource = "SELECT NumNomina, Tipo, CodEmpleado, FechaIni, FechaFin, Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, Agosto, Septiembre, Octubre, Noviembre, Diciembre From HistorialSalarioMes WHERE     (CodEmpleado = '" & CodEmpleado & "') AND (FechaIni = CONVERT(DATETIME, '" & Fecha1 & "', 102)) AND (FechaFin = CONVERT(DATETIME, '" & Fecha2 & "',102))"
                 Me.AdoHistorialSalarial.Refresh
                  If Me.AdoHistorialSalarial.Recordset.EOF Then
                        Me.AdoHistorialSalarial.Recordset.AddNew
                        Me.AdoHistorialSalarial.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
                        Me.AdoHistorialSalarial.Recordset("FechaIni") = CDate(Me.TxtFINI13.Value)
                        Me.AdoHistorialSalarial.Recordset("FechaFin") = CDate(Me.TxtFFIN13.Value)
                        Me.AdoHistorialSalarial.Recordset("NumNomina") = val(Me.TxtNumNom13.Text)
                        Me.AdoHistorialSalarial.Recordset("Tipo") = "Aguinaldo"
                        Select Case Mes
                          Case 1
                            Me.AdoHistorialSalarial.Recordset("Enero") = SalTemp
                          Case 2
                            Me.AdoHistorialSalarial.Recordset("Febrero") = SalTemp
                          Case 3
                            Me.AdoHistorialSalarial.Recordset("Marzo") = SalTemp
                          Case 4
                            Me.AdoHistorialSalarial.Recordset("Abril") = SalTemp
                          Case 5
                            Me.AdoHistorialSalarial.Recordset("Mayo") = SalTemp
                          Case 6
                            Me.AdoHistorialSalarial.Recordset("Junio") = SalTemp
                          Case 7
                            Me.AdoHistorialSalarial.Recordset("Julio") = SalTemp
                          Case 8
                            Me.AdoHistorialSalarial.Recordset("Agosto") = SalTemp
                          Case 9
                            Me.AdoHistorialSalarial.Recordset("Septiembre") = SalTemp
                          Case 10
                            Me.AdoHistorialSalarial.Recordset("Octubre") = SalTemp
                          Case 11
                            Me.AdoHistorialSalarial.Recordset("Noviembre") = SalTemp
                          Case 12
                            Me.AdoHistorialSalarial.Recordset("Diciembre") = SalTemp
                         End Select
                        Me.AdoHistorialSalarial.Recordset.Update
                  Else
                         Me.AdoHistorialSalarial.Recordset("FechaIni") = CDate(Me.TxtFINI13.Value)
                         Me.AdoHistorialSalarial.Recordset("FechaFin") = CDate(Me.TxtFFIN13.Value)
                         Me.AdoHistorialSalarial.Recordset("NumNomina") = val(Me.TxtNumNom13.Text)
                         Me.AdoHistorialSalarial.Recordset("Tipo") = "Aguinaldo"
                        Select Case Mes
                          Case 1
                            Me.AdoHistorialSalarial.Recordset("Enero") = SalTemp
                          Case 2
                            Me.AdoHistorialSalarial.Recordset("Febrero") = SalTemp
                          Case 3
                            Me.AdoHistorialSalarial.Recordset("Marzo") = SalTemp
                          Case 4
                            Me.AdoHistorialSalarial.Recordset("Abril") = SalTemp
                          Case 5
                            Me.AdoHistorialSalarial.Recordset("Mayo") = SalTemp
                          Case 6
                            Me.AdoHistorialSalarial.Recordset("Junio") = SalTemp
                          Case 7
                            Me.AdoHistorialSalarial.Recordset("Julio") = SalTemp
                          Case 8
                            Me.AdoHistorialSalarial.Recordset("Agosto") = SalTemp
                          Case 9
                            Me.AdoHistorialSalarial.Recordset("Septiembre") = SalTemp
                          Case 10
                            Me.AdoHistorialSalarial.Recordset("Octubre") = SalTemp
                          Case 11
                            Me.AdoHistorialSalarial.Recordset("Noviembre") = SalTemp
                          Case 12
                            Me.AdoHistorialSalarial.Recordset("Diciembre") = SalTemp
                        End Select
                        Me.AdoHistorialSalarial.Recordset.Update
       
                  End If
            
            Else
'                 Me.AdoHistorialSalarial.RecordSource = "SELECT NumNomina, CodEmpleado, FechaIni, FechaFin, Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, Agosto, Septiembre, Octubre, Noviembre, Diciembre From HistorialSalarioMes WHERE  (CodEmpleado = '" & CodEmpleado & "')"
'                 Me.AdoHistorialSalarial.Refresh
'                  If Not Me.AdoHistorialSalarial.Recordset.EOF Then
'
'                        Select Case Mes
'                          Case 1
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Enero")
'                          Case 2
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Febrero")
'                          Case 3
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Marzo")
'                          Case 4
'                           SalTemp = Me.AdoHistorialSalarial.Recordset("Abril")
'                          Case 5
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Mayo")
'                          Case 6
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Junio")
'                          Case 7
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Julio")
'                          Case 8
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Agosto")
'                          Case 9
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Septiembre")
'                          Case 10
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Octubre")
'                          Case 11
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Noviembre")
'                          Case 12
'                            SalTemp = Me.AdoHistorialSalarial.Recordset("Diciembre")
'                         End Select
'
'                 End If
            
            End If
            
        If SalMayor < SalTemp Then
           SalMayor = SalTemp
        End If

        If SalMayorCHE < SalTempHE Then
            SalMayorCHE = SalTempHE
        End If
        
        
        
  
       Next
   Else
   
      '///////////////////////////////////////////////////////////////////////////////////////////////////////
      '/////////////////////////////////77SI EL EMPLEADO ES SALARIO FIJO ///////////////////////////////
      '///////////////////////////////////////////////////////////////////////////////////////////////
   
      MDIPrimero.DtaConsulta.RecordSource = "SELECT  Historico.*, Empleado.SueldoActualBasico,Empleado.SalPorcentaje FROM Historico INNER JOIN Empleado ON Historico.Codempleado = Empleado.CodEmpleado  Where (Historico.CodEmpleado = " & CodEmpleado & ")"
      MDIPrimero.DtaConsulta.Refresh
      If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
           SalPorciento = MDIPrimero.DtaConsulta.Recordset("SalPorcentaje")
           SalPorciento = SalPorciento / 100
      Else
           SalPorciento = 0
      End If

     SalMayor = DtaEmpleados.Recordset("SueldoPeriodo") * (1 + SalPorciento)
     
        'Else
       '  SalMayor = DtaEmpleados.Recordset("SueldoPeriodo")
       'End If
    
    'dependiendo del tipo de pago se hace el calculo del salario básico
     If DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
           SalMayor = SalMayor * 6
     ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
           SalMayor = SalMayor * 12
     ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
         'If CantRegistros > 0 Then
           SalMayor = SalMayor * 2
          'End If
          
     ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
           SalMayor = DiasMes * (DtaEmpleados.Recordset("SueldoPeriodo") / 14)
     
     End If
     
       
'       If CodEmpleado = "10709" Or CodEmpleado = "10796" Then
'             SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.OtrosIngresos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo] + [DetalleNomina].[Antiguedad] + DetalleNomina.Incentivos + DetalleNomina.OtrosIngresos  AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where (Nomina.Mes =" & Mes & ") And (Nomina.Ano =" & Anno & ") and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
'       Else
'              SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.OtrosIngresos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo] + [DetalleNomina].[Antiguedad] + DetalleNomina.Incentivos + DetalleNomina.Comisiones + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos   AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where (Nomina.Mes =" & Mes & ") And (Nomina.Ano =" & Anno & ") and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
'       End If
'            DtaNominas.RecordSource = SqlNominas
'            DtaNominas.Refresh
'               SalTemp = 0
'               Ingresos = 0
'               HorasExtra = 0
'               SalTempHE = 0
'              Do While Not DtaNominas.Recordset.EOF
'
'
'                   HorasExtra = HorasExtra + DtaNominas.Recordset("HorasExtras")
'                   SalTemp = SalTemp + DtaNominas.Recordset("Total")
'                   SalTempHE = SalTempHE + DtaNominas.Recordset("Total") + DtaNominas.Recordset("HorasExtras")
'                   SalTemp = SalTempHE  '''PARA QUE CALCULE CON HORAS EXTRAS
'                   Ingresos = Ingresos + (DtaNominas.Recordset("Destajo") + DtaNominas.Recordset("OtrosIngresos") + DtaNominas.Recordset("Incentivos"))
'
'
'                  '/////////////////////////////////////////////////////////////////////////////////////
'                  '//////////////////////////////RECORRO LA TABLA DE PERIODO PARA SELECCIONAR LOS MESES /
'                  '///////////////////////////////////////////////////////////////////////////////////////
'                  If Me.DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
'                    Me.DtaConsulta.RecordSource = "SELECT año, CodTipoNomina, COUNT(mes) AS Cont, mes AS Mes From Fecha_Planilla GROUP BY año, CodTipoNomina, mes HAVING (año = " & Anno & ") AND (CodTipoNomina = '" & CodTipoNomina & "')  AND (mes = '" & Format(Mes, "0#") & "')" 'AND (COUNT(mes) = 3)
'                    Me.DtaConsulta.Refresh
'                    Do While Not Me.DtaConsulta.Recordset.EOF
'                       If Me.DtaConsulta.Recordset("Cont") = 3 Then
'                         SalTemp = (SalTemp / 42) * 30
'                       Else
'                         SalTemp = (SalTemp / 28) * 30
'
'                       End If
'
'
'                      Me.DtaConsulta.Recordset.MoveNext
'                    Loop
'                  End If
'
'               CantRegistros = CantRegistros + 1
'               DtaNominas.Recordset.MoveNext
'
'
'              Loop

     
     
              '///////////////////////////////////////////////////////////////////////////////////
              '//////////////////////////SI EL SALARIO ES BASICO NO SE TOMA EN CUENTA PROMEDIO/////////////
              '///////////////////////////////////////////////////////////////////////////////////
              MDIPrimero.DtaConsulta.RecordSource = "SELECT  Historico.*, Empleado.SueldoActualBasico FROM Historico INNER JOIN Empleado ON Historico.Codempleado = Empleado.CodEmpleado  Where (Historico.CodEmpleado = " & CodEmpleado & ")"
              MDIPrimero.DtaConsulta.Refresh
              If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
                If MDIPrimero.DtaConsulta.Recordset("SueldoActualBasico") = True Then
                 If Not IsNull(MDIPrimero.DtaConsulta.Recordset("SueldoActual")) Then
                    SalarioBasico = MDIPrimero.DtaConsulta.Recordset("SueldoActual")
                 Else
                    SalarioBasico = 0
                 End If
                    SalMayor = SalarioBasico '+ Ingresos
                End If
              End If
    
    
 End If
        
    If CalculaHE = True Then
      If SalMayorCHE > SalMayor Then
        SalMayor = SalMayorCHE
      End If
    End If
    SalMayorCHE = 0
    
    
    '//// Calculo cuanto voy a deducir en base a pension alimenticia //////
      
    If PorcentajePension > 0 Then
        PorcentajePension = SalMayor * (PorcentajePension / 100)
    End If
       
DtaNom13Mes.Refresh
Do While Not DtaNom13Mes.Recordset.EOF
    If DtaNom13Mes.Recordset("NumNom13Mes") = val(TxtNumNom13.Text) And DtaNom13Mes.Recordset("Activa") = True Then
       'DtaNom13Mes.Recordset.Edit
       DtaNom13Mes.Recordset("montopagado") = DtaNom13Mes.Recordset("montopagado") + (SalMayor - MontoPension - PorcentajePension)
       DtaNom13Mes.Recordset.Update
   Exit Do
    End If
DtaNom13Mes.Recordset.MoveNext
Loop

       Me.DtaDetalleNom13Mes.RecordSource = "SELECT Id, NumNom13Mes, MontoPension, PorcentajePension, CodEmpleado, SalarioMensual, SalarioAPagar, DiasAPagar, Adelanto13vo, MontoSuspension, DiasSuspension, TotalDeducciones From DetalleNom13Mes"
       Me.DtaDetalleNom13Mes.Refresh
       If Me.DtaDetalleNom13Mes.Recordset.EOF Then
       Id = 1
       Else
       Me.DtaDetalleNom13Mes.Recordset.MoveLast
       Id = Me.DtaDetalleNom13Mes.Recordset("id") + 1
       End If
       
    LlenarHistorialAguinaldo (DtaEmpleados.Recordset("CodEmpleado"))
       
    If (SalMayor - MontoPension - PorcentajePension) <> 0 Then
       
        DtaDetalleNom13Mes.Recordset.AddNew
        Me.DtaDetalleNom13Mes.Recordset("id") = Id
        DtaDetalleNom13Mes.Recordset("Adelanto13vo") = Adelanto13vo
        DtaDetalleNom13Mes.Recordset("NumNom13Mes") = val(TxtNumNom13.Text)
        DtaDetalleNom13Mes.Recordset("MontoPension") = Format(MontoPension, "##,##0.00")
        DtaDetalleNom13Mes.Recordset("PorcentajePension") = Format(PorcentajePension, "##,##0.00")
        
        If DtaEmpleados.Recordset("CodEmpleado") = "10706" Then
        Dim asd As Integer
            asd = 0
        End If
        
        DtaDetalleNom13Mes.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
        DtaDetalleNom13Mes.Recordset("SalarioMensual") = (SalMayor - MontoPension - PorcentajePension)
        If (Not Dias = 0) And Dias < 30 Then
         DtaDetalleNom13Mes.Recordset("SalarioAPagar") = (((SalMayor - MontoPension - PorcentajePension) * Format(Dias, "##,##0.00")) / DiasMes)
        ElseIf Dias >= 30 Then
         DtaDetalleNom13Mes.Recordset("SalarioAPagar") = (SalMayor - MontoPension - PorcentajePension)
        Else
         DtaDetalleNom13Mes.Recordset("SalarioAPagar") = 0
        End If
        
        
        Dim CantidadSuspension As Double
        Dim DiasSuspension As Double
        AdoConsulta.RecordSource = "Select sum(DiasDisfrutar) as DiasDisfrutar from solicitudvacaciones Where TipoSolicitud = 'Suspension' and CodigoEmpleado = '" & CodEmpleado1 & "' and FechaInicio >= '" & Format(Me.TxtFINI13.Value, "yyMMdd") & " 00:00' and FechaFin <= '" & Format(Me.TxtFFIN13.Value, "yyMMdd") & " 23:59'"
        AdoConsulta.Refresh
        
        If Not AdoConsulta.Recordset.EOF Then
            If Not IsNull(AdoConsulta.Recordset("Diasdisfrutar")) Then
               CantidadSuspension = Format(AdoConsulta.Recordset("DiasDisfrutar") * 0.08333333, "##,##0.00")
               DiasSuspension = AdoConsulta.Recordset("DiasDisfrutar")
             Else
               CantidadSuspension = 0
               DiasSuspension = 0
            End If

        End If
        
        DtaDetalleNom13Mes.Recordset("MontoSuspension") = (Format(DtaDetalleNom13Mes.Recordset("SalarioMensual"), "##,##0.00") / 30) * (CantidadSuspension)
        DtaDetalleNom13Mes.Recordset("DiasSuspension") = DiasSuspension
        DtaDetalleNom13Mes.Recordset("TotalDeducciones") = MontoPension + PorcentajePension
        DtaDetalleNom13Mes.Recordset("DiasAPagar") = Dias
        DtaDetalleNom13Mes.Recordset.Update
   End If
End If


.Value = i
i = i + 1

DtaEmpleados.Recordset.MoveNext
Loop
End With
NumFecha1 = CDate(Me.TxtFINI13)
NumFecha2 = CDate(Me.TxtFFIN13)
Sql13voMes = "SELECT Nom13Mes.NumNom13Mes, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, DetalleNom13Mes.SalarioAPagar-DetalleNom13Mes.Adelanto13vo AS MontoPagar FROM Nom13Mes INNER JOIN (Empleado INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNom13Mes & ")) ORDER BY Empleado.CodEmpleado1"

Dta13voMes.RecordSource = Sql13voMes
Dta13voMes.Refresh

Me.Dbgr13Mes.Columns(0).Visible = False
Me.Dbgr13Mes.Columns(0).Locked = True
Me.Dbgr13Mes.Columns(1).Locked = True
Me.Dbgr13Mes.Columns(2).Locked = True
Me.Dbgr13Mes.Columns(3).Locked = True
Me.Dbgr13Mes.Columns(4).Locked = True
Me.Dbgr13Mes.Columns(5).Locked = True
Me.Dbgr13Mes.Columns(6).Locked = True
Me.Dbgr13Mes.Columns(7).Locked = True
Me.Dbgr13Mes.Columns(9).Locked = True
Me.Dbgr13Mes.Columns(6).NumberFormat = "##,##0.00"
Me.Dbgr13Mes.Columns(7).NumberFormat = "##,##0.00"
Me.Dbgr13Mes.Columns(8).NumberFormat = "##,##0.00"
Me.Dbgr13Mes.Columns(9).NumberFormat = "##,##0.00"
'CmdPr13mes.Enabled = True
Me.CmdPrnNomina.Enabled = True
Me.CmdExportaBAC.Enabled = True
Me.CmdprNomina.Enabled = True
Me.CmdExporta2.Enabled = True
Me.CmdDenominacion.Enabled = True
MousePointer = 1

End Sub

Private Sub CmdCalcularVacaciones_Click()
Dim SalarioBasico As Double, DiasDescuento As Double
Dim SqlEmpleados As String, MontoSubsidio As Double
Dim CodTipoNomina As String, TotalSubsidio As Double
Dim SqlNominas As String, NumVaca As Integer
Dim SalMayor As Double, TarifaHoraria As Double
Dim SalTemp As Double, Dias As Double
Dim CodEmpleado As String, AdelantoVaca As Double
Dim Edicion As Boolean, FechaContratos As String, MesContrato As Double, MesVaca As Double
Dim Anno As Integer, SqlSalarios As String
Dim Mes As Integer, Fecha As Date, MontoInss As Double
Dim CantEmpleados As Long
Dim i As Integer, CantMeses As Integer, CantRegistros As Integer
Dim DiasMes As Double
Dim DiasSemana As Double, DiasPagar As Double
Dim FechaHoy As Date, FechaInicio As Date, FechaFin As Date, FechaInicioVaca As Date, FechaFinVaca As Date
Dim rsDB As New ADODB.Recordset
Dim rs As New ADODB.Recordset, Ejecutar As ADODB.Connection
Dim cnDB As New ADODB.Connection
Dim iMes As Integer, Monto As Double, DiasAcumulados As Double
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer, SalarioAcumulado As Double, IR As Double, IrAcumulado As Double, InssAcumulado As Double



Me.CmdExportar.Enabled = True

DtaControles.Refresh
DiasMes = DtaControles.Recordset("DiasMes")
DiasSemana = DtaControles.Recordset("DiasSemana")



FechaHoy = Format(Me.TxtFechaAplica.Value, "dd/mm/yyyy")

If FechaHoy < Me.TxtFINIVaca.Value Then
 MsgBox "La fecha Actual no Coincide con la Nomina", vbCritical, "Sistema de NOminas"
 Exit Sub
ElseIf FechaHoy > Me.TxtFFinVaca.Value Then
  FechaHoy = Me.TxtFFinVaca
End If

Anno = Year(FechaHoy)

Edicion = False

MousePointer = 11
If DBCNominas.Text = "Listado de Nominas" Then
   MsgBox "No ha seleccionado el tipo de nomina al cual le desea calcular las Vacaciones"
   MousePointer = 1
   DBCNominas.SetFocus
   Exit Sub
End If


If Not IsNumeric(TxtDiasDescuento.Text) Then
   MsgBox "Los dias de descuendo de Vacaciones son erróneos"
   MousePointer = 1
   TxtDiasDescuento.SetFocus
   Exit Sub
ElseIf val(TxtDiasDescuento.Text) > 15 Then
   MsgBox "Los dias de descuendo de Vacaciones no pueden ser mayor de 15"
   MousePointer = 1
   TxtDiasDescuento.SetFocus
   Exit Sub
End If

DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop

NumVaca = val(TxtNumNomVaca.Text)

If Me.ChkEliminar.Value = 1 Then
 Fecha1 = Format(Me.TxtFINIVaca.Value, "yyyy-mm-dd")
 Fecha2 = Format(Me.TxtFFinVaca.Value, "yyyy-mm-dd")
 Set Ejecutar = New ADODB.Connection
 Ejecutar.ConnectionString = Conexion
 Ejecutar.Open
 Ejecutar.Execute "DELETE FROM DetalleNomVaca WHERE (NumNomVaca =  " & NumVaca & ")"
 Ejecutar.Execute "DELETE FROM HistorialSalarioMes WHERE (FechaIni = CONVERT(DATETIME, '" & Fecha1 & "', 102)) AND (FechaFin = CONVERT(DATETIME, '" & Fecha2 & "', 102))"
End If


'pregunto si ya existe una nomina de vacaciones elaborada
Me.DtaNomVaca.RecordSource = "SELECT CodTipoNomina, NumNomVaca, FechaAplica, MontoPagado, FechaIni, FechaFin, Activa From NomVaca Where (Activa = 1) And (NumNomVaca = " & NumVaca & ")"
DtaNomVaca.Refresh
If Not DtaNomVaca.Recordset.EOF Then
    'If DtaNomVaca.Recordset("NumNomVaca") = Val(TxtNumNomVaca.Text) And DtaNomVaca.Recordset.Activa = True Then
       'DtaNomVaca.Recordset.Edit
       DtaNomVaca.Recordset("FechaAplica") = Me.TxtFechaAplica.Value
       DtaNomVaca.Recordset("montopagado") = 0
       DtaNomVaca.Recordset("Fechaini") = CDate(TxtFINIVaca.Value)
       DtaNomVaca.Recordset("Fechafin") = CDate(TxtFFinVaca.Value)
       DtaNomVaca.Recordset("CodTipoNomina") = CodTipoNomina
       DtaNomVaca.Recordset.Update
       Edicion = True
    
    'End If
End If

If Not Edicion Then 'es primera vez que se crearan las vacaciones

       DtaNomVaca.Recordset.AddNew
       DtaNomVaca.Recordset("NumNomVaca") = val(TxtNumNomVaca.Text)
       DtaNomVaca.Recordset("FechaAplica") = Format(Now, "DD/MM/YYYY")
       DtaNomVaca.Recordset("montopagado") = 0
       DtaNomVaca.Recordset("Fechaini") = CDate(TxtFINIVaca.Value)
       DtaNomVaca.Recordset("Fechafin") = CDate(TxtFFinVaca.Value)
       DtaNomVaca.Recordset("CodTipoNomina") = CodTipoNomina
       DtaNomVaca.Recordset("Activa") = 1
       DtaNomVaca.Recordset.Update
       
Else 'se borran los movimientos anteriores

     'DtaDetalleNomVaca.Refresh
     'Do While Not DtaDetalleNomVaca.Recordset.EOF
     'If DtaDetalleNomVaca.Recordset("NumNomVaca") = Val(TxtNumNomVaca) Then
      '  DtaDetalleNomVaca.Recordset.Delete
     'End If
     
     'DtaDetalleNomVaca.Recordset.MoveNext
     'Loop
     
     
End If

'hago el sql de los empleados que pertenezcan solo a la nomina seleccionada

'SqlEmpleados = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.CodTipoNomina From Empleado WHERE Empleado.CodTipoNomina=  '" & CodTipoNomina & "'"
'SqlEmpleados = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, Empleado.ExentoIr, Empleado.OtrosIngresos,  Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, Empleado.Ausente, Empleado.SalarioFijo From Empleado WHERE Empleado.CodTipoNomina= '" & CodTipoNomina & "' AND Empleado.Activo=1 AND Empleado.Ausente=0"

'SqlEmpleados = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, " & _
'               "Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, " & _
'               "Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, " & _
'               "Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, " & _
'               "Empleado.ExentoIr, Empleado.OtrosIngresos, Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, " & _
'               "Empleado.Ausente, Empleado.SalarioFijo, Empleado.CodEmpleado1, Historico.FechaContrato,Historico.FechaContratoVac " & _
'               "FROM  Empleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado " & _
'               "WHERE  (Empleado.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (Empleado.Ausente = 0)"


                  
                  '/////////////////////////////////////////////////////////////////////////////////////////////////////////
                  '///////////////////////////////BUSCA LA FECHA DE HACE 6 MESES////////////////////////////////////////////
                  '////////////////////////////////////////////////////////////////////////////////////////////////////////
                  
                  FechaInicioVaca = DateSerial(Year(Me.TxtFFinVaca.Value), Month(Me.TxtFFinVaca.Value) - i, 1)
                  For i = 1 To 6
                   If i = 5 Then
                     FechaFinVaca = DateSerial(Year(Me.TxtFFinVaca.Value), Month(Me.TxtFFinVaca.Value) - i, 0)
                   End If

                  Next
                  
                  
               
                  
              

'                  FechaFinVaca = Me.dtpFPInicio.Value - 1
'                  FechaInicioVaca = FechaVacaciones(FechaFinVaca)
'                  mes =
                  



'SqlEmpleados = "SELECT * FROM  Historico INNER JOIN Empleado ON Historico.Codempleado = Empleado.CodEmpleado " & _
               "WHERE (Empleado.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Historico.Codempleado "
               
 SqlEmpleados = "SELECT  * FROM Historico INNER JOIN Empleado ON Historico.Codempleado = Empleado.CodEmpleado WHERE (Empleado.Activo = 1) AND (Empleado.CodTipoNomina = '" & CodTipoNomina & "') AND (Historico.FechaContratoVac BETWEEN CONVERT(DATETIME, '" & Format(FechaInicioVaca, "yyyy-mm-dd") & "',102) AND CONVERT(DATETIME, '" & Format(FechaFinVaca, "yyyy-mm-dd") & "', 102)) ORDER BY Historico.Codempleado"

DtaEmpleados.RecordSource = SqlEmpleados
DtaEmpleados.Refresh

If Not Me.DtaEmpleados.Recordset.EOF Then
 DtaEmpleados.Recordset.MoveLast
 CantEmpleados = DtaEmpleados.Recordset.RecordCount
End If

PBVacaciones.Value = 0

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////BUSCO LAS FECHAS DE LAS VACACIONES EN LAS NOMINAS CALCULADAS////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


'                        FechaFin = Me.dtpFPInicio.Value - 1
'
'                        SQlSalarios = "SELECT DISTINCT TOP 100 PERCENT SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos, SUM (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO, SUM(DetalleNomina.Comisiones) As Comisiones FROM  DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY Nomina.Mes, Nomina.Ano  " & _
'                                      "HAVING  (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.Comisiones) <> 0) AND (MAX(Nomina.FechaNomina) <= CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) ORDER BY Nomina.Ano, Nomina.Mes"
'
'                        Me.DtaConsulta.RecordSource = SQlSalarios
'                        Me.DtaConsulta.Refresh
'                        If Not Me.DtaConsulta.Recordset.EOF Then
'                         Me.DtaConsulta.Recordset.MoveLast
'                        Else
''                         FechaFin = Format(Now, "dd/mm/yyyy")
'                         FechaInicio = Format(Now, "dd/mm/yyyy")
'                        End If
'                        I = 0
'                        Do While Not Me.DtaConsulta.Recordset.BOF
'                          If I = 1 Then
'                            FechaInicio = Me.DtaConsulta.Recordset("FechaInicio")
'
'                          ElseIf I = 5 Then
'                            FechaInicio = Me.DtaConsulta.Recordset("FechaInicio")
'                            Exit Do
'                          ElseIf I = 0 Then
'                            FechaInicio = Me.DtaConsulta.Recordset("FechaInicio")
''                            FechaFin = Me.DtaConsulta.Recordset("FechaFin")
'                          Else
'                            FechaInicio = Me.DtaConsulta.Recordset("FechaInicio")
'                          End If
'                          I = I + 1
'
'                          Me.DtaConsulta.Recordset.MovePrevious
'                        Loop
                        
                        
                  FechaFin = Me.dtpFPInicio.Value - 1
                  FechaInicio = FechaVacaciones(FechaFin)
'                   FechaFin = FechaFinVaca
'                   FechaInicio = FechaInicioVaca
                  
                  


With PBVacaciones
         .Min = 0
         .Max = CantEmpleados
         .Value = 0
         i = 1
        DtaEmpleados.Refresh
        
        cnDB.ConnectionString = Conexion
        cnDB.Open
        
                  
                  
        'recorro la BD empleados y a cada uno le busco su salario mayor (sies destajo) si no solo extraigo su salario
        Do While Not DtaEmpleados.Recordset.EOF
        
                              .Value = i
                              'Me.xp_canvas1.Caption = "Procesando:  " & i & " de " & CantEmpleados & " Empleados "
                              Me.Caption = "Procesando:  " & i & " de " & CantEmpleados & " Empleados "
                              Me.LblTotal.Caption = "Procesando:  " & i & " de " & CantEmpleados & " Empleados "
                              DoEvents
                             CodEmpleado = DtaEmpleados.Recordset("CodEmpleado")
                             DiasAcumulados = 0
                             SalarioAcumulado = 0
                             
                           
                                                     

                             
                          '/////////////////ESTE ES EL SALARIO BASICO, SIN PRODUCCION////////////////
                             TarifaHoraria = DtaEmpleados.Recordset("TarifaHoraria")
                             SalarioBasico = Format(DiasMes * 8 * TarifaHoraria, "####0.00")

        
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '//////////////////VERIFICO SI TIENE LOS 6 MESES PARA INCLUIRLO EN LAS VACACIONES///////////////////////////////
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////
        
 
         
           If Not IsNull(Me.DtaEmpleados.Recordset("FechaContratoVac")) Then
'            FechaContratos = Format(Day(Me.DtaEmpleados.Recordset("FechaContratoVac")) & "/" & Month(Me.DtaEmpleados.Recordset("FechaContratoVac")) & "/" & Year(Me.TxtFINIVaca.Value), "dd/mm/yyyy")
            FechaContratos = Me.DtaEmpleados.Recordset("FechaContratoVac")
            MesContrato = Month(Me.DtaEmpleados.Recordset("FechaContratoVac"))
           Else
            MsgBox "El Empleado: " & Me.DtaEmpleados.Recordset("CodEmpleado1") & " tiene la Fecha Contrato Vac nulo"
            FechaContratos = Format(Now, "dd/mm/yyyy")
            MesContrato = Month(Format(Now, "dd/mm/yyyy"))
            
           End If
           MesVaca = Month(Me.TxtFFinVaca.Value) + 1
           
           Meses = ((CDate(Me.TxtFFinVaca.Value) - CDate(FechaContratos)) + 1) / DiasMes
           Meses = ((CDate(Me.dtpFPFinal.Value) - CDate(FechaContratos)) + 1) / DiasMes
            
           Dias = (CDate(Me.TxtFFinVaca.Value) - CDate(FechaContratos)) + 1
           If Dias >= 182 And Dias <= 215 Then
'           If Meses >= 6 And Meses < 7 Then


                        




              
                         '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                         '////////////////////////BUSCOS LOS PERIODOS DE LAS NOMINAS DEL EMPLEADO//////////////////////////////////////////////////
                         '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        
                          rsDB.Open "SELECT TipoNomina.Nomina, Fecha_Planilla.Periodo, Fecha_Planilla.año, Fecha_Planilla.mes, Fecha_Planilla.Inicio, Fecha_Planilla.Final, " & _
                                  "Fecha_Planilla.Actual , Fecha_Planilla.NumNomina, Fecha_Planilla.Calculada " & _
                                  "FROM Fecha_Planilla INNER JOIN TipoNomina ON Fecha_Planilla.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                                  "WHERE TipoNomina.Nomina LIKE '" & Me.DBCNominas.Text & "' AND Fecha_Planilla.Inicio >= CONVERT(DATETIME, '" & Format(FechaInicio, "yyyy-mm-dd") & " 00:00:00', 102) AND Fecha_Planilla.Final <= CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & " 00:00:00', 102) ORDER BY Fecha_Planilla.Inicio ASC", cnDB
                        
                           
                        Me.DtaConsulta.RecordSource = "SELECT TipoNomina.Nomina, Fecha_Planilla.Periodo, Fecha_Planilla.año, Fecha_Planilla.mes, Fecha_Planilla.Inicio, Fecha_Planilla.Final, " & _
                                  "Fecha_Planilla.Actual , Fecha_Planilla.NumNomina, Fecha_Planilla.Calculada " & _
                                  "FROM Fecha_Planilla INNER JOIN TipoNomina ON Fecha_Planilla.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                                  "WHERE TipoNomina.Nomina LIKE '" & Me.DBCNominas.Text & "' AND Fecha_Planilla.Inicio >= CONVERT(DATETIME, '" & Format(FechaInicio, "yyyy-mm-dd") & " 00:00:00', 102) AND Fecha_Planilla.Final <= CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & " 00:00:00', 102) ORDER BY Fecha_Planilla.Inicio ASC "
                        

                             CodEmpleado = DtaEmpleados.Recordset("CodEmpleado")
                            'tengo que hacer un SQL de Solo los que esten en el rango de fechas
                            'solo se veran las nominas de cada mes
                            'se deben de hacer ciclos por cada mes, seis ciclos por los seis meses.
                         SalMayor = 0
                         CantMeses = 0
                         CantRegistros = 0
                         
                        
                          
                         
                         
                     
                        
                          DtaHistorico.RecordSource = "SELECT Historico.Codempleado, Historico.FechaBaja, Historico.FechaContrato From Historico Where (((Historico.CodEmpleado) = '" & CodEmpleado & "'))"
                          DtaHistorico.Refresh
                          NumFecha1 = CDate(Me.TxtFFinVaca)
                          NumFecha2 = CDate(Me.TxtFFinVaca)
                         
                         
                         
                          '////////Verifico cuantos dias tiene de trabajar///////////////////
                          If Not Me.DtaHistorico.Recordset.EOF Then
                           If Not IsNull(DtaHistorico.Recordset("FechaContrato")) Then
                             FechaContrato = DtaHistorico.Recordset("FechaContrato")
                             NumFecha2 = FechaContrato
                        '     NumFecha1 = Me.TxtFFinVaca.Value
                              NumFecha1 = Me.dtpFPFinal.Value
                             annos = ((CDbl(NumFecha1) - CDbl(NumFecha2)) + 1) / 365
                             Meses = ((CDbl(NumFecha1) - CDbl(NumFecha2)) + 1) / DiasMes
                                 
                             Dias = (CDate(Me.TxtFFinVaca.Value) - CDate(FechaContrato)) + 1
                        '     Dias = (CDate(Me.dtpFPFinal.Value) - CDate(FechaContrato)) + 1
                                 
                             If Dias < 0 Then
                               Dias = 0
                             ElseIf Dias > 182 Then
'                               Dias = (CDate(Me.TxtFFinVaca.Value) - CDate(Me.TxtFINIVaca.Value)) + 1
'                                Dias = (CDate(Me.TxtFFinVaca.Value) - CDate(FechaContratos)) + 1
                                Dias = 182.5
                             End If
                            End If
                           End If
                           
                           
                         
                         
                         If Dias > CInt(Me.txtAntiguedad.Text) Then
                         
                             If DtaEmpleados.Recordset("SalarioFijo") = "N" Then
                                 Año = Year(Me.TxtFFinVaca.Value)
                         '///////////Si el Salario es Variable Busco el Salario Mayor/////////
                              
                                '/////////////CAlculo las vacaciones Enero - Junio ////////////////////////////
                                SalMayor = 0
                                
                                
                                '/////////////////////////////////////////BUSCO LOS SALARIOS EN LOS PERIODOS ENCONTRADOS///////////////////////
                                
                                 Do While Not rsDB.EOF And Dias > CInt(Me.txtAntiguedad.Text)
                                 
                                 '30   '    For Mes = 1 To 6
                                 '+ DetalleNomina.HorasExtras
                                 
                                 
                                 If Me.ChkExtraVaca.Value = 1 Then
                                 
                                         SqlNominas = "SELECT DISTINCT " & _
                                                    "DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, " & _
                                                    "SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos, " & _
                                                    "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion+ DetalleNomina.HorasExtras) AS TotalIngresos, " & _
                                                    "MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO " & _
                                                    "FROM  DetalleNomina INNER JOIN " & _
                                                    "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina " & _
                                                    "GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                                                    "HAVING      (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = " & CodEmpleado & ") AND (Nomina.Mes = " & CInt(rsDB.Fields("mes")) & ") AND " & _
                                                    "(Nomina.Ano = " & rsDB.Fields("año") & ") ORDER BY Nomina.Ano, Nomina.Mes "
                                       DtaNominas.RecordSource = SqlNominas
                                       DtaNominas.Refresh
                                 
                                 
                                 Else
                              
                                       SqlNominas = "SELECT DISTINCT " & _
                                                    "DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, " & _
                                                    "SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos, " & _
                                                    "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) AS TotalIngresos, " & _
                                                    "MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO " & _
                                                    "FROM  DetalleNomina INNER JOIN " & _
                                                    "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina " & _
                                                    "GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                                                    "HAVING      (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = " & CodEmpleado & ") AND (Nomina.Mes = " & CInt(rsDB.Fields("mes")) & ") AND " & _
                                                    "(Nomina.Ano = " & rsDB.Fields("año") & ") ORDER BY Nomina.Ano, Nomina.Mes "
                                       DtaNominas.RecordSource = SqlNominas
                                       DtaNominas.Refresh
                                       
                                End If
                                    
'                                       If Month(Me.TxtFFinVaca.Value) <= 6 Then
'                                          Fecha = "01/01/" & Anno
'                                          NumFecha1 = Fecha
'                                          Fecha = "30/06/" & Anno
'                                          NumFecha2 = Fecha
'                                       Else
'                                          Fecha = "01/06/" & Anno
'                                          NumFecha1 = Fecha
'                                          Fecha = "30/11/" & Anno
'                                          NumFecha2 = Fecha
'
'                                       End If
                        
                                
                                
                                          SalTemp = 0
                                       If Not DtaNominas.Recordset.EOF Then
                                        CantMeses = CantMeses + 1
                                       End If
                                       
                                      Do While Not DtaNominas.Recordset.EOF
                                         If Not IsNull(DtaNominas.Recordset("TotalIngresos")) Then
                                             SalTemp = SalTemp + CDbl(DtaNominas.Recordset("Totalingresos"))
                                                    
                                         End If
                                       
                                         CantRegistros = CantRegistros + 1
                                         DtaNominas.Recordset.MoveNext
                                      Loop
                                      
                        '///////////////////////////////////////////////////////////////////////////////////
                        '//////////////////////////SI NO TIENE SALARIO EN UN MES NO SUMA PARA EL PROMEDIO/////////////
                        '///////////////////////////////////////////////////////////////////////////////////
                                      If SalTemp = 0 Then
                        '               CantMeses = CantMeses - 1
                                      End If
                                      
                                      
                                
                                      Fecha1 = Year(Me.TxtFINIVaca.Value) & "-" & Month(Me.TxtFINIVaca.Value) & "-" & Day(Me.TxtFINIVaca.Value)
                                      Fecha2 = Year(Me.TxtFFinVaca.Value) & "-" & Month(Me.TxtFFinVaca.Value) & "-" & Day(Me.TxtFFinVaca.Value)
                                      
                                      If SalTemp <> 0 Then
                                         Me.AdoHistorialSalarial.RecordSource = "SELECT NumNomina, Tipo, CodEmpleado, FechaIni, FechaFin, Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, Agosto, Septiembre, Octubre, Noviembre, Diciembre From HistorialSalarioMes WHERE     (CodEmpleado = '" & CodEmpleado & "') AND (FechaIni = CONVERT(DATETIME, '" & Fecha1 & "', 102)) AND (FechaFin = CONVERT(DATETIME, '" & Fecha2 & "',102))"
                                         Me.AdoHistorialSalarial.Refresh
                                          If Me.AdoHistorialSalarial.Recordset.EOF Then
                                                Me.AdoHistorialSalarial.Recordset.AddNew
                                                Me.AdoHistorialSalarial.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
                                                Me.AdoHistorialSalarial.Recordset("FechaIni") = CDate(Me.TxtFINIVaca.Value)
                                                Me.AdoHistorialSalarial.Recordset("FechaFin") = CDate(Me.TxtFFinVaca.Value)
                                                Me.AdoHistorialSalarial.Recordset("NumNomina") = val(Me.TxtNumNomVaca.Text)
                                                Me.AdoHistorialSalarial.Recordset("Tipo") = "Vacaciones"
                                                
                                                Select Case (CInt(rsDB.Fields("mes")))
                                                  Case 1
                                                    Me.AdoHistorialSalarial.Recordset("Enero") = SalTemp
                                                  Case 2
                                                    Me.AdoHistorialSalarial.Recordset("Febrero") = SalTemp
                                                  Case 3
                                                    Me.AdoHistorialSalarial.Recordset("Marzo") = SalTemp
                                                  Case 4
                                                    Me.AdoHistorialSalarial.Recordset("Abril") = SalTemp
                                                  Case 5
                                                    Me.AdoHistorialSalarial.Recordset("Mayo") = SalTemp
                                                  Case 6
                                                    Me.AdoHistorialSalarial.Recordset("Junio") = SalTemp
                                                  Case 7
                                                    Me.AdoHistorialSalarial.Recordset("Julio") = SalTemp
                                                  Case 8
                                                    Me.AdoHistorialSalarial.Recordset("Agosto") = SalTemp
                                                  Case 9
                                                    Me.AdoHistorialSalarial.Recordset("Septiembre") = SalTemp
                                                  Case 10
                                                    Me.AdoHistorialSalarial.Recordset("Octubre") = SalTemp
                                                  Case 11
                                                    Me.AdoHistorialSalarial.Recordset("Noviembre") = SalTemp
                                                  Case 12
                                                    Me.AdoHistorialSalarial.Recordset("Diciembre") = SalTemp
                                                 End Select
                                                Me.AdoHistorialSalarial.Recordset.Update
                                          Else
                                                 Me.AdoHistorialSalarial.Recordset("FechaIni") = CDate(Me.TxtFINIVaca.Value)
                                                 Me.AdoHistorialSalarial.Recordset("FechaFin") = CDate(Me.TxtFFinVaca.Value)
                                                 Me.AdoHistorialSalarial.Recordset("NumNomina") = val(Me.TxtNumNomVaca.Text)
                                                 Me.AdoHistorialSalarial.Recordset("Tipo") = "Vacaciones"
                                                Select Case (CInt(rsDB.Fields("mes")))
                                                  Case 1
                                                    Me.AdoHistorialSalarial.Recordset("Enero") = SalTemp
                                                  Case 2
                                                    Me.AdoHistorialSalarial.Recordset("Febrero") = SalTemp
                                                  Case 3
                                                    Me.AdoHistorialSalarial.Recordset("Marzo") = SalTemp
                                                  Case 4
                                                    Me.AdoHistorialSalarial.Recordset("Abril") = SalTemp
                                                  Case 5
                                                    Me.AdoHistorialSalarial.Recordset("Mayo") = SalTemp
                                                  Case 6
                                                    Me.AdoHistorialSalarial.Recordset("Junio") = SalTemp
                                                  Case 7
                                                    Me.AdoHistorialSalarial.Recordset("Julio") = SalTemp
                                                  Case 8
                                                    Me.AdoHistorialSalarial.Recordset("Agosto") = SalTemp
                                                  Case 9
                                                    Me.AdoHistorialSalarial.Recordset("Septiembre") = SalTemp
                                                  Case 10
                                                    Me.AdoHistorialSalarial.Recordset("Octubre") = SalTemp
                                                  Case 11
                                                    Me.AdoHistorialSalarial.Recordset("Noviembre") = SalTemp
                                                  Case 12
                                                    Me.AdoHistorialSalarial.Recordset("Diciembre") = SalTemp
                                                End Select
                                                Me.AdoHistorialSalarial.Recordset.Update
                               
                                          End If
                                    
                                    Else
                                    Me.AdoHistorialSalarial.RecordSource = "SELECT NumNomina, CodEmpleado, FechaIni, FechaFin, Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, Agosto, Septiembre, Octubre, Noviembre, Diciembre, Tipo From HistorialSalarioMes WHERE     (Tipo = N'Vacaciones') AND (CodEmpleado = '" & CodEmpleado & "') AND (FechaIni = CONVERT(DATETIME, '" & Fecha1 & "', 102)) AND (FechaFin = CONVERT(DATETIME, '" & Fecha2 & "', 102)) "
                                    
                                         Me.AdoHistorialSalarial.Refresh
                                          If Not Me.AdoHistorialSalarial.Recordset.EOF Then
                         
                                                Select Case (CInt(rsDB.Fields("mes")))
                                                  Case 1
                                                    SalTemp = Me.AdoHistorialSalarial.Recordset("Enero")
                                                  Case 2
                                                    SalTemp = Me.AdoHistorialSalarial.Recordset("Febrero")
                                                  Case 3
                                                    SalTemp = Me.AdoHistorialSalarial.Recordset("Marzo")
                                                  Case 4
                                                   SalTemp = Me.AdoHistorialSalarial.Recordset("Abril")
                                                  Case 5
                                                    SalTemp = Me.AdoHistorialSalarial.Recordset("Mayo")
                                                  Case 6
                                                    SalTemp = Me.AdoHistorialSalarial.Recordset("Junio")
                                                  Case 7
                                                    SalTemp = Me.AdoHistorialSalarial.Recordset("Julio")
                                                  Case 8
                                                    SalTemp = Me.AdoHistorialSalarial.Recordset("Agosto")
                                                  Case 9
                                                    SalTemp = Me.AdoHistorialSalarial.Recordset("Septiembre")
                                                  Case 10
                                                    SalTemp = Me.AdoHistorialSalarial.Recordset("Octubre")
                                                  Case 11
                                                    SalTemp = Me.AdoHistorialSalarial.Recordset("Noviembre")
                                                  Case 12
                                                    SalTemp = Me.AdoHistorialSalarial.Recordset("Diciembre")
                                                 End Select
                                    
                                         End If
                                    
                                    End If
                        
                                 
                                 
                                 
                                   SalMayor = SalTemp + SalMayor
                                   
                                 
                                   iMes = CInt(rsDB.Fields("mes"))
                                 
                                   Do While Not rsDB.EOF
                                 
                                     If iMes <> CInt(rsDB.Fields("mes")) Then
                                        Exit Do
                                     End If
                                   
                                     rsDB.MoveNext
                                   Loop
                                 
                                 
                                
                                 
                                Loop
                                       
                          '///////////////////////Busco si el Empleado tiene Subsidio para Sumarlo////////////////////
                          
                         
                                Me.DtaConsulta.RecordSource = "SELECT NomSubsidio.NumNomina, NomSubsidio.FechaPago, DetalleNomSubsidio.CodEmpleado, DetalleNomSubsidio.Subsidio FROM NomSubsidio INNER JOIN DetalleNomSubsidio ON NomSubsidio.NumNomina = DetalleNomSubsidio.NumNominaSubsidio WHERE (((NomSubsidio.FechaPago) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((DetalleNomSubsidio.CodEmpleado)='" & CodEmpleado & "'))"
                                Me.DtaConsulta.Refresh
                                MontoSubsidio = 0
                                
                                Do While Not Me.DtaConsulta.Recordset.EOF
                                   MontoSubsidio = MontoSubsidio + Me.DtaConsulta.Recordset("Subsidio")
                                   Me.DtaConsulta.Recordset.MoveNext
                                Loop
                               
                               
                              '/////////////Busco el Adelanto de Vacaciones Registrados//////////////
                        '////////////////////////////////////////////////////////////////////////
                                
                               
                                 
                                Me.DtaAdelanto.RecordSource = "SELECT  CodEmpleado, FechaAdelanto, MontoAdelanto, [Ref/Cheque], TipoAdelanto From Adelanto13vo WHERE     (FechaAdelanto BETWEEN CONVERT(DATETIME, '" & Format(Me.dtpFPInicio.Value, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.dtpFPFinal.Value, "yyyymmdd") & "', 102)) AND (TipoAdelanto = 'Vacaciones')AND (CodEmpleado = '" & CodEmpleado & "') "
                        '       Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo WHERE (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Adelanto13vo.FechaAdelanto) Between '" & Format(Me.dtpFPInicio.Value, "yyyy/mm/dd") & "' And '" & Format(Me.dtpFPFinal.Value, "yyyy/mm/dd") & "') AND ((Adelanto13vo.TipoAdelanto)='Vacaciones'))"
                        '        Me.DtaAdelanto.RecordSource = "SELECT CodEmpleado, FechaAdelanto, MontoAdelanto, [Ref/Cheque], TipoAdelanto From Adelanto13vo WHERE     (CodEmpleado = '" & CodEmpleado & "') AND (FechaAdelanto Between '" & Format(Me.dtpFPInicio.Value, "yyyymmdd") & "' And '" & Format(Me.dtpFPFinal.Value, "yyyymmdd") & "')) AND (TipoAdelanto = 'Vacaciones')"
                                  Me.DtaAdelanto.Refresh
                                  AdelantoVaca = 0
                                
                                  Do While Not DtaAdelanto.Recordset.EOF
                                     AdelantoVaca = AdelantoVaca + DtaAdelanto.Recordset("MontoAdelanto")
                                     DtaAdelanto.Recordset.MoveNext
                                  Loop
                               
                                 If val(SalMayor) < 0 Then
                                  SalMayor = 0
                                 End If
                                   
                                If CantMeses <> 0 Then
                                  If DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
                                    SalMayor = (SalMayor / CantMeses)
                                  Else
                                    SalMayor = SalMayor + MontoSubsidio
                                  End If
                               End If
                                  Año = Year(Me.TxtFFinVaca.Value)
                         
                         
                         
                         
                         Else '///Si es salario fijo le calculo el ultimo salario ////////
                         
                        '/////////////Busco el Adelanto de Vacaciones Registrados//////////////
                        '///////////////////////////////////////////////////////////////////////
                        '      cnDB.ConnectionString = Conexion
                        '      cnDB.Open
                        
                              
                        '      If Month(FechaHoy) <= 6 Then
                                '/////////////CAlculo las vacaciones Enero - Junio ////////////////////////////
                                SalMayor = 0
                                CantMeses = 0
                                 
                        '         rsDB.Open "SELECT TipoNomina.Nomina, Fecha_Planilla.Periodo, Fecha_Planilla.año, Fecha_Planilla.mes, Fecha_Planilla.Inicio, Fecha_Planilla.Final, " & _
                        '          "Fecha_Planilla.Actual , Fecha_Planilla.NumNomina, Fecha_Planilla.Calculada " & _
                        '          "FROM Fecha_Planilla INNER JOIN TipoNomina ON Fecha_Planilla.CodTipoNomina = TipoNomina.CodTipoNomina " & _
                        '          "WHERE TipoNomina.Nomina LIKE '" & Me.DBCNominas.Text & "' AND Fecha_Planilla.Inicio >= CONVERT(DATETIME, '" & Format(Me.dtpFPInicio.Value, "yyyy-mm-dd") & " 00:00:00', 102) AND Fecha_Planilla.Final <= CONVERT(DATETIME, '" & Format(Me.dtpFPFinal.Value, "yyyy-mm-dd") & " 00:00:00', 102) ORDER BY Fecha_Planilla.Inicio ASC", cnDB
                        '
                        '
                                
                                
                                Do While Not rsDB.EOF
                                 
                                   If iMes <> CInt(rsDB.Fields("mes")) Then
                                      CantMeses = CantMeses + 1
                                      iMes = CInt(rsDB.Fields("mes"))
                                   End If
                                   
                                   rsDB.MoveNext
                                 Loop
                        
                        
                        
                        
                        
                        '        For Mes = 1 To 6
                        '         SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where ((Month([Nomina].[FechaNomina])) =  '" & Mes & "' ) And ((Year([Nomina].[FechaNomina])) ='" & Anno & "') and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
                        '            DtaNominas.RecordSource = SqlNominas
                        '            DtaNominas.Refresh
                        '            Fecha = "01/01/" & Anno
                        '         NumFecha1 = Fecha
                                 Fecha = "30/06/" & Anno
                                 NumFecha2 = Fecha
                               ' Next
                               
                        '      Else
                        '       For Mes = 7 To 12
                        '       '///////////////Busco las Vacaciones de Julio a Diciembre/////////////////////////////////////
                        '       'Julio- Diciembre
                        '        SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where ((Month([Nomina].[FechaNomina])) =  '" & Mes & "' ) And ((Year([Nomina].[FechaNomina])) ='" & Anno & "') and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
                        '            DtaNominas.RecordSource = SqlNominas
                        '            DtaNominas.Refresh
                        '               SalTemp = 0
                        '
                        '        Fecha = "01/07/" & Anno
                        '        NumFecha1 = Fecha
                        '        Fecha = "31/12/" & Anno
                        '        NumFecha2 = Fecha
                        '       Next
                        '
                        '      End If
                        
                                 Me.DtaAdelanto.RecordSource = "SELECT  CodEmpleado, FechaAdelanto, MontoAdelanto, [Ref/Cheque], TipoAdelanto From Adelanto13vo WHERE     (FechaAdelanto BETWEEN CONVERT(DATETIME, '" & Format(Me.dtpFPInicio.Value, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.dtpFPFinal.Value, "yyyymmdd") & "', 102)) AND (TipoAdelanto = 'Vacaciones')AND (CodEmpleado = '" & CodEmpleado & "') "
                        '         Me.DtaAdelanto.RecordSource = "SELECT  CodEmpleado, FechaAdelanto, MontoAdelanto, [Ref/Cheque], TipoAdelanto From Adelanto13vo WHERE     (FechaAdelanto BETWEEN CONVERT(DATETIME, '2007-07-01 00:00:00', 102) AND CONVERT(DATETIME, '2007-12-31 00:00:00', 102)) AND (TipoAdelanto = 'Vacaciones')AND (CodEmpleado = '" & CodEmpleado & "') "
                        '        Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo WHERE (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Adelanto13vo.FechaAdelanto) Between  " & NumFecha1 & " And " & NumFecha2 & ") AND ((Adelanto13vo.TipoAdelanto)='Vacaciones'))"
                                'Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo WHERE (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Adelanto13vo.FechaAdelanto) Between " & NumFecha1 & " And " & NumFecha2 & "))"
                                Me.DtaAdelanto.Refresh
                                AdelantoVaca = 0
                                
                                Do While Not DtaAdelanto.Recordset.EOF
                                 AdelantoVaca = AdelantoVaca + DtaAdelanto.Recordset("MontoAdelanto")
                                 DtaAdelanto.Recordset.MoveNext
                                Loop
                         
                          
                         
                         
                            SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]+[DetalleNomina].[IncetivoProduccion]+ +[DetalleNomina].[OtrosIngresos] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where (((DetalleNomina.CodEmpleado) = '" & CodEmpleado & "'))"
                            DtaNominas.RecordSource = SqlNominas
                            DtaNominas.Refresh
                            If DtaNominas.Recordset.EOF Then
                             Edicion = False
                             'DtaNominas.Recordset.MoveLast
                            End If
                             
                          '///////Selecciono el Salario Mayor de la Tabla Empleados/////////////////
                             SalMayor = DtaEmpleados.Recordset("SueldoPeriodo")
                            
                            'dependiendo del tipo de pago se hace el calculo del salario básico
                            
                            If DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
                                 SalMayor = SalMayor
                            ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
                                  SalMayor = SalMayor
                            ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
                                 SalMayor = SalMayor
                            End If
                            
                            
                            
                            
                            
                            
                           If Month(FechaHoy) <= 6 Then
                              SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]+ [DetalleNomina].[IncentivoProduccion]+ [DetalleNomina].[OtrosIngresos] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina WHERE (((DetalleNomina.CodEmpleado)='" & CodEmpleado & "') AND ((Month([Nomina].[FechaNomina])) Between 1 And 6) AND ((Year([Nomina].[FechaNomina]))= " & Anno & " ))"
                              DtaNominas.RecordSource = SqlNominas
                              DtaNominas.Refresh
                              If Not DtaNominas.Recordset.EOF Then
                               DtaNominas.Recordset.MoveLast
                              End If
                              CantRegistros = 0
                              
                           Else
                           
                             SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]+ [DetalleNomina].[IncetivoProduccion]+ [DetalleNomina].[OtrosIngresos] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina WHERE (((DetalleNomina.CodEmpleado)='" & CodEmpleado & "') AND ((Month([Nomina].[FechaNomina])) Between 7 And 12) AND ((Year([Nomina].[FechaNomina]))= " & Anno & " ))"
                              DtaNominas.RecordSource = SqlNominas
                              DtaNominas.Refresh
                             If Not DtaNominas.Recordset.EOF Then
                              DtaNominas.Recordset.MoveLast
                             End If
                              CantRegistros = 0
                           End If
                        
                           
                         End If  '/////////Fin del If salario fijo/////////////////
                                
                            'dependiendo del tipo de pago se hace el calculo del salario básico
                            
                            
                            
                             If DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
                                 If CantRegistros > 0 Then
                                   SalMayor = SalMayor / CantMeses
                                 Else
                                   SalMayor = SalMayor * 3
                                   CantRegistros = DtaNominas.Recordset.RecordCount
                                 End If
                             ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
                                  If CantRegistros > 0 Then
                                   SalMayor = SalMayor / CantMeses
                                  Else
                                   SalMayor = SalMayor * 6
                                   CantRegistros = DtaNominas.Recordset.RecordCount
                                  End If
                             ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
                        '         If CantRegistros > 0 Then
                                  If Dias < 182 Then
                                   'SalMayor = (Dias * (SalMayor * 2) / 30.4167) * 0.08333333
                        '          Else
                        '           SalMayor = 0
                                  End If
                        '         Else
                        '           SalMayor = SalMayor * 2
                        '           CantRegistros = DtaNominas.Recordset.RecordCount
                        '          End If
                             ElseIf DtaTipoNomina.Recordset("Periodo") = "Semanal Viernes" Then
                                 If CantRegistros > 0 Then
                                  If CantMeses <> 0 Then
'                                   SalMayor = (SalMayor / CantMeses)
                                    SalMayor = (SalMayor / 6)
                                    '//////////////////VERIFICO SI EL SALARIO PROMEDIO ES MENOR QUE EL SALARIO MINIMO ///////
                                    If SalMayor < Format(TarifaHoraria * DiasMes * 8, "####0.00") Then
                                      SalMayor = Format(TarifaHoraria * DiasMes * 8, "####0.00")
                                    End If
                                    
                                  Else
                                   SalMayor = 0
                                  End If
                                 Else
                                   SalMayor = SalMayor * 2
'                                   CantRegistros = Me.DtaNominas.Recordset.RecordCount
                                  End If
                             ElseIf DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
                                 If CantRegistros > 0 Then
                                  If CantMeses <> 0 Then
'                                   SalMayor = (SalMayor / CantMeses)
                                    SalMayor = (SalMayor / 6)
                                    '//////////////////VERIFICO SI EL SALARIO PROMEDIO ES MENOR QUE EL SALARIO MINIMO ///////
                                    If SalMayor < Format(TarifaHoraria * DiasMes * 8, "####0.00") Then
                                      SalMayor = Format(TarifaHoraria * DiasMes * 8, "####0.00")
                                    End If
                                    
                                  Else
                                   SalMayor = 0
                                  End If
                                 Else
                                   SalMayor = SalMayor * 2
'                                   CantRegistros = Me.DtaNominas.Recordset.RecordCount
                                  End If
                             End If
                             
                         
                        DtaNomVaca.Refresh
                        Do While Not DtaNomVaca.Recordset.EOF
                            If DtaNomVaca.Recordset("NumNomVaca") = val(TxtNumNomVaca.Text) And DtaNomVaca.Recordset("Activa") = True Then
                               'DtaNomVaca.Recordset.Edit
                               DtaNomVaca.Recordset("montopagado") = SalMayor + DtaNomVaca.Recordset("montopagado")
                               'DtaNomVaca.Recordset.adelantovacaciones = AdelantoVaca
                               DtaNomVaca.Recordset.Update
                           Exit Do
                            End If
                        DtaNomVaca.Recordset.MoveNext
                        Loop
                        

                                
                            If SalarioBasico > SalMayor Then
                             SalMayor = SalarioBasico
                           End If
                           
                                                                                     
                               '---------------------------------------------------------------------------------------------------------------------
                               '-----------------------------------BUSCO LOS SALARIOS DEL PERIODO DE VACACIONES --------------------------------------
                               '----------------------------------------------------------------------------------------------------------------------
                                NumNomina = val(Me.TxtNumNomVaca.Text)
                                Me.DtaConsulta.RecordSource = "SELECT  * From Reembolso WHERE (NumNomina = " & NumNomina & " ) AND (CodEmpleado = '" & CodEmpleado & "')"
                                Me.DtaConsulta.Refresh
                                If Not Me.DtaConsulta.Recordset.EOF Then
                                  Monto = Me.DtaConsulta.Recordset("Monto")
                                Else
                                  Monto = 0
                                End If
                               
                               SqlSalarios = "SELECT DISTINCT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.HorasExtras) AS HorasExttras, SUM(DetalleNomina.BonoProduccion) AS BonoProduccion, SUM(DetalleNomina.SeptimoDia) AS SeptimoDia, SUM(DetalleNomina.OtrosIngresos) AS OtrosIngresos, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.Incentivos) AS Incentivos, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.Comisiones + DetalleNomina.HorasExtras + DetalleNomina.BonoProduccion + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos+DetalleNomina.IncetivoProduccion) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin,Nomina.Mes AS MES,Nomina.Ano AS AÑO,SUM(DetalleNomina.Comisiones) AS Comisiones,SUM(DetalleNomina.MontoIR) AS MontoIR,SUM(DetalleNomina.MontoINSS) As MontoINSS  " & _
                                             "FROM  DetalleNomina INNER JOIN  Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = " & CodEmpleado & ") AND (MAX(Nomina.FechaNomina) <= CONVERT(DATETIME, '" & Format(Me.dtpFPFinal.Value, "yyyy-mm-dd") & "', 102)) AND (MIN(Nomina.FechaNominaINI) >= CONVERT(DATETIME, '" & Format(Me.dtpFPInicio.Value, "yyyy-mm-dd") & "', 102))"
                           
                               Me.AdoBusca.RecordSource = SqlSalarios
                               Me.AdoBusca.Refresh
                               If Not Me.AdoBusca.Recordset.EOF Then
                                 
                                 IrAcumulado = Format(Me.AdoBusca.Recordset("MontoIR"), "####0.00")
                                 InssAcumulado = Format(Me.AdoBusca.Recordset("MontoINSS"), "####0.00")
                                 SalarioAcumulado = Format(Me.AdoBusca.Recordset("TotalIngresos"), "####0.00") + ((SalMayor / DiasMes) * (Dias / 12)) + Monto - InssAcumulado - (((SalMayor / DiasMes) * (Dias / 12)) + Monto) * (TasaInss / 100)
                                 
                               Else
                                 IrAcumulado = 0
                                 InssAcumulado = 0
                                 SalarioAcumulado = 0
                               End If
                               
                               IR = CalcularIr(SalarioAcumulado, "Semanal Viernes")
                               IrAcumulado = IR - IrAcumulado
                               If IrAcumulado < 0 Then
                                 IrAcumulado = 0
                               End If
                               
                                DiasDescuento = 0
                          If Edicion = True Then
                          
                          
                        
                           
                           Me.DtaDetalleNomVaca.RecordSource = "SELECT DetalleNomVaca.Id, DetalleNomVaca.TotalDevengado, DetalleNomVaca.Inss, DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones,DetalleNomVaca.Ir From DetalleNomVaca Where (((DetalleNomVaca.NumNomVaca) = " & val(TxtNumNomVaca.Text) & ") And ((DetalleNomVaca.CodEmpleado) = '" & CodEmpleado & "'))"
                           Me.DtaDetalleNomVaca.Refresh
                             
                           '///////Busco si el Empleado ya Existe en la Nomina de Vacaciones/////
                             If Not Me.DtaDetalleNomVaca.Recordset.EOF Then
                                DiasDescuento = DtaDetalleNomVaca.Recordset("DiasDescuento") + val(TxtDiasDescuento)
                                'DtaDetalleNomVaca.Recordset.Edit
                                DtaDetalleNomVaca.Recordset("NumNomVaca") = val(TxtNumNomVaca.Text)
                                DtaDetalleNomVaca.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
                                DtaDetalleNomVaca.Recordset("Ir") = IrAcumulado
                        '        If Me.DBCNominas.Text = "Administracion" Then
                                
                                  If DtaEmpleados.Recordset("SalarioFijo") = "S" Then
                                   
                                         
                        '         Else
                                    
                                    If Dias = 182 Then
                                       DtaDetalleNomVaca.Recordset("Inss") = SalMayor * (TasaInss / 100)
                                       DtaDetalleNomVaca.Recordset("TotalDevengado") = SalMayor
                                    Else
                                       DtaDetalleNomVaca.Recordset("Inss") = (Dias * (SalMayor * 2) / DiasMes) * 0.08333333 * (TasaInss / 100)
                                       DtaDetalleNomVaca.Recordset("TotalDevengado") = (Dias * (SalMayor * 2) / DiasMes) * 0.08333333
                                    End If
                                      
                                     DtaDetalleNomVaca.Recordset("SalarioMensual") = SalMayor * 2
                                      
                        '          End If
                                  
                                Else
                                    
                                    DtaDetalleNomVaca.Recordset("SalarioMensual") = SalMayor
'                                    DtaDetalleNomVaca.Recordset("Inss") = ((SalMayor / DiasMes) * ((Dias - DiasDescuento) * 0.0833333) - AdelantoVaca) * (TasaInss/100)
                                    DtaDetalleNomVaca.Recordset("Inss") = ((SalMayor / DiasMes) * ((Dias * 0.0833333) - DiasDescuento)) * (TasaInss / 100)
'                                    DtaDetalleNomVaca.Recordset("TotalDevengado") = ((SalMayor / DiasMes) * (Dias - DiasDescuento) * 0.0833333)
                                    DtaDetalleNomVaca.Recordset("TotalDevengado") = ((SalMayor / DiasMes) * ((Dias * 0.0833333) - DiasDescuento))
                                    
                                    
                                End If
                                
                                
                                
                                
                                
                             '///////////////////////////////////////////////////////////////////////////////////////////////////////
                             '////////////////CALCULO CUANTOS DIAS SE CONSIDERAN PARA EL PAGO///////////////////////////////////////
                             '//////SI ES MAYOR DE 15 REDONDEO A 15//////////////////////////////////////////////////////////////
                               DiasPagar = Format(Dias * 0.08333333, "##,##0.00")
                               If DiasPagar > 15 Then
                                 DiasPagar = 15
                               End If
                                
                               DtaDetalleNomVaca.Recordset("DiasAPagar") = DiasPagar
                                 DtaDetalleNomVaca.Recordset("AdelantoVacaciones") = AdelantoVaca
                                If IsNull(DtaDetalleNomVaca.Recordset("DiasDescuento")) Then
                                    DtaDetalleNomVaca.Recordset("DiasDescuento") = 0
                                Else
                                  DtaDetalleNomVaca.Recordset("DiasDescuento") = DtaDetalleNomVaca.Recordset("DiasDescuento") + val(TxtDiasDescuento)
                                
                                End If
                                DtaDetalleNomVaca.Recordset.Update
                             Else
                             
                                                          '///////////////////////////////////////////////////////////////////////////////////////////////////////
                             '////////////////CALCULO CUANTOS DIAS SE CONSIDERAN PARA EL PAGO///////////////////////////////////////
                             '//////SI ES MAYOR DE 15 REDONDEO A 15//////////////////////////////////////////////////////////////
                               DiasPagar = Format(Dias * 0.08333333, "##,##0.00")
                               If DiasPagar > 15 Then
                                 DiasPagar = 15
                               End If
                                

              
                             
                               Me.DtaDetalleNomVaca.RecordSource = "SELECT DetalleNomVaca.Id, DetalleNomVaca.Inss, DetalleNomVaca.Ir,DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones From DetalleNomVaca"
                               Me.DtaDetalleNomVaca.Refresh
                               If Me.DtaDetalleNomVaca.Recordset.EOF Then
                                 Id = 1
                               Else
                                 Me.DtaDetalleNomVaca.Recordset.MoveLast
                                 Id = Me.DtaDetalleNomVaca.Recordset("id") + 1
                               End If
                        '       DiasDescuento = DtaDetalleNomVaca.Recordset("DiasDescuento") + Val(TxtDiasDescuento)
                                DtaDetalleNomVaca.Recordset.AddNew
                                DtaDetalleNomVaca.Recordset("id") = Id
                                DtaDetalleNomVaca.Recordset("NumNomVaca") = val(TxtNumNomVaca.Text)
                                DtaDetalleNomVaca.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
                                DtaDetalleNomVaca.Recordset("SalarioMensual") = SalMayor
                        '        DtaDetalleNomVaca.Recordset("DiasAPagar") = 1.25 * CantRegistros
                                DtaDetalleNomVaca.Recordset("DiasAPagar") = DiasPagar
                                'DtaDetalleNomVaca.Recordset("Inss") = (SalMayor * (((Dias - DiasDescuento) * 0.0833333) / DiasMes) - AdelantoVaca) * (TasaInss/100)
                                DtaDetalleNomVaca.Recordset("Ir") = IrAcumulado
                                DtaDetalleNomVaca.Recordset("Inss") = (SalMayor * (((Dias * 0.0833333) - DiasDescuento) / DiasMes)) * (TasaInss / 100)
                                DtaDetalleNomVaca.Recordset("AdelantoVacaciones") = AdelantoVaca
                                'MsgBox DtaDetalleNomVaca.Recordset("DiasDescuento")
                                If IsNull(DtaDetalleNomVaca.Recordset("DiasDescuento")) Then
                                 DtaDetalleNomVaca.Recordset("DiasDescuento") = 0
                                Else
                                DtaDetalleNomVaca.Recordset("DiasDescuento") = DtaDetalleNomVaca.Recordset("DiasDescuento") + val(TxtDiasDescuento)
                                End If
                                DtaDetalleNomVaca.Recordset.Update
                             
                             End If
                           Else
                           
                               Me.DtaDetalleNomVaca.RecordSource = "SELECT DetalleNomVaca.Id, DetalleNomVaca.Ir,DetalleNomVaca.Inss, DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones From DetalleNomVaca"
                               Me.DtaDetalleNomVaca.Refresh
                               If Me.DtaDetalleNomVaca.Recordset.EOF Then
                                 Id = 1
                               Else
                                 Me.DtaDetalleNomVaca.Recordset.MoveLast
                                 Id = Me.DtaDetalleNomVaca.Recordset("id") + 1
                               End If
                                DiasDescuento = DtaDetalleNomVaca.Recordset("DiasDescuento") + val(TxtDiasDescuento)
                                DtaDetalleNomVaca.Recordset.AddNew
                                DtaDetalleNomVaca.Recordset("id") = Id
                                DtaDetalleNomVaca.Recordset("NumNomVaca") = val(TxtNumNomVaca.Text)
                                DtaDetalleNomVaca.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
                                DtaDetalleNomVaca.Recordset("SalarioMensual") = SalMayor
                                DtaDetalleNomVaca.Recordset("DiasAPagar") = DiasPagar
'                                DtaDetalleNomVaca.Recordset("Inss") = (SalMayor * (((Dias - DiasDescuento) * 0.0833333) / DiasMes) - AdelantoVaca) * (TasaInss/100)
                                DtaDetalleNomVaca.Recordset("Ir") = IrAcumulado
                                DtaDetalleNomVaca.Recordset("Inss") = (SalMayor * (((Dias * 0.0833333) - DiasDescuento) / DiasMes)) * (TasaInss / 100)
                                'DtaDetalleNomVaca.Recordset("DiasAPagar") = 2.5 * CantMeses
                                DtaDetalleNomVaca.Recordset("AdelantoVacaciones") = AdelantoVaca
                                'MsgBox DtaDetalleNomVaca.Recordset("DiasDescuento")
                                If IsNull(DtaDetalleNomVaca.Recordset("DiasDescuento")) Then
                                    DtaDetalleNomVaca.Recordset("DiasDescuento") = 0
                                Else
                                  DtaDetalleNomVaca.Recordset("DiasDescuento") = DtaDetalleNomVaca.Recordset("DiasDescuento") + val(TxtDiasDescuento)
                                
                                End If
                                DtaDetalleNomVaca.Recordset.Update
                         End If
                         
                         
                         Else
                            Me.DtaDetalleNomVaca.RecordSource = "SELECT DetalleNomVaca.Inss, DetalleNomVaca.Id, DetalleNomVaca.Inss, DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones From DetalleNomVaca Where (((DetalleNomVaca.NumNomVaca) = " & val(TxtNumNomVaca.Text) & ") And ((DetalleNomVaca.CodEmpleado) = '" & CodEmpleado & "'))"
                            Me.DtaDetalleNomVaca.Refresh
                            If Not Me.DtaDetalleNomVaca.Recordset.EOF Then
                            DtaDetalleNomVaca.Recordset("DiasAPagar") = 0
                             DtaDetalleNomVaca.Recordset("Inss") = 0
                            DtaDetalleNomVaca.Recordset.Update
                            End If
                        
                         End If
                         
                       
                    rsDB.Close
        
          End If
          
                        DtaEmpleados.Recordset.MoveNext
                        i = i + 1
                        DoEvents
                        
        Loop
End With

                 NumNomina = val(Me.TxtNumNomVaca.Text)
                 Me.DtaEmpleados.RecordSource = "SELECT  * From Reembolso Where (NumNomina = " & NumNomina & ")"
                 Me.DtaEmpleados.Refresh
                 If Not Me.DtaEmpleados.Recordset.EOF Then
                   
                        
                          
                           Me.DtaEmpleados.Recordset.MoveLast
                           CantRegistros = Me.DtaEmpleados.Recordset.RecordCount
                           Me.DtaEmpleados.Recordset.MoveFirst
                          
                          PBVacaciones.Min = 0
                          PBVacaciones.Max = CantRegistros
                          PBVacaciones.Value = 0
                          i = 0
                          
                          Do While Not Me.DtaEmpleados.Recordset.EOF
                          
                                                        Monto = DtaEmpleados.Recordset("Monto")
                          
                                                        Me.DtaDetalleNomVaca.RecordSource = "SELECT DetalleNomVaca.Id, DetalleNomVaca.Inss, DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones,DetalleNomVaca.TotalDevengado From DetalleNomVaca"
                                                        Me.DtaDetalleNomVaca.Refresh
                                                        If Me.DtaDetalleNomVaca.Recordset.EOF Then
                                                          Id = 1
                                                        Else
                                                          Me.DtaDetalleNomVaca.Recordset.MoveLast
                                                          Id = Me.DtaDetalleNomVaca.Recordset("id") + 1
                                                        End If
                                                        
                                                         CodEmpleado = DtaEmpleados.Recordset("CodEmpleado")
                                                         If ExisteEmpleado2(CodEmpleado) = True Then
                                                        
                                                                 Me.DtaConsulta.RecordSource = "SELECT  * From DetalleNomVaca WHERE (NumNomVaca = " & NumNomina & ") AND (CodEmpleado = " & CodEmpleado & ")"
                                                                 Me.DtaConsulta.Refresh
                                                                 If Me.DtaConsulta.Recordset.EOF Then
                                                                     DtaConsulta.Recordset.AddNew
                                                                     DtaConsulta.Recordset("id") = Id
                                                                     DtaConsulta.Recordset("NumNomVaca") = val(TxtNumNomVaca.Text)
                                                                     DtaConsulta.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
                                                                     DtaConsulta.Recordset("SalarioMensual") = Monto
                                                                     DtaConsulta.Recordset("DiasAPagar") = 30
                                                                     DtaConsulta.Recordset("AdelantoVacaciones") = 0
                                                                     DtaConsulta.Recordset("DiasDescuento") = 0
                                                                     DtaConsulta.Recordset("Inss") = Monto * (TasaInss / 100)
        '                                                             DtaConsulta.Recordset("Ir") = 0
                                                                     DtaConsulta.Recordset("TotalDevengado") = Monto
                                                                     DtaConsulta.Recordset.Update
                                                                 ElseIf DtaConsulta.Recordset("DiasAPagar") = 30 Then
                                                                     DtaConsulta.Recordset("SalarioMensual") = Monto
                                                                     DtaConsulta.Recordset("DiasAPagar") = 30
                                                                     DtaConsulta.Recordset("AdelantoVacaciones") = 0
                                                                     DtaConsulta.Recordset("DiasDescuento") = 0
                                                                     DtaConsulta.Recordset("Inss") = Monto * (TasaInss / 100)
        '                                                             DtaConsulta.Recordset("Ir") = 0
                                                                     DtaConsulta.Recordset("TotalDevengado") = Monto
                                                                     DtaConsulta.Recordset.Update
                                                                 Else
                                                                     
                                                                     DtaConsulta.Recordset("SalarioMensual") = DtaConsulta.Recordset("SalarioMensual") + Monto
                                                                     DtaConsulta.Recordset("Inss") = DtaConsulta.Recordset("Inss") + Monto * (TasaInss / 100)
                                                                     DtaConsulta.Recordset("TotalDevengado") = DtaConsulta.Recordset("TotalDevengado") + Monto
        '                                                             ((SalMayor / DiasMes) * ((Dias * 0.0833333) - DiasDescuento))
                                                                     DtaConsulta.Recordset.Update
                                                                 End If
                                                         End If
                                                              
                                  Me.DtaConsulta.Refresh
                           
                           DtaEmpleados.Recordset.MoveNext
                           i = i + 1
                           PBVacaciones.Value = i
                           DoEvents
                          Loop
                  End If
 
 
 
 
 
DtaVacaciones.Refresh
Me.DbgrVacaciones.Columns(0).Visible = False
Me.DbgrVacaciones.Columns(0).Locked = True
Me.DbgrVacaciones.Columns(1).Locked = True
Me.DbgrVacaciones.Columns(2).Locked = True
Me.DbgrVacaciones.Columns(3).Locked = True
Me.DbgrVacaciones.Columns(4).Locked = True
Me.DbgrVacaciones.Columns(5).Locked = True
Me.DbgrVacaciones.Columns(6).Locked = True
Me.DbgrVacaciones.Columns(7).Locked = True
Me.DbgrVacaciones.Columns(10).Locked = True
Me.DbgrVacaciones.Columns(6).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(7).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(8).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(9).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(10).NumberFormat = "##,##0.00"
Me.CmdPRVaca.Enabled = True
Me.CmdExportar.Visible = True
'Me.CmdMonedas.Visible = True
Me.CmdMonedasvaca.Enabled = True
Me.CmdNominaVaca.Enabled = True
Me.CmdColillaVaca.Enabled = True
MousePointer = 1

cnDB.Close
End Sub

Private Sub CmdCalVaca_Click()
Dim SalarioBasico As Double, DiasDescuento As Double
Dim SqlEmpleados As String, MontoSubsidio As Double
Dim CodTipoNomina As String, TotalSubsidio As Double
Dim SqlNominas As String, NumVaca As Integer
Dim SalMayor As Double, TarifaHoraria As Double
Dim SalTemp As Double, Dias As Double
Dim CodEmpleado As String, AdelantoVaca As Double
Dim Edicion As Boolean
Dim Anno As Integer, DiasAcumulados As Double
Dim Mes As Integer, Fecha As Date
Dim CantEmpleados As Long
Dim i As Integer, CantMeses As Integer, CantRegistros As Integer
Dim DiasMes As Double
Dim DiasSemana As Double, DiasPagar As Double
Dim FechaHoy As Date, DiasNomVaca As Double
Dim rsDB As New ADODB.Recordset, rs As New ADODB.Recordset
Dim cnDB As New ADODB.Connection
Dim iMes As Integer, DiasMenos As Double, CodEmpleado1 As String, TasaInss As Double
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer, SalarioAcumulado As Double, IR As Double, IrAcumulado As Double, InssAcumulado As Double
Dim SueldoActual As Boolean

Me.CmdExportar.Enabled = True

DtaControles.Refresh
DiasMes = DtaControles.Recordset("DiasMes")
DiasSemana = DtaControles.Recordset("DiasSemana")

    

FechaHoy = Format(Now, "dd/mm/yyyy")

If FechaHoy < Me.TxtFINIVaca.Value Then
 MsgBox "La fecha Actual no Coincide con la Nomina", vbCritical, "Sistema de NOminas"
 Exit Sub
ElseIf FechaHoy > Me.TxtFFinVaca.Value Then
  FechaHoy = Me.TxtFFinVaca
End If

Anno = Year(FechaHoy)

Edicion = False

MousePointer = 11
If DBCNominas.Text = "Listado de Nominas" Then
   MsgBox "No ha seleccionado el tipo de nomina al cual le desea calcular las Vacaciones"
   MousePointer = 1
   DBCNominas.SetFocus
   Exit Sub
End If


If Not IsNumeric(TxtDiasDescuento.Text) Then
   MsgBox "Los dias de descuendo de Vacaciones son erróneos"
   MousePointer = 1
   TxtDiasDescuento.SetFocus
   Exit Sub
ElseIf val(TxtDiasDescuento.Text) > 15 Then
   MsgBox "Los dias de descuendo de Vacaciones no pueden ser mayor de 15"
   MousePointer = 1
   TxtDiasDescuento.SetFocus
   Exit Sub
End If

DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop


     MDIPrimero.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, PorcientoInss, TasaInssPatronal, TasaInss, PorcientoIr, TasaIr From TipoNomina WHERE (PorcientoInss = 1) AND (CodTipoNomina = '" & CodTipoNomina & "' )"
     MDIPrimero.DtaConsulta.Refresh
     If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
       TasaInss = MDIPrimero.DtaConsulta.Recordset("TasaInss")
     End If


NumVaca = val(TxtNumNomVaca.Text)

If Me.ChkEliminar.Value = 1 Then
 Fecha1 = Format(Me.TxtFINIVaca.Value, "yyyy-mm-dd")
 Fecha2 = Format(Me.TxtFFinVaca.Value, "yyyy-mm-dd")
 Set Ejecutar = New ADODB.Connection
 Ejecutar.ConnectionString = Conexion
 Ejecutar.Open
 Ejecutar.Execute "DELETE FROM DetalleNomVaca WHERE (NumNomVaca =  " & NumVaca & ")"
 Ejecutar.Execute "DELETE FROM HistorialSalarioMes WHERE (FechaIni = CONVERT(DATETIME, '" & Fecha1 & "', 102)) AND (FechaFin = CONVERT(DATETIME, '" & Fecha2 & "', 102))"
End If

'pregunto si ya existe una nomina de vacaciones elaborada
Me.DtaNomVaca.RecordSource = "SELECT CodTipoNomina, NumNomVaca, FechaAplica, MontoPagado, FechaIni, FechaFin, Activa, Transfereir From NomVaca Where (Activa = 1) And (NumNomVaca = " & NumVaca & ")"
DtaNomVaca.Refresh
If Not DtaNomVaca.Recordset.EOF Then
    'If DtaNomVaca.Recordset("NumNomVaca") = Val(TxtNumNomVaca.Text) And DtaNomVaca.Recordset.Activa = True Then
       'DtaNomVaca.Recordset.Edit
       DtaNomVaca.Recordset("FechaAplica") = Me.TxtFechaAplica.Value
       DtaNomVaca.Recordset("montopagado") = 0
       DtaNomVaca.Recordset("Fechaini") = CDate(TxtFINIVaca.Value)
       DtaNomVaca.Recordset("Fechafin") = CDate(TxtFFinVaca.Value)
       DtaNomVaca.Recordset("CodTipoNomina") = CodTipoNomina
       
       If Me.CHKTranferir.Value = 1 Then
        DtaNomVaca.Recordset("Transfereir") = True
       Else
        DtaNomVaca.Recordset("Transfereir") = False
       End If
       
       DtaNomVaca.Recordset.Update
       Edicion = True
    
    'End If
End If

If Not Edicion Then 'es primera vez que se crearan las vacaciones

       DtaNomVaca.Recordset.AddNew
       DtaNomVaca.Recordset("NumNomVaca") = val(TxtNumNomVaca.Text)
       DtaNomVaca.Recordset("FechaAplica") = Format(Now, "DD/MM/YYYY")
       DtaNomVaca.Recordset("montopagado") = 0
       DtaNomVaca.Recordset("Fechaini") = CDate(TxtFINIVaca.Value)
       DtaNomVaca.Recordset("Fechafin") = CDate(TxtFFinVaca.Value)
       DtaNomVaca.Recordset("CodTipoNomina") = CodTipoNomina
       DtaNomVaca.Recordset("Activa") = 1
       DtaNomVaca.Recordset.Update
       
Else 'se borran los movimientos anteriores

     'DtaDetalleNomVaca.Refresh
     'Do While Not DtaDetalleNomVaca.Recordset.EOF
     'If DtaDetalleNomVaca.Recordset("NumNomVaca") = Val(TxtNumNomVaca) Then
      '  DtaDetalleNomVaca.Recordset.Delete
     'End If
     
     'DtaDetalleNomVaca.Recordset.MoveNext
     'Loop
     
     
End If

'hago el sql de los empleados que pertenezcan solo a la nomina seleccionada

'SqlEmpleados = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, Empleado.ExentoIr, Empleado.OtrosIngresos,  Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, Empleado.Ausente, Empleado.SalarioFijo From Empleado WHERE Empleado.CodTipoNomina= '" & CodTipoNomina & "' AND Empleado.Activo=1 AND Empleado.Ausente=0"
SqlEmpleados = "SELECT  Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, Empleado.ExentoIr, Empleado.OtrosIngresos, Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, Empleado.Ausente, Empleado.SalarioFijo, Historico.FechaContrato FROM  Empleado INNER JOIN  Historico ON Empleado.CodEmpleado = Historico.Codempleado  " & _
               "WHERE (Empleado.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (Empleado.Ausente = 0) AND (Historico.FechaContrato <= CONVERT(DATETIME, '" & Format(Me.dtpFPFinal.Value, "yyyy-mm-dd") & "', 102))"
DtaEmpleados.RecordSource = SqlEmpleados
DtaEmpleados.Refresh

DtaEmpleados.Recordset.MoveLast
CantEmpleados = DtaEmpleados.Recordset.RecordCount

PBVacaciones.Value = 0

With PBVacaciones
 .Min = 0
 .Max = CantEmpleados
 .Value = 0
 i = 1
DtaEmpleados.Refresh

cnDB.ConnectionString = Conexion
cnDB.Open

'/////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////ELIMINO LOS HISTORIALES PARA LA NOMINA ////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////
rs.Open "DELETE FROM HistorialSalarioMes Where (NumNomina = " & TxtNumNomVaca.Text & ")", Conexion

          
          
'recorro la BD empleados y a cada uno le busco su salario mayor (sies destajo) si no solo extraigo su salario
Do While Not DtaEmpleados.Recordset.EOF

  rsDB.Open "SELECT TipoNomina.Nomina, Fecha_Planilla.Periodo, Fecha_Planilla.año, Fecha_Planilla.mes, Fecha_Planilla.Inicio, Fecha_Planilla.Final, " & _
          "Fecha_Planilla.Actual , Fecha_Planilla.NumNomina, Fecha_Planilla.Calculada " & _
          "FROM Fecha_Planilla INNER JOIN TipoNomina ON Fecha_Planilla.CodTipoNomina = TipoNomina.CodTipoNomina " & _
          "WHERE TipoNomina.Nomina LIKE '" & Me.DBCNominas.Text & "' AND Fecha_Planilla.Inicio >= CONVERT(DATETIME, '" & Format(Me.dtpFPInicio.Value, "yyyy-mm-dd") & " 00:00:00', 102) AND Fecha_Planilla.Final <= CONVERT(DATETIME, '" & Format(Me.dtpFPFinal.Value, "yyyy-mm-dd") & " 00:00:00', 102) ORDER BY Fecha_Planilla.Inicio ASC", cnDB

   

     .Value = i
      'Me.xp_canvas1.Caption = "Procesando:  " & i & " de " & CantEmpleados & " Empleados "
      Me.Caption = "Procesando:  " & i & " de " & CantEmpleados & " Empleados "
      Me.LblTotal.Caption = "Procesando:  " & i & " de " & CantEmpleados & " Empleados "
      DoEvents
     CodEmpleado = DtaEmpleados.Recordset("CodEmpleado")
     CodEmpleado1 = DtaEmpleados.Recordset("CodEmpleado1")
     
    'tengo que hacer un SQL de Solo los que esten en el rango de fechas
    'solo se veran las nominas de cada mes
    'se deben de hacer ciclos por cada mes, seis ciclos por los seis meses.
 SalMayor = 0
 CantMeses = 0
 CantRegistros = 0
 DiasAcumulados = 0
 Dias = 0
 
                             If CodEmpleado1 = "S117090014" Then
                              CodEmpleado1 = "S117090014"
                             End If

 TasaCambio = BuscaTasaCambio(Me.TxtFFinVaca.Value)
 
 
  '/////////////////ESTE ES EL SALARIO BASICO, SIN PRODUCCION////////////////
     TarifaHoraria = DtaEmpleados.Recordset("TarifaHoraria")

     
     If TarifaHoraria = 0 Then
                '//////////////////////////////////////////////////////////////////////////////////////
                '//////////////////CALCULO EL SALARIO X HORA PARA LOS EMPLEADS QUE SON FIJOS////
                '////////////////////////////////////////////////////////////////////////////////////
                   Select Case DtaTipoNomina.Recordset("Periodo")
                        Case "Catorcenal los Sabados"
                        
                            TarifaHoraria = Format(DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio / 112, "###,##0.00")
                            
                        Case "Quincenal"
                            TarifaHoraria = Format(DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio / ((DiasMes * 8) / 2), "###,##0.00")
                            TotalHoras = 15 '* DtaTipoNomina.Recordset("Horas") '////LE ESCRIBO EL TOTAL DE DIAS ///
                        Case "Mensual"
                            TarifaHoraria = Format(DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio / (DiasMes * 8), "###,##0.00")
                        Case "Trimestral"
                            TarifaHoraria = Format(DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio / (DiasMes * 8 * 3), "###,##0.00")
                        Case "Semestral"
                            TarifaHoraria = Format(DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio / (DiasMes * 8 * 6), "###,##0.00")
                  End Select
                        
                        
     End If
     
     
  SalarioBasico = Format(DiasMes * 8 * TarifaHoraria, "##,##0.00")


  DtaHistorico.RecordSource = "SELECT Historico.Codempleado, Historico.FechaBaja, Historico.FechaContrato, Historico.FechaContratoVac From Historico Where (((Historico.CodEmpleado) = '" & CodEmpleado & "'))"
  DtaHistorico.Refresh
  NumFecha1 = CDate(Me.TxtFINI13)
  NumFecha2 = CDate(Me.TxtFFIN13)
 
 
 
  '////////Verifico cuantos dias tiene de trabajar///////////////////
  If Not Me.DtaHistorico.Recordset.EOF Then
   If Not IsNull(DtaHistorico.Recordset("FechaContrato")) Then
     FechaContrato = DtaHistorico.Recordset("FechaContrato")
     NumFecha2 = FechaContrato
'     NumFecha1 = Me.TxtFFinVaca.Value
      NumFecha1 = Me.dtpFPFinal.Value
     annos = ((CDbl(NumFecha1) - CDbl(NumFecha2)) + 1) / 365
     Meses = ((CDbl(NumFecha1) - CDbl(NumFecha2)) + 1) / DiasMes
         

      FechaInicioAgui = DateSerial(Year(Me.TxtFFinVaca.Value), 1, 1)
      If FechaContrato > FechaInicioAgui Then
        If FechaContrato < CDate(Me.TxtFFinVaca.Value) Then
          FechaInicioAgui = FechaContrato
        End If
      End If

        FechaInicioAgui = DtaHistorico.Recordset("FechaContratoVac")
'         DiasNomVaca = CalcularDiasVaca(CDate(FechaInicioAgui), Me.TxtFFinVaca.Value)
         Dias = CalculoDiasVacaciones(CodEmpleado1, Me.TxtFFinVaca.Value)
'         CalculoDiasVacaciones(CodigoEmpleado, Me.dtpFin.Value)
         DiasNomVaca = CalculoDiasVacaSFicha(CodEmpleado1, Me.TxtFFinVaca.Value)
         
     If Dias < 0 Then
      
       Dias = 0
     ElseIf Dias > 15 Then
       Dias = 15
     End If
    End If
   End If
   
'   DiasNomVaca = Format(DiasNomVaca * 0.0833, "####0.00")
   If DiasNomVaca > 15 Then
     DiasNomVaca = 15
   End If
   
   
If CodEmpleado1 = "S117090014" Then
 CodEmpleado1 = "S117090014"
End If

SueldoActual = False
 
 
 If Dias >= CInt(Me.txtAntiguedad.Text) Then
 
     If DtaEmpleados.Recordset("SalarioFijo") = "N" Then
         Año = Year(Me.TxtFFinVaca.Value)
 '///////////Si el Salario es Variable Busco el Salario Mayor/////////
      
        '/////////////CAlculo las vacaciones Enero - Junio ////////////////////////////
        SalMayor = 0
        
       
        
         Do While Not rsDB.EOF And Dias > CInt(Me.txtAntiguedad.Text)  '30   '    For Mes = 1 To 6
      
      
              '///////////////////////////////////////////////////////////////////////////////////
              '//////////////////////////SI EL SALARIO ES BASICO NO SE TOMA EN CUENTA PROMEDIO/////////////
              '///////////////////////////////////////////////////////////////////////////////////
              MDIPrimero.DtaConsulta.RecordSource = "SELECT  Historico.*, Empleado.SueldoActualBasico FROM Historico INNER JOIN Empleado ON Historico.Codempleado = Empleado.CodEmpleado  Where (Historico.CodEmpleado = " & CodEmpleado & ")"
              MDIPrimero.DtaConsulta.Refresh
              If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
                If MDIPrimero.DtaConsulta.Recordset("SueldoActualBasico") = True Then
                
                   SueldoActual = True
                   If Not IsNull(MDIPrimero.DtaConsulta.Recordset("SueldoActual")) Then
                    SalarioBasico = MDIPrimero.DtaConsulta.Recordset("SueldoActual")
                   Else
                    SalarioBasico = 0
                   End If
                   
                   
                     
                    '"SUM(DetalleNomina.Destajo + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion) + " & SalarioBasico & " AS TotalIngresos, " & _

                 SqlNominas = "SELECT DISTINCT " & _
                            "DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, " & _
                            "SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos, " & _
                            "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion + DetalleNomina.HorasExtras)  AS TotalIngresos, " & _
                            "MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO " & _
                            "FROM  DetalleNomina INNER JOIN " & _
                            "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina " & _
                            "GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                            "HAVING      (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = " & CodEmpleado & ") AND (Nomina.Mes = " & CInt(rsDB.Fields("mes")) & ") AND " & _
                            "(Nomina.Ano = " & rsDB.Fields("año") & ") ORDER BY Nomina.Ano, Nomina.Mes "
                    
                Else
                     
                 SueldoActual = False
                 
                 SqlNominas = "SELECT DISTINCT " & _
                            "DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, " & _
                            "SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos, " & _
                            "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion + DetalleNomina.HorasExtras) AS TotalIngresos, " & _
                            "MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO " & _
                            "FROM  DetalleNomina INNER JOIN " & _
                            "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina " & _
                            "GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                            "HAVING      (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = " & CodEmpleado & ") AND (Nomina.Mes = " & CInt(rsDB.Fields("mes")) & ") AND " & _
                            "(Nomina.Ano = " & rsDB.Fields("año") & ") ORDER BY Nomina.Ano, Nomina.Mes "
 
                    
                    
                End If
              End If
              
               DtaNominas.RecordSource = SqlNominas
               DtaNominas.Refresh
            
               If Month(Me.TxtFFinVaca.Value) <= 6 Then
                  Fecha = "01/01/" & Anno
                  NumFecha1 = Fecha
                  Fecha = "30/06/" & Anno
                  NumFecha2 = Fecha
               Else
                  Fecha = "01/06/" & Anno
                  NumFecha1 = Fecha
                  Fecha = "30/11/" & Anno
                  NumFecha2 = Fecha
               
               End If

        
        
                  SalTemp = 0
               If Not DtaNominas.Recordset.EOF Then
                CantMeses = CantMeses + 1
               End If
               
              Do While Not DtaNominas.Recordset.EOF
                 If Not IsNull(DtaNominas.Recordset("TotalIngresos")) Then
                     SalTemp = SalTemp + CDbl(DtaNominas.Recordset("Totalingresos"))
                            
                 End If
               
                 CantRegistros = CantRegistros + 1
                 DtaNominas.Recordset.MoveNext
              Loop
              
'///////////////////////////////////////////////////////////////////////////////////
'//////////////////////////SI NO TIENE SALARIO EN UN MES NO SUMA PARA EL PROMEDIO/////////////
'///////////////////////////////////////////////////////////////////////////////////
              If SalTemp = 0 Then
'               CantMeses = CantMeses - 1
              End If
              
             
              
        
              Fecha1 = Year(Me.TxtFINIVaca.Value) & "-" & Month(Me.TxtFINIVaca.Value) & "-" & Day(Me.TxtFINIVaca.Value)
              Fecha2 = Year(Me.TxtFFinVaca.Value) & "-" & Month(Me.TxtFFinVaca.Value) & "-" & Day(Me.TxtFFinVaca.Value)
              
              If SalTemp <> 0 Then
                 Me.AdoHistorialSalarial.RecordSource = "SELECT NumNomina, Tipo, CodEmpleado, FechaIni, FechaFin, Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, Agosto, Septiembre, Octubre, Noviembre, Diciembre From HistorialSalarioMes WHERE (CodEmpleado = " & CodEmpleado & ") AND (NumNomina = " & Me.TxtNumNomVaca.Text & " )"
                 Me.AdoHistorialSalarial.Refresh
                 
                 '///////////////////////////////////////////////////VALIDO LA FECHA //////////////////////////
                 If Not Me.AdoHistorialSalarial.Recordset.EOF Then
                 
                   Dim FechaIni As Date, FechaFin As Date
                   FechaIni = Me.AdoHistorialSalarial.Recordset("FechaIni")
                   FechaFin = Me.AdoHistorialSalarial.Recordset("FechaFin")
                   
                  
                   If CDate(Fecha1) <> FechaIni Then
                      rs.Open "DELETE FROM HistorialSalarioMes WHERE (CodEmpleado = " & CodEmpleado & ") AND (NumNomina = " & Me.TxtNumNomVaca.Text & " )", Conexion
                   End If

                   If CDate(Fecha2) <> FechaFin Then
                      rs.Open "DELETE FROM HistorialSalarioMes WHERE (CodEmpleado = " & CodEmpleado & ") AND (NumNomina = " & Me.TxtNumNomVaca.Text & " )", Conexion
                   End If
                   
                    Me.AdoHistorialSalarial.Refresh
                 End If
                 
                  If Me.AdoHistorialSalarial.Recordset.EOF Then
                        Me.AdoHistorialSalarial.Recordset.AddNew
                        Me.AdoHistorialSalarial.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
                        Me.AdoHistorialSalarial.Recordset("FechaIni") = CDate(Me.TxtFINIVaca.Value)
                        Me.AdoHistorialSalarial.Recordset("FechaFin") = CDate(Me.TxtFFinVaca.Value)
                        Me.AdoHistorialSalarial.Recordset("NumNomina") = val(Me.TxtNumNomVaca.Text)
                        Me.AdoHistorialSalarial.Recordset("Tipo") = "Vacaciones"
                        
                        Select Case (CInt(rsDB.Fields("mes")))
                          Case 1
                            Me.AdoHistorialSalarial.Recordset("Enero") = SalTemp
                          Case 2
                            Me.AdoHistorialSalarial.Recordset("Febrero") = SalTemp
                          Case 3
                            Me.AdoHistorialSalarial.Recordset("Marzo") = SalTemp
                          Case 4
                            Me.AdoHistorialSalarial.Recordset("Abril") = SalTemp
                          Case 5
                            Me.AdoHistorialSalarial.Recordset("Mayo") = SalTemp
                          Case 6
                            Me.AdoHistorialSalarial.Recordset("Junio") = SalTemp
                          Case 7
                            Me.AdoHistorialSalarial.Recordset("Julio") = SalTemp
                          Case 8
                            Me.AdoHistorialSalarial.Recordset("Agosto") = SalTemp
                          Case 9
                            Me.AdoHistorialSalarial.Recordset("Septiembre") = SalTemp
                          Case 10
                            Me.AdoHistorialSalarial.Recordset("Octubre") = SalTemp
                          Case 11
                            Me.AdoHistorialSalarial.Recordset("Noviembre") = SalTemp
                          Case 12
                            Me.AdoHistorialSalarial.Recordset("Diciembre") = SalTemp
                         End Select
                        Me.AdoHistorialSalarial.Recordset.Update
                  Else
                         Me.AdoHistorialSalarial.Recordset("FechaIni") = CDate(Me.TxtFINIVaca.Value)
                         Me.AdoHistorialSalarial.Recordset("FechaFin") = CDate(Me.TxtFFinVaca.Value)
                         Me.AdoHistorialSalarial.Recordset("NumNomina") = val(Me.TxtNumNomVaca.Text)
                         Me.AdoHistorialSalarial.Recordset("Tipo") = "Vacaciones"
                        Select Case (CInt(rsDB.Fields("mes")))
                          Case 1
                            Me.AdoHistorialSalarial.Recordset("Enero") = SalTemp
                          Case 2
                            Me.AdoHistorialSalarial.Recordset("Febrero") = SalTemp
                          Case 3
                            Me.AdoHistorialSalarial.Recordset("Marzo") = SalTemp
                          Case 4
                            Me.AdoHistorialSalarial.Recordset("Abril") = SalTemp
                          Case 5
                            Me.AdoHistorialSalarial.Recordset("Mayo") = SalTemp
                          Case 6
                            Me.AdoHistorialSalarial.Recordset("Junio") = SalTemp
                          Case 7
                            Me.AdoHistorialSalarial.Recordset("Julio") = SalTemp
                          Case 8
                            Me.AdoHistorialSalarial.Recordset("Agosto") = SalTemp
                          Case 9
                            Me.AdoHistorialSalarial.Recordset("Septiembre") = SalTemp
                          Case 10
                            Me.AdoHistorialSalarial.Recordset("Octubre") = SalTemp
                          Case 11
                            Me.AdoHistorialSalarial.Recordset("Noviembre") = SalTemp
                          Case 12
                            Me.AdoHistorialSalarial.Recordset("Diciembre") = SalTemp
                        End Select
                        Me.AdoHistorialSalarial.Recordset.Update
       
                  End If
            
            Else
            Me.AdoHistorialSalarial.RecordSource = "SELECT NumNomina, CodEmpleado, FechaIni, FechaFin, Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, Agosto, Septiembre, Octubre, Noviembre, Diciembre, Tipo From HistorialSalarioMes WHERE     (Tipo = N'Vacaciones') AND (CodEmpleado = '" & CodEmpleado & "') AND (FechaIni = CONVERT(DATETIME, '" & Fecha1 & "', 102)) AND (FechaFin = CONVERT(DATETIME, '" & Fecha2 & "', 102)) "
            
                 Me.AdoHistorialSalarial.Refresh
                  If Not Me.AdoHistorialSalarial.Recordset.EOF Then
 
                        Select Case (CInt(rsDB.Fields("mes")))
                          Case 1
                            SalTemp = Me.AdoHistorialSalarial.Recordset("Enero")
                          Case 2
                            SalTemp = Me.AdoHistorialSalarial.Recordset("Febrero")
                          Case 3
                            SalTemp = Me.AdoHistorialSalarial.Recordset("Marzo")
                          Case 4
                           SalTemp = Me.AdoHistorialSalarial.Recordset("Abril")
                          Case 5
                            SalTemp = Me.AdoHistorialSalarial.Recordset("Mayo")
                          Case 6
                            SalTemp = Me.AdoHistorialSalarial.Recordset("Junio")
                          Case 7
                            SalTemp = Me.AdoHistorialSalarial.Recordset("Julio")
                          Case 8
                            SalTemp = Me.AdoHistorialSalarial.Recordset("Agosto")
                          Case 9
                            SalTemp = Me.AdoHistorialSalarial.Recordset("Septiembre")
                          Case 10
                            SalTemp = Me.AdoHistorialSalarial.Recordset("Octubre")
                          Case 11
                            SalTemp = Me.AdoHistorialSalarial.Recordset("Noviembre")
                          Case 12
                            SalTemp = Me.AdoHistorialSalarial.Recordset("Diciembre")
                         End Select
            
                 End If
            
            End If

         
         
         
           SalMayor = SalTemp + SalMayor
           
         
           iMes = CInt(rsDB.Fields("mes"))
         
           Do While Not rsDB.EOF
         
             If iMes <> CInt(rsDB.Fields("mes")) Then
                Exit Do
             End If
           
             rsDB.MoveNext
           Loop
         
         
        
         
        Loop
               
  '///////////////////////Busco si el Empleado tiene Subsidio para Sumarlo////////////////////
  
 
        Me.DtaConsulta.RecordSource = "SELECT NomSubsidio.NumNomina, NomSubsidio.FechaPago, DetalleNomSubsidio.CodEmpleado, DetalleNomSubsidio.Subsidio FROM NomSubsidio INNER JOIN DetalleNomSubsidio ON NomSubsidio.NumNomina = DetalleNomSubsidio.NumNominaSubsidio WHERE (((NomSubsidio.FechaPago) Between " & NumFecha1 & " And " & NumFecha2 & ") AND ((DetalleNomSubsidio.CodEmpleado)='" & CodEmpleado & "'))"
        Me.DtaConsulta.Refresh
        MontoSubsidio = 0
        
        Do While Not Me.DtaConsulta.Recordset.EOF
           MontoSubsidio = MontoSubsidio + Me.DtaConsulta.Recordset("Subsidio")
           Me.DtaConsulta.Recordset.MoveNext
        Loop
       
       
      '/////////////Busco el Adelanto de Vacaciones Registrados//////////////
'////////////////////////////////////////////////////////////////////////
        
       
         
        Me.DtaAdelanto.RecordSource = "SELECT  CodEmpleado, FechaAdelanto, MontoAdelanto, [Ref/Cheque], TipoAdelanto From Adelanto13vo WHERE     (FechaAdelanto BETWEEN CONVERT(DATETIME, '" & Format(Me.dtpFPInicio.Value, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.dtpFPFinal.Value, "yyyymmdd") & "', 102)) AND (TipoAdelanto = 'Vacaciones')AND (CodEmpleado = '" & CodEmpleado & "') "
'       Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo WHERE (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Adelanto13vo.FechaAdelanto) Between '" & Format(Me.dtpFPInicio.Value, "yyyy/mm/dd") & "' And '" & Format(Me.dtpFPFinal.Value, "yyyy/mm/dd") & "') AND ((Adelanto13vo.TipoAdelanto)='Vacaciones'))"
'        Me.DtaAdelanto.RecordSource = "SELECT CodEmpleado, FechaAdelanto, MontoAdelanto, [Ref/Cheque], TipoAdelanto From Adelanto13vo WHERE     (CodEmpleado = '" & CodEmpleado & "') AND (FechaAdelanto Between '" & Format(Me.dtpFPInicio.Value, "yyyymmdd") & "' And '" & Format(Me.dtpFPFinal.Value, "yyyymmdd") & "')) AND (TipoAdelanto = 'Vacaciones')"
          Me.DtaAdelanto.Refresh
          AdelantoVaca = 0
        
          Do While Not DtaAdelanto.Recordset.EOF
             AdelantoVaca = AdelantoVaca + DtaAdelanto.Recordset("MontoAdelanto")
             DtaAdelanto.Recordset.MoveNext
          Loop
       
         If val(SalMayor) < 0 Then
          SalMayor = 0
         End If
           
        If CantMeses <> 0 Then
'          If DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
            SalMayor = (SalMayor / CantMeses)
'          Else
'            SalMayor = SalMayor + MontoSubsidio
'          End If
       End If
          Año = Year(Me.TxtFFinVaca.Value)
 
 
 
 
 Else '///Si es salario fijo le calculo el ultimo salario ////////
 
'/////////////Busco el Adelanto de Vacaciones Registrados//////////////
'///////////////////////////////////////////////////////////////////////
'      cnDB.ConnectionString = Conexion
'      cnDB.Open

      
'      If Month(FechaHoy) <= 6 Then
        '/////////////CAlculo las vacaciones Enero - Junio ////////////////////////////
        SalMayor = 0
        CantMeses = 0
         
'         rsDB.Open "SELECT TipoNomina.Nomina, Fecha_Planilla.Periodo, Fecha_Planilla.año, Fecha_Planilla.mes, Fecha_Planilla.Inicio, Fecha_Planilla.Final, " & _
'          "Fecha_Planilla.Actual , Fecha_Planilla.NumNomina, Fecha_Planilla.Calculada " & _
'          "FROM Fecha_Planilla INNER JOIN TipoNomina ON Fecha_Planilla.CodTipoNomina = TipoNomina.CodTipoNomina " & _
'          "WHERE TipoNomina.Nomina LIKE '" & Me.DBCNominas.Text & "' AND Fecha_Planilla.Inicio >= CONVERT(DATETIME, '" & Format(Me.dtpFPInicio.Value, "yyyy-mm-dd") & " 00:00:00', 102) AND Fecha_Planilla.Final <= CONVERT(DATETIME, '" & Format(Me.dtpFPFinal.Value, "yyyy-mm-dd") & " 00:00:00', 102) ORDER BY Fecha_Planilla.Inicio ASC", cnDB
'
'
        
        
        Do While Not rsDB.EOF
         
           If iMes <> CInt(rsDB.Fields("mes")) Then
              CantMeses = CantMeses + 1
              iMes = CInt(rsDB.Fields("mes"))
           End If
           
           rsDB.MoveNext
         Loop





'        For Mes = 1 To 6
'         SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where ((Month([Nomina].[FechaNomina])) =  '" & Mes & "' ) And ((Year([Nomina].[FechaNomina])) ='" & Anno & "') and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
'            DtaNominas.RecordSource = SqlNominas
'            DtaNominas.Refresh
'            Fecha = "01/01/" & Anno
'         NumFecha1 = Fecha
         Fecha = "30/06/" & Anno
         NumFecha2 = Fecha
       ' Next
       
'      Else
'       For Mes = 7 To 12
'       '///////////////Busco las Vacaciones de Julio a Diciembre/////////////////////////////////////
'       'Julio- Diciembre
'        SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where ((Month([Nomina].[FechaNomina])) =  '" & Mes & "' ) And ((Year([Nomina].[FechaNomina])) ='" & Anno & "') and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
'            DtaNominas.RecordSource = SqlNominas
'            DtaNominas.Refresh
'               SalTemp = 0
'
'        Fecha = "01/07/" & Anno
'        NumFecha1 = Fecha
'        Fecha = "31/12/" & Anno
'        NumFecha2 = Fecha
'       Next
'
'      End If

         Me.DtaAdelanto.RecordSource = "SELECT  CodEmpleado, FechaAdelanto, MontoAdelanto, [Ref/Cheque], TipoAdelanto From Adelanto13vo WHERE     (FechaAdelanto BETWEEN CONVERT(DATETIME, '" & Format(Me.dtpFPInicio.Value, "yyyymmdd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.dtpFPFinal.Value, "yyyymmdd") & "', 102)) AND (TipoAdelanto = 'Vacaciones')AND (CodEmpleado = '" & CodEmpleado & "') "
'         Me.DtaAdelanto.RecordSource = "SELECT  CodEmpleado, FechaAdelanto, MontoAdelanto, [Ref/Cheque], TipoAdelanto From Adelanto13vo WHERE     (FechaAdelanto BETWEEN CONVERT(DATETIME, '2007-07-01 00:00:00', 102) AND CONVERT(DATETIME, '2007-12-31 00:00:00', 102)) AND (TipoAdelanto = 'Vacaciones')AND (CodEmpleado = '" & CodEmpleado & "') "
'        Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo WHERE (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Adelanto13vo.FechaAdelanto) Between  " & NumFecha1 & " And " & NumFecha2 & ") AND ((Adelanto13vo.TipoAdelanto)='Vacaciones'))"
        'Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo WHERE (((Adelanto13vo.CodEmpleado)='" & CodEmpleado & "') AND ((Adelanto13vo.FechaAdelanto) Between " & NumFecha1 & " And " & NumFecha2 & "))"
        Me.DtaAdelanto.Refresh
        AdelantoVaca = 0
        
        Do While Not DtaAdelanto.Recordset.EOF
         AdelantoVaca = AdelantoVaca + DtaAdelanto.Recordset("MontoAdelanto")
         DtaAdelanto.Recordset.MoveNext
        Loop
 
  
 
 
    SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]+[DetalleNomina].[IncetivoProduccion]+ +[DetalleNomina].[OtrosIngresos] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where (((DetalleNomina.CodEmpleado) = '" & CodEmpleado & "'))"
    DtaNominas.RecordSource = SqlNominas
    DtaNominas.Refresh
    If DtaNominas.Recordset.EOF Then
     Edicion = False
     'DtaNominas.Recordset.MoveLast
    End If
     
  '///////Selecciono el Salario Mayor de la Tabla Empleados/////////////////
     
     
                  Select Case DtaTipoNomina.Recordset("Periodo")
                        Case "Catorcenal los Sabados"
                        
                            TarifaHoraria = Format(DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio / 112, "###,##0.00")
                            
                        Case "Quincenal"
                            TarifaHoraria = Format(DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio / ((DiasMes * 8) / 2), "###,##0.00")
                            TotalHoras = 15 '* DtaTipoNomina.Recordset("Horas") '////LE ESCRIBO EL TOTAL DE DIAS ///
                        Case "Mensual"
                            TarifaHoraria = Format(DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio / (DiasMes * 8), "###,##0.00")
                        Case "Trimestral"
                            TarifaHoraria = Format(DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio / (DiasMes * 8 * 3), "###,##0.00")
                        Case "Semestral"
                            TarifaHoraria = Format(DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio / (DiasMes * 8 * 6), "###,##0.00")
                  End Select
    
    
'    SalMayor = DtaEmpleados.Recordset("SueldoPeriodo")
     SalMayor = Format(DiasMes * 8 * TarifaHoraria, "##,##0.00")
    
    'dependiendo del tipo de pago se hace el calculo del salario básico
    
    If DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
         SalMayor = SalMayor
    ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
          SalMayor = SalMayor
    ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
         SalMayor = SalMayor
    End If
    
    
    
    
    
    
   If Month(FechaHoy) <= 6 Then
      SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]+ DetalleNomina.IncetivoProduccion + [DetalleNomina].[OtrosIngresos] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina WHERE (((DetalleNomina.CodEmpleado)='" & CodEmpleado & "') AND ((Month([Nomina].[FechaNomina])) Between 1 And 6) AND ((Year([Nomina].[FechaNomina]))= " & Anno & " ))"
      DtaNominas.RecordSource = SqlNominas
      DtaNominas.Refresh
      If Not DtaNominas.Recordset.EOF Then
       DtaNominas.Recordset.MoveLast
      End If
      CantRegistros = 0
      
   Else
   
     SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos]+ [DetalleNomina].[IncetivoProduccion]+ [DetalleNomina].[OtrosIngresos] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina WHERE (((DetalleNomina.CodEmpleado)='" & CodEmpleado & "') AND ((Month([Nomina].[FechaNomina])) Between 7 And 12) AND ((Year([Nomina].[FechaNomina]))= " & Anno & " ))"
      DtaNominas.RecordSource = SqlNominas
      DtaNominas.Refresh
     If Not DtaNominas.Recordset.EOF Then
      DtaNominas.Recordset.MoveLast
     End If
      CantRegistros = 0
   End If

   
 End If  '/////////Fin del If salario fijo/////////////////
        
    'dependiendo del tipo de pago se hace el calculo del salario básico
    

    
    
     If DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
         If CantRegistros > 0 Then
           SalMayor = SalMayor / CantMeses
         Else
           SalMayor = SalMayor * 3
           CantRegistros = DtaNominas.Recordset.RecordCount
         End If
     ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
          If CantRegistros > 0 Then
           SalMayor = SalMayor / CantMeses
          Else
           SalMayor = SalMayor * 6
           CantRegistros = DtaNominas.Recordset.RecordCount
          End If
     ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
'         If CantRegistros > 0 Then
          If Dias < 182 Then
           'SalMayor = (Dias * (SalMayor * 2) / 30.4167) * 0.08333333
'          Else
'           SalMayor = 0
          End If
'         Else
'           SalMayor = SalMayor * 2
'           CantRegistros = DtaNominas.Recordset.RecordCount
'          End If
     ElseIf DtaTipoNomina.Recordset("Periodo") = "Semanal Viernes" Then
         If CantRegistros > 0 Then
          If CantMeses <> 0 Then
           SalMayor = (SalMayor / CantMeses)
          Else
           SalMayor = 0
          End If
         Else
           SalMayor = SalMayor * 2
           CantRegistros = DtaNominas.Recordset.RecordCount
          End If
     End If
     
 
DtaNomVaca.Refresh
Do While Not DtaNomVaca.Recordset.EOF
    If DtaNomVaca.Recordset("NumNomVaca") = val(TxtNumNomVaca.Text) And DtaNomVaca.Recordset("Activa") = True Then
       'DtaNomVaca.Recordset.Edit
       DtaNomVaca.Recordset("montopagado") = SalMayor + DtaNomVaca.Recordset("montopagado")
       'DtaNomVaca.Recordset.adelantovacaciones = AdelantoVaca
       DtaNomVaca.Recordset.Update
   Exit Do
    End If
DtaNomVaca.Recordset.MoveNext
Loop



                             If CodEmpleado1 = "S117080012" Then
                              CodEmpleado1 = "S117080012"
                             End If
                             


  If SueldoActual = False Then
    If SalarioBasico > SalMayor Then
      SalMayor = SalarioBasico
    End If
  Else
       '///////////////SI TIENE ACTIVO LA OPCION DE SALARIO, DEJO LO BASICO ///////////////
       SalMayor = SalarioBasico
  End If

        
        
                               '---------------------------------------------------------------------------------------------------------------------
                               '-----------------------------------BUSCO LOS SALARIOS DEL PERIODO DE VACACIONES --------------------------------------
                               '----------------------------------------------------------------------------------------------------------------------
                                NumNomina = val(Me.TxtNumNomVaca.Text)
                                Me.DtaConsulta.RecordSource = "SELECT  * From Reembolso WHERE (NumNomina = " & NumNomina & " ) AND (CodEmpleado = '" & CodEmpleado & "')"
                                Me.DtaConsulta.Refresh
                                If Not Me.DtaConsulta.Recordset.EOF Then
                                  Monto = Me.DtaConsulta.Recordset("Monto")
                                Else
                                  Monto = 0
                                End If
                               
                               SqlSalarios = "SELECT DISTINCT DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.HorasExtras) AS HorasExttras, SUM(DetalleNomina.BonoProduccion) AS BonoProduccion, SUM(DetalleNomina.SeptimoDia) AS SeptimoDia, SUM(DetalleNomina.OtrosIngresos) AS OtrosIngresos, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.Incentivos) AS Incentivos, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.Comisiones + DetalleNomina.HorasExtras + DetalleNomina.BonoProduccion + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos+DetalleNomina.IncetivoProduccion) AS TotalIngresos, MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin,Nomina.Mes AS MES,Nomina.Ano AS AÑO,SUM(DetalleNomina.Comisiones) AS Comisiones,SUM(DetalleNomina.MontoIR) AS MontoIR,SUM(DetalleNomina.MontoINSS) As MontoINSS  " & _
                                             "FROM  DetalleNomina INNER JOIN  Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = " & CodEmpleado & ") AND (MAX(Nomina.FechaNomina) <= CONVERT(DATETIME, '" & Format(Me.dtpFPFinal.Value, "yyyy-mm-dd") & "', 102)) AND (MIN(Nomina.FechaNominaINI) >= CONVERT(DATETIME, '" & Format(Me.dtpFPInicio.Value, "yyyy-mm-dd") & "', 102))"
                           
                               Me.AdoBusca.RecordSource = SqlSalarios
                               Me.AdoBusca.Refresh
                               If Not Me.AdoBusca.Recordset.EOF Then
                                 
                                 IrAcumulado = Format(Me.AdoBusca.Recordset("MontoIR"), "####0.00")
                                 InssAcumulado = Format(Me.AdoBusca.Recordset("MontoINSS"), "####0.00")
                                 SalarioAcumulado = Format(Me.AdoBusca.Recordset("TotalIngresos"), "####0.00") + ((SalMayor / DiasMes) * (Dias / 12)) + Monto - InssAcumulado - (((SalMayor / DiasMes) * (Dias / 12)) + Monto) * (TasaInss / 100)
                                 
                               Else
                                 IrAcumulado = 0
                                 InssAcumulado = 0
                                 SalarioAcumulado = 0
                               End If
                               
                               IR = CalcularIr(SalarioAcumulado, DtaTipoNomina.Recordset("Periodo"))
                               IrAcumulado = IR - IrAcumulado
                               If IrAcumulado < 0 Then
                                 IrAcumulado = 0
                               End If
                               
                                DiasDescuento = 0
        
                             If CodEmpleado = "39177" Then
                              CodEmpleado = "39177"
                             End If
                             

        

  If Edicion = True Then
   Me.DtaDetalleNomVaca.RecordSource = "SELECT DetalleNomVaca.Id, DetalleNomVaca.TotalDevengado, DetalleNomVaca.Inss, DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones From DetalleNomVaca Where (((DetalleNomVaca.NumNomVaca) = " & val(TxtNumNomVaca.Text) & ") And ((DetalleNomVaca.CodEmpleado) = '" & CodEmpleado & "'))"
   Me.DtaDetalleNomVaca.Refresh
     
   '///////Busco si el Empleado ya Existe en la Nomina de Vacaciones/////
     If Not Me.DtaDetalleNomVaca.Recordset.EOF Then
     
     DiasMenos = 0
     '/////////////////////////////////////////////////////////////////////////////////////////////////
     '///////////////////////BUSCO LOS DIAS DE VACACIONES /////////////////////////////////////////////
     '/////////////////////////////////////////////////////////////////////////////////////////////////
'      MDIPrimero.DtaConsulta.RecordSource = "SELECT CodigoEmpleado, SUM(DiasDisfrutar) AS Dias From SolicitudVacaciones WHERE (TipoSolicitud = 'Vacaciones') AND (Anulado = 0) AND (FechaInicio >= CONVERT(DATETIME, '" & Format(Me.TxtFINIVaca.Value, "yyyy-mm-dd") & "', 102)) AND (FechaFin <= CONVERT(DATETIME,'" & Format(Me.TxtFFinVaca.Value, "yyyy-MM-dd") & "', 102)) GROUP BY CodigoEmpleado, TipoSolicitud HAVING  (SolicitudVacaciones.CodigoEmpleado = '" & DtaEmpleados.Recordset("CodEmpleado1") & "')"
'      MDIPrimero.DtaConsulta.Refresh
'      If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
'          DiasMenos = MDIPrimero.DtaConsulta.Recordset("Dias")
'      Else
'          DiasMenos = 0
'      End If
      
      
      
       DiasDescuento = val(TxtDiasDescuento) + DiasMenos

        DtaDetalleNomVaca.Recordset("NumNomVaca") = val(TxtNumNomVaca.Text)
        DtaDetalleNomVaca.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
'        If Me.DBCNominas.Text = "Administracion" Then
        
       If DtaEmpleados.Recordset("SalarioFijo") = "S" Then
           
                 

            
            If DiasPagar = 15 Then
               DtaDetalleNomVaca.Recordset("Inss") = SalMayor * (TasaInss / 100)
               DtaDetalleNomVaca.Recordset("TotalDevengado") = SalMayor
            Else
            
              If DiasPagar = 0 Then
               DtaDetalleNomVaca.Recordset("Inss") = 0
               DtaDetalleNomVaca.Recordset("TotalDevengado") = 0
              Else
               DtaDetalleNomVaca.Recordset("Inss") = ((SalMayor / DiasMes) * (Dias - DiasDescuento - AdelantoVaca)) * (TasaInss / 100)
               DtaDetalleNomVaca.Recordset("TotalDevengado") = Dias * (SalMayor / DiasMes)
              End If
            End If
              
             DtaDetalleNomVaca.Recordset("SalarioMensual") = SalMayor
              
'          End If
          
       Else
            
            If Dias = 0 Then
                DtaDetalleNomVaca.Recordset("SalarioMensual") = SalMayor
                DtaDetalleNomVaca.Recordset("Inss") = 0
                DtaDetalleNomVaca.Recordset("TotalDevengado") = 0
            Else
                DtaDetalleNomVaca.Recordset("SalarioMensual") = SalMayor
                DtaDetalleNomVaca.Recordset("Inss") = ((SalMayor / DiasMes) * (Dias - DiasDescuento - AdelantoVaca)) * (TasaInss / 100)
                DtaDetalleNomVaca.Recordset("TotalDevengado") = (SalMayor / DiasMes) * (Dias - DiasDescuento)
            End If
            
        End If
        

        
        
     '///////////////////////////////////////////////////////////////////////////////////////////////////////
     '////////////////CALCULO CUANTOS DIAS SE CONSIDERAN PARA EL PAGO///////////////////////////////////////
     '//////SI ES MAYOR DE 15 REDONDEO A 15//////////////////////////////////////////////////////////////
     
       DiasPagar = Dias
       If DiasPagar > 15 Then
         DiasPagar = 15
       End If
       
       
    '////////////////////////////////SETEO LOS CALCULOS PARA QUE PODER MOSTRAR EL DETALLE DE DESCUENTO EN DIAS /
    '/////////////////////////POR QUE EN DIAS YA INCLUYE LA DEDUCCIO DE DIAS ///////////////////////////////////////
    DiasPagar = DiasNomVaca
    DiasMenos = DiasNomVaca - Dias
    
    
    If CodEmpleado1 = "S117090014" Then
      CodEmpleado1 = "S117090014"
    End If
       
       
     '////////////////////////////////////////////////////////////////////////////////////////////////////////
     '/////////////////////////////CALCULO DE LOS DIAS ACUMULADO EN OTRAS VACACIONES /////////////////////////
     '////////////////////////////////////////////////////////////////////////////////////////////////////////
     MDIPrimero.DtaConsulta.RecordSource = "SELECT  DetalleNomVaca.CodEmpleado, SUM(DetalleNomVaca.DiasAPagar) AS DiasAPagar, SUM(DetalleNomVaca.DiasDescuento) AS DiasDescuento, SUM(DetalleNomVaca.AdelantoVacaciones) As AdelantoVacaciones FROM  NomVaca INNER JOIN  DetalleNomVaca ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca Where (NomVaca.Activa = 0) GROUP BY DetalleNomVaca.CodEmpleado Having (DetalleNomVaca.CodEmpleado = " & CodEmpleado & ")"
     MDIPrimero.DtaConsulta.Refresh
     If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
'       DiasAcumulados = MDIPrimero.DtaConsulta.Recordset("DiasAPagar")
     Else
       DiasAcumulados = 0
     End If
     
     
     
     
'      '///////////////////SI LOS DIAS A PAGAR ES MAYOR QUE LOS QUE YA SE HAN PAGADO LOS HAGO CERO //////////////////////////
'     If DiasAcumulados > DiasPagar Then
'        DiasPagar = 0
'      Else
'        DiasPagar = DiasPagar - DiasAcumulados
'      End If
      
      '/////////////////////////////////COMPLEMENTO LOS DIAS A PAGAR PARA CUANDO SE HACE LA IMPORTACION DE FICHAS DE EXCEL //////////
      
'      If DiasNomVaca > DiasPagar Then
'        If DiasAcumulados = 0 Then
'          If DiasMenos = 0 Then
'           DiasMenos = DiasNomVaca - DiasPagar
'           DiasPagar = DiasNomVaca
'          End If
'        End If
'      End If
      
      
      
                              
        
       DtaDetalleNomVaca.Recordset("DiasAPagar") = DiasPagar
         DtaDetalleNomVaca.Recordset("AdelantoVacaciones") = AdelantoVaca
        If IsNull(DtaDetalleNomVaca.Recordset("DiasDescuento")) Then
            DtaDetalleNomVaca.Recordset("DiasDescuento") = 0
        Else
          DtaDetalleNomVaca.Recordset("DiasDescuento") = val(TxtDiasDescuento) + DiasMenos
        
        End If
        DtaDetalleNomVaca.Recordset.Update
     Else
       Me.DtaDetalleNomVaca.RecordSource = "SELECT DetalleNomVaca.Id, DetalleNomVaca.Inss, DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones From DetalleNomVaca"
       Me.DtaDetalleNomVaca.Refresh
       If Me.DtaDetalleNomVaca.Recordset.EOF Then
         Id = 1
       Else
         Me.DtaDetalleNomVaca.Recordset.MoveLast
         Id = Me.DtaDetalleNomVaca.Recordset("id") + 1
       End If
'       DiasDescuento = DtaDetalleNomVaca.Recordset("DiasDescuento") + Val(TxtDiasDescuento)
      If SalMayor > 0 Then
            DtaDetalleNomVaca.Recordset.AddNew
            DtaDetalleNomVaca.Recordset("id") = Id
            DtaDetalleNomVaca.Recordset("NumNomVaca") = val(TxtNumNomVaca.Text)
            DtaDetalleNomVaca.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
            DtaDetalleNomVaca.Recordset("SalarioMensual") = SalMayor
    '        DtaDetalleNomVaca.Recordset("DiasAPagar") = 1.25 * CantRegistros
            DtaDetalleNomVaca.Recordset("DiasAPagar") = DiasPagar
             'DtaDetalleNomVaca.Recordset("Ir") = IrAcumulado
            DtaDetalleNomVaca.Recordset("Inss") = ((SalMayor / DiasMes) * (Dias - DiasDescuento - AdelantoVaca)) * (TasaInss / 100)
            DtaDetalleNomVaca.Recordset("AdelantoVacaciones") = AdelantoVaca
            If IsNull(DtaDetalleNomVaca.Recordset("DiasDescuento")) Then
             DtaDetalleNomVaca.Recordset("DiasDescuento") = 0
            Else
            DtaDetalleNomVaca.Recordset("DiasDescuento") = DtaDetalleNomVaca.Recordset("DiasDescuento") + val(TxtDiasDescuento) + DiasMenos
            End If
            DtaDetalleNomVaca.Recordset.Update
       End If
     End If
   Else
   
       Me.DtaDetalleNomVaca.RecordSource = "SELECT DetalleNomVaca.Id, DetalleNomVaca.Inss, DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, DetalleNomVaca.ir From DetalleNomVaca"
       Me.DtaDetalleNomVaca.Refresh
       If Me.DtaDetalleNomVaca.Recordset.EOF Then
         Id = 1
       Else
         Me.DtaDetalleNomVaca.Recordset.MoveLast
         Id = Me.DtaDetalleNomVaca.Recordset("id") + 1
       End If
        DiasDescuento = DtaDetalleNomVaca.Recordset("DiasDescuento") + val(TxtDiasDescuento)
         If SalMayor > 0 Then
            DtaDetalleNomVaca.Recordset.AddNew
            DtaDetalleNomVaca.Recordset("id") = Id
            DtaDetalleNomVaca.Recordset("NumNomVaca") = val(TxtNumNomVaca.Text)
            DtaDetalleNomVaca.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
            DtaDetalleNomVaca.Recordset("SalarioMensual") = SalMayor
            DtaDetalleNomVaca.Recordset("DiasAPagar") = DiasPagar
            DtaDetalleNomVaca.Recordset("Ir") = IrAcumulado
            DtaDetalleNomVaca.Recordset("Inss") = ((SalMayor / DiasMes) * (Dias - DiasDescuento - AdelantoVaca)) * (TasaInss / 100)
            'DtaDetalleNomVaca.Recordset("DiasAPagar") = 2.5 * CantMeses
            DtaDetalleNomVaca.Recordset("AdelantoVacaciones") = AdelantoVaca
            'MsgBox DtaDetalleNomVaca.Recordset("DiasDescuento")
            If IsNull(DtaDetalleNomVaca.Recordset("DiasDescuento")) Then
                DtaDetalleNomVaca.Recordset("DiasDescuento") = 0
            Else
              DtaDetalleNomVaca.Recordset("DiasDescuento") = val(TxtDiasDescuento) + DiasMenos
            
            End If
            DtaDetalleNomVaca.Recordset.Update
        End If
 End If
 
 
 Else
    Me.DtaDetalleNomVaca.RecordSource = "SELECT DetalleNomVaca.Inss, DetalleNomVaca.Id, DetalleNomVaca.Inss, DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones From DetalleNomVaca Where (((DetalleNomVaca.NumNomVaca) = " & val(TxtNumNomVaca.Text) & ") And ((DetalleNomVaca.CodEmpleado) = '" & CodEmpleado & "'))"
    Me.DtaDetalleNomVaca.Refresh
    If Not Me.DtaDetalleNomVaca.Recordset.EOF Then
    DtaDetalleNomVaca.Recordset("DiasAPagar") = 0
     DtaDetalleNomVaca.Recordset("Inss") = 0
    DtaDetalleNomVaca.Recordset.Update
    End If

 End If
 
 
DtaEmpleados.Recordset.MoveNext
i = i + 1

rsDB.Close

Loop
End With
DtaVacaciones.Refresh
Me.DbgrVacaciones.Columns(0).Visible = False
Me.DbgrVacaciones.Columns(0).Locked = True
Me.DbgrVacaciones.Columns(1).Locked = True
Me.DbgrVacaciones.Columns(2).Locked = True
Me.DbgrVacaciones.Columns(3).Locked = True
Me.DbgrVacaciones.Columns(4).Locked = True
Me.DbgrVacaciones.Columns(5).Locked = True
Me.DbgrVacaciones.Columns(6).Locked = True
Me.DbgrVacaciones.Columns(7).Locked = True
Me.DbgrVacaciones.Columns(10).Locked = True
Me.DbgrVacaciones.Columns(6).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(7).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(8).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(9).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(10).NumberFormat = "##,##0.00"
Me.CmdPRVaca.Enabled = True
Me.CmdExportar.Visible = True
'Me.CmdMonedas.Visible = True
Me.CmdMonedasvaca.Enabled = True
Me.CmdNominaVaca.Enabled = True
Me.CmdColillaVaca.Enabled = True
MousePointer = 1

cnDB.Close

End Sub

Private Sub CmdCerrar13_Click()
DtaConsulta.RecordSource = "SELECT Nom13Mes.NumNom13Mes, Nom13Mes.Activa From Nom13Mes Where (((Nom13Mes.Activa) = 1))"
DtaConsulta.Refresh
Do While Not DtaConsulta.Recordset.EOF
If Me.DtaConsulta.Recordset("NumNom13Mes") = NumNom13Mes Then
   'DtaConsulta.Recordset.Edit
   DtaConsulta.Recordset("Activa") = False
   DtaConsulta.Recordset.Update
    Exit Do
End If
DtaConsulta.Recordset.MoveNext
Loop

DtaConsecutivos.Refresh
'DtaConsecutivos.Recordset.Edit
DtaConsecutivos.Recordset("nom13") = DtaConsecutivos.Recordset("nom13") + 1
DtaConsecutivos.Recordset.Update

MsgBox "La Nomina de 13vo Mes ha sido cerrada"
Unload Me
End Sub

Private Sub CmdCerrarVacaciones_Click()
On Error GoTo TipoErrs
Dim SqlEmpleados As String, CantEmpleados As Integer
Dim Respuesta As Integer, Cadena As String

If DBCNominas.Text = "Lista de Nóminas" Then
   MsgBox "No ha seleccionado el tipo de nomina al cual le desea calcular las Vacaciones"
   MousePointer = 1
   DBCNominas.SetFocus
   Exit Sub
End If

Respuesta = MsgBox("¿Esta Seguro de Cerrar la Nomina?", vbYesNo, "Sistema de Nominas")
If Respuesta = 6 Then

 If Me.CHKTranferir.Value = 0 Then
  '///////////////Verifico si existe Nomina Activa para Tranferirla///////////////////////
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
  Me.DtaNominas.RecordSource = "SELECT Nomina.NumNomina, Nomina.CodTipoNomina, Nomina.FechaNominaINI, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada From Nomina WHERE (((Nomina.CodTipoNomina)='" & CodTipoNomina & "') AND ((Nomina.Activa)=1))AND(Nomina.Procesada=1)"
  DtaNominas.Refresh
  If DtaNominas.Recordset.EOF Then
     Cadena = "No existe Ninguna Nomina, Activa" & vbLf
     Cadena = Cadena & "ó no se ha caluculado Nunca!!!!" & vbLf
     Cadena = Cadena & "Active la Nomina y luego Calcule"
    MsgBox Cadena, vbCritical, "Sistema de Nominas"
   Exit Sub
  End If
 End If

  DtaConsecutivos.Refresh
  'DtaConsecutivos.Recordset.Edit
'  DtaConsecutivos.Recordset("NomVaca") = DtaConsecutivos.Recordset("NomVaca") + 1
  DtaConsecutivos.Recordset.Update
 
  If Me.CHKTranferir.Value = 0 Then
    SqlEmpleados = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, Empleado.ExentoIr, Empleado.OtrosIngresos,  Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, Empleado.Ausente, Empleado.SalarioFijo From Empleado WHERE Empleado.CodTipoNomina= '" & CodTipoNomina & "' AND Empleado.Activo=1 AND Empleado.Ausente=0"
    DtaEmpleados.RecordSource = SqlEmpleados
    DtaEmpleados.Refresh

    DtaEmpleados.Recordset.MoveLast
    
    CantEmpleados = DtaEmpleados.Recordset.RecordCount

    With PBVacaciones
        .Min = 0
        .Max = CantEmpleados
        .Value = 0
         i = 1

  
        NumNomina = DtaNominas.Recordset("NumNomina")
'        Me.DtaVacaciones.Recordset.MoveFirst
        
        SqlVacaciones = "SELECT NomVaca.NumNomVaca, Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2,DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, DetalleNomVaca.SalarioMensual * (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento)/ '" & DiasMes & "' - DetalleNomVaca.AdelantoVacaciones AS MontoAPagar, NomVaca.CodTipoNomina FROM NomVaca INNER JOIN Empleado INNER JOIN       DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca WHERE (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1"
        Me.DtaConsulta.RecordSource = SqlVacaciones
        Me.DtaConsulta.Refresh
        Do While Not DtaConsulta.Recordset.EOF
            CodEmpleado = DtaConsulta.Recordset("CodEmpleado")
            If i > CantEmpleados Then
                Exit Do
            End If
            .Value = i
            Me.Caption = "Procesando:  " & i & " de " & CantEmpleados & " Empleados "
            DoEvents
        '/////////////Agrego el Monto de Vacaciones Ganadas a la Nomina Activa////////////////////////////////////////////////
            Me.DtaDetalleNominas.RecordSource = "SELECT DetalleNomina.AdelantosVacaciones,DetalleNomina.DiasVacaciones,DetalleNomina.NumNomina, DetalleNomina.CodEmpleado, DetalleNomina.VacacionesPagadas From DetalleNomina Where (((DetalleNomina.NumNomina) = " & NumNomina & ") And ((DetalleNomina.CodEmpleado) = '" & CodEmpleado & "'))"
            Me.DtaDetalleNominas.Refresh
            If Not DtaDetalleNominas.Recordset.EOF Then
                'DtaDetalleNominas.Recordset.Edit
                DtaDetalleNominas.Recordset("DiasVacaciones") = DtaConsulta.Recordset("DiasAPagar") - DtaConsulta.Recordset("DiasDescuento")
                DtaDetalleNominas.Recordset("VacacionesPagadas") = Me.DtaConsulta.Recordset("MontoAPagar")
                DtaDetalleNominas.Recordset("AdelantosVacaciones") = Me.DtaConsulta.Recordset("AdelantoVacaciones")
                DtaDetalleNominas.Recordset.Update
            End If
             DtaConsulta.Recordset.MoveNext
             i = i + 1
        Loop
  
    End With
  End If
  
  DtaNomVaca.Refresh
  Do While Not DtaNomVaca.Recordset.EOF
  If DtaNomVaca.Recordset("NumNomVaca") = val(TxtNumNomVaca.Text) Then
    'DtaNomVaca.Recordset.Edit
    DtaNomVaca.Recordset("Activa") = 0
    DtaNomVaca.Recordset.Update
   End If
   DtaNomVaca.Recordset.MoveNext
   Loop




   MsgBox "La Nomina de Vacaciones ha sido cerrada"
   Unload Me
End If
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdPr13mes_Click()


End Sub

Private Sub ÇmdColillaVaca_Click()

End Sub

Private Sub CmdColillaVaca_Click()
On Error GoTo TipoErrs
Dim Espacio As String
Dim NumNomina As Integer
Espacio = " "
NumNomina = Me.TxtNumNomVaca.Text
Quien = "Colilla de Pago Vacaciones"
 
DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop

If Me.ChkImprimirDptoVaca.Value = 1 Then

    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////////CONSULTO LOS DEPARTAMENTOS PARA EL TIPO DE NOMINA ////////////////////////
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Me.AdoBusca.RecordSource = "SELECT DISTINCT Departamento.Departamento, Departamento.CodDepartamento FROM NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN  Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento  WHERE (NomVaca.NumNomVaca = " & NumNomina & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Departamento.CodDepartamento"
    Me.AdoBusca.Refresh
    Do While Not Me.AdoBusca.Recordset.EOF
        
        Quien = "Colilla de Pago Vacaciones"
        NumNomina = Me.TxtNumNom13.Text
        
        CodDepartamento = Me.AdoBusca.Recordset("CodDepartamento")
        NombreDepartamento = Me.AdoBusca.Recordset("Departamento")
        
        MsgBox "SE IMPRIMIRA DEPARTAMENTO: " & NombreDepartamento, vbInformation, "Zeus Nominas"

             If Me.DBCNominas.Text <> "Administracion" Then
             
                      If Me.ChkRestar.Value = 1 Then
'                           SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina,DetalleNomVaca.Ir FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
'                                         "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (Empleado.CodDepartamento = '" & CodDepartamento & "') ORDER BY Empleado.CodEmpleado1 "
                            SQlReportes = "SELECT  NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo, DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoPagar, DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina, DetalleNomVaca.IR , departamento.CodDepartamento, departamento.departamento FROM  NomVaca INNER JOIN Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN  Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN " & _
                                          "Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento WHERE (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (Empleado.CodDepartamento = '" & CodDepartamento & "') AND (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento > 0) ORDER BY Empleado.CodEmpleado1"
                     Else
'                          SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado,Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina, DetalleNomVaca.Ir FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
'                                        "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (Empleado.CodDepartamento = '" & CodDepartamento & "') ORDER BY Empleado.CodEmpleado1 "
                           SQlReportes = "SELECT  NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo, DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar, DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina, DetalleNomVaca.Ir, departamento.CodDepartamento , departamento.departamento FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN  " & _
                                         "Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento WHERE (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (Empleado.CodDepartamento = '" & CodDepartamento & "') AND (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento > 0) ORDER BY Empleado.CodEmpleado1"
                     
                     End If
                     
                     
                        ArepColillaVaca.AdoColillas.Source = SQlReportes
                        ArepColillaVaca.LblTipo.Caption = Me.DBCNominas.Text
                        ArepColillaVaca.LblPeriodos.Caption = "Desde   " & Me.TxtFINIVaca.Value & " Hasta    " & Me.TxtFFinVaca.Value
                        ArepColillaVaca.LblPeriodo.Caption = Format(Me.TxtFINIVaca.Value, "dd/mm/yyyy") & "   Hasta   " & Format(Me.TxtFFinVaca.Value, "dd/mm/yyyy")
                        ArepColillaVaca.lbltitulo.Caption = Titulo
                        ArepColillaVaca.AdoColillas.ConnectionString = ConexionReporte
                        ArepColillaVaca.Show 1
            
            
            
             Else
            
                    If Me.ChkRestar.Value = 1 Then
                        SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina, DetalleNomVaca.Ir FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                                   "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) AND (Empleado.CodDepartamento = '" & CodDepartamento & "') AND (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento > 0) ORDER BY Empleado.CodEmpleado1 "
                    
                    Else
                          SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina, DetalleNomVaca.Ir FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                                        "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) AND (Empleado.CodDepartamento = '" & CodDepartamento & "') AND (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento > 0) ORDER BY Empleado.CodEmpleado1 "
                    End If
            
            
            
             
                    ArepColillaVaca.AdoColillas.Source = SQlReportes
                    ArepColillaVaca.LblTipo.Caption = Me.DBCNominas.Text
                    ArepColillaVaca.LblPeriodos.Caption = "Desde   " & Me.TxtFINIVaca.Value & " Hasta    " & Me.TxtFFinVaca.Value
                    ArepColillaVaca.LblPeriodo.Caption = Format(Me.TxtFINIVaca.Value, "dd/mm/yyyy") & "   Hasta   " & Format(Me.TxtFFinVaca.Value, "dd/mm/yyyy")
                    ArepColillaVaca.lbltitulo.Caption = Titulo
                    ArepColillaVaca.AdoColillas.ConnectionString = ConexionReporte
                    ArepColillaVaca.Show 1
            
            
                    Exit Sub
            
            
            
             End If

      Me.AdoBusca.Recordset.MoveNext
    Loop

Else

         If Me.DBCNominas.Text <> "Administracion" Then
         
'                  If Me.ChkRestar.Value = 1 Then
'                       SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina,DetalleNomVaca.Ir FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
'                                     "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "
'                 Else
'                      SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado,Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina, DetalleNomVaca.Ir FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
'                                    "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "
'                 End If
                 
                    If Me.ChkRestar.Value = 1 Then
'                           SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina,DetalleNomVaca.Ir FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
'                                         "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (Empleado.CodDepartamento = '" & CodDepartamento & "') ORDER BY Empleado.CodEmpleado1 "
                            SQlReportes = "SELECT  NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo, DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoPagar, DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina, DetalleNomVaca.IR , departamento.CodDepartamento, departamento.departamento FROM  NomVaca INNER JOIN Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN  Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN " & _
                                          "Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento WHERE (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1"
                     Else
'                          SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado,Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina, DetalleNomVaca.Ir FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
'                                        "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (Empleado.CodDepartamento = '" & CodDepartamento & "') ORDER BY Empleado.CodEmpleado1 "
                           SQlReportes = "SELECT  NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo, DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar, DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina, DetalleNomVaca.Ir, departamento.CodDepartamento , departamento.departamento FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN  " & _
                                         "Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento WHERE (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1"
                     
                     End If
                 
                 
                    ArepColillaVaca.AdoColillas.Source = SQlReportes
                    ArepColillaVaca.LblTipo.Caption = Me.DBCNominas.Text
                    ArepColillaVaca.LblPeriodos.Caption = "Desde   " & Me.TxtFINIVaca.Value & " Hasta    " & Me.TxtFFinVaca.Value
                    ArepColillaVaca.LblPeriodo.Caption = Format(Me.TxtFINIVaca.Value, "dd/mm/yyyy") & "   Hasta   " & Format(Me.TxtFFinVaca.Value, "dd/mm/yyyy")
                    ArepColillaVaca.lbltitulo.Caption = Titulo
                    ArepColillaVaca.AdoColillas.ConnectionString = ConexionReporte
                    ArepColillaVaca.Show 1
        
        
        
         Else
        
                If Me.ChkRestar.Value = 1 Then
                    SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina, DetalleNomVaca.Ir FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                               "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
                
                Else
                      SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina, DetalleNomVaca.Ir FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                                    "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
                End If
        
        
        
         
                ArepColillaVaca.AdoColillas.Source = SQlReportes
                ArepColillaVaca.LblTipo.Caption = Me.DBCNominas.Text
                ArepColillaVaca.LblPeriodos.Caption = "Desde   " & Me.TxtFINIVaca.Value & " Hasta    " & Me.TxtFFinVaca.Value
                ArepColillaVaca.LblPeriodo.Caption = Format(Me.TxtFINIVaca.Value, "dd/mm/yyyy") & "   Hasta   " & Format(Me.TxtFFinVaca.Value, "dd/mm/yyyy")
                ArepColillaVaca.lbltitulo.Caption = Titulo
                ArepColillaVaca.AdoColillas.ConnectionString = ConexionReporte
                ArepColillaVaca.Show 1
        
        
                Exit Sub
        
        
        
         End If


End If

Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdDenominacion_Click()
On Error GoTo TipoErr
DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop

'SQLNomina = "SELECT Nomina.* From Nomina WHERE Nomina.Activa=1 AND Nomina.CodTipoNomina= '" & CodTipoNomina & "'"
'DtaNomina.RecordSource = SQLNomina
'DtaNomina.Refresh

NumNomina = Me.TxtNumNom13.Text
Quien = "13vo"
FrmMonedas13vo.Show 1
Exit Sub
TipoErr:
    ControlErrores
End Sub

Private Sub CmdExporta2_Click()
On Error GoTo TipoErrs
Dim Espacio As String
Espacio = " "
NumNomina = Me.TxtNumNom13.Text
Me.CommonDialog1.ShowSave
Directorio = ""
Directorio = Me.CommonDialog1.FileName + ".xls"

      Nom13vo.lbltitulo.Caption = Titulo
      Nom13vo.LblSubtitulo.Caption = SubTitulo
      Nom13vo.ImgLogo.Picture = LoadPicture(RutaLogo)
      
      Nom13vo.LblFecha.Caption = "Desde " + Format(Me.TxtFINI13.Value, "mm/dd/yyyy") + " Hasta " + Format(Me.TxtFFIN13.Value, "mm/dd/yyyy")
      Nom13vo.LblFechaHoy = Format(Now, "dddddd")
      Nom13vo.DataControl1.ConnectionString = ConexionReporte
'      SQLReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, ([DetalleNom13Mes].[SalarioMensual]) AS MontoPagar, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, Cargo.Cargo, ([DetalleNom13Mes].[SalarioMensual]) AS TotalDevengado, Empleado.CodEmpleado1 FROM Nom13Mes INNER JOIN ((Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNomina & ")) ORDER BY Empleado.CodEmpleado1"

      SQlReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, (DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo) AS MontoPagar, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, Cargo.Cargo, ([DetalleNom13Mes].SalarioAPagar) AS TotalDevengado, Empleado.CodEmpleado1 FROM Nom13Mes INNER JOIN ((Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNomina & ")) ORDER BY Nombres"

      Nom13vo.DataControl1.Source = SQlReportes
      Nom13vo.ImgLogo.Picture = LoadPicture(RutaLogo)
      Exportar = True
      Nom13vo.Show 1

Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub CmdMonedas_Click()
On Error GoTo TipoErr
DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop

'SQLNomina = "SELECT Nomina.* From Nomina WHERE Nomina.Activa=1 AND Nomina.CodTipoNomina= '" & CodTipoNomina & "'"
'DtaNomina.RecordSource = SQLNomina
'DtaNomina.Refresh

NumNomina = Me.TxtNumNomVaca.Text
FrmMonedas13vo.Show 1
Exit Sub
TipoErr:
    ControlErrores
End Sub

Private Sub CmdExportaBAC_Click()
Quien = "Nomina13vo"
FrmExportaBac.Show 1
End Sub

Private Sub CmdExportaBanpro_Click()
On Error GoTo TipoErrs
Dim SQlReportes As String, V As Integer, H As Integer, i As Integer
Dim Año As String, MesLetra As String, Neto As String, Dias As String
Dim CanDias As String, QuinLetra As String, Nombres As String, Espacio As String
Dim TotalNomina As Double, Neto1 As Double, Cod As String, NetoT As String, Longitud As Integer
Dim CodigoCuenta As String, NombreEmpresa As String, MontoSubsidio As Double

Espacio = " "
Quien = "CalcularNomina"
Select Case Quien
 Case "CalcularNomina"
       '//////////////////////Cargo la Consulta de la Nomina///////////////////////
  
   NumNomina = Me.TxtNumNom13.Text
   SQlReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, (DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo) AS MontoPagar, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, Cargo.Cargo, ([DetalleNom13Mes].SalarioAPagar) AS TotalDevengado, Empleado.CodEmpleado1, Empleado.CuentaBanco FROM Nom13Mes INNER JOIN ((Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNomina & ")) ORDER BY Empleado.CodEmpleado1"

       Me.DtaConsulta.RecordSource = SQlReportes
       Me.DtaConsulta.Refresh

    
'    NumFecha1 = CDate(Me.TxtFINI13)
'    NumFecha2 = CDate(Me.TxtFFIN13)
'    Sql13voMes = "SELECT Nom13Mes.NumNom13Mes, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, DetalleNom13Mes.SalarioAPagar-DetalleNom13Mes.Adelanto13vo AS MontoPagar FROM Nom13Mes INNER JOIN (Empleado INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNomina & ")) ORDER BY Empleado.CodEmpleado1"
'
'    Dta13voMes.RecordSource = Sql13voMes
'    Dta13voMes.Refresh

  
   Case "NominaVacaciones"
      NumNomVaca = Frm13Vaca.TxtNumNomVaca.Text
      '///////////////////////////Cargo la Consulta de Vacaciones////////////////////////////////
      SQlReportes = "SELECT NomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, ([DetalleNomVaca].[SalarioMensual]*([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento])/" & DiasMes & ")-[DetalleNomVaca].[AdelantoVacaciones] AS MontoAPagar, [DetalleNomVaca].[SalarioMensual]*([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento])/" & DiasMes & " AS TotalDevengado, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, ([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento]) AS TotalDescuento " & vbLf
      SQlReportes = SQlReportes & "FROM NomVaca INNER JOIN (Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado) ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca Where (((NomVaca.NumNomVaca) = " & NumNomVaca & " )) ORDER BY Nombres"
'       Me.DtaExporta.Refresh
'        'Me.'Me.DtaExporta.Recordset.Edit
'       Me.DtaExporta.Recordset("CodigoBAC") = val(Me.TxtCod.Text)
'       Me.DtaExporta.Recordset.Update

       Mes = Month(Me.DtaConsulta.Recordset("FechaNomina"))
       Año = Year(Me.DtaConsulta.Recordset("FechaNomina"))
       CanDias = Day(Me.DtaConsulta.Recordset("FechaNomina"))
       Dias = Day(Me.DtaConsulta.Recordset("FechaNomina"))
'       Cod = Me.TxtCod.Text
End Select

            
   
    Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
    'Call Formato_Excel(8, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
V = 1
H = 0
i = 1

 
  Do While Not Me.DtaConsulta.Recordset.EOF 'esto nos sirve pa leer los datos desde
       
       CodEmpleado = DtaConsulta.Recordset("CodEmpleado")
       
       MontoSubsidio = 0
       
       If Not IsNull(DtaConsulta.Recordset("CuentaBanco")) Then
         CodigoCuenta = DtaConsulta.Recordset("CuentaBanco")
       Else
         CodigoCuenta = ""
       End If
 'la tabla de access para despues colocarlos en las celdas correspondientes
       
       Nombre = Me.DtaConsulta.Recordset("Nombres")
       Neto = Format(Me.DtaConsulta.Recordset("MontoPagar") + MontoSubsidio, "####0.00")
       Neto1 = Format(Me.DtaConsulta.Recordset("MontoPagar") + MontoSubsidio, "##,##0.00")
       Longitud = Len(Neto)
       NetoT = Mid(Neto, Longitud - 1, 3)
       NetoT = (Mid(Neto, 1, Longitud - 3)) & NetoT
       With DtaConsulta.Recordset

       
'           If Not (V = 1) Then
'             objExcel.ActiveSheet.Cells(V, H) = "T"
'           End If
            objExcel.ActiveSheet.Cells(V, H + 1) = Nombre
            objExcel.ActiveSheet.Cells(V, H + 2) = CodigoCuenta
            objExcel.ActiveSheet.Cells(V, H + 3) = "13vo Mes"
            objExcel.ActiveSheet.Cells(V, H + 4) = Format(Neto, "##,##0.00")
            objExcel.ActiveSheet.Cells(V, H + 5) = "C"
            V = V + 1
            i = i + 1
            TotalNomina = TotalNomina + Neto1
            .MoveNext

   
        End With
  Loop
  
  '/////////////////////////////SELECCION SOLO LOS EMPLEADOS QUE TIENEN SUBSIDIO Y NO TIENEN SALARIO
  Me.DtaConsulta.RecordSource = "SELECT TOP (200) DetalleNomSubsidio.id, DetalleNomSubsidio.NumNominaSubsidio, DetalleNomSubsidio.CodEmpleado, DetalleNomSubsidio.Subsidio, Empleado.CodEmpleado1, (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS Neto, Empleado.CuentaBanco, Empleado.Dolarizado, Empleado.FechaAntiguedad, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres " & _
                                "FROM DetalleNomSubsidio INNER JOIN Empleado ON DetalleNomSubsidio.CodEmpleado = Empleado.CodEmpleado INNER JOIN Nomina ON DetalleNomSubsidio.NumNominaSubsidio = Nomina.NumNomina INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado AND Nomina.NumNomina = DetalleNomina.NumNomina  " & _
                                "WHERE (DetalleNomSubsidio.NumNominaSubsidio = " & NumNomina & ") AND ((DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) = 0) AND (DetalleNomSubsidio.Subsidio <> 0) Order by Nombres"
  Me.DtaConsulta.Refresh
  Do While Not Me.DtaConsulta.Recordset.EOF
       If Not IsNull(DtaConsulta.Recordset("CuentaBanco")) Then
         CodigoCuenta = DtaConsulta.Recordset("CuentaBanco")
       Else
         CodigoCuenta = ""
       End If
 'la tabla de access para despues colocarlos en las celdas correspondientes
       
       Nombre = Me.DtaConsulta.Recordset("Nombres")
       Neto = Format(Me.DtaConsulta.Recordset("Subsidio"), "####0.00")
       Neto1 = Format(Me.DtaConsulta.Recordset("Subsidio"), "##,##0.00")
       Longitud = Len(Neto)
       NetoT = Mid(Neto, Longitud - 1, 3)
       NetoT = (Mid(Neto, 1, Longitud - 3)) & NetoT
       With DtaConsulta.Recordset

       
'           If Not (V = 1) Then
'             objExcel.ActiveSheet.Cells(V, H) = "T"
'           End If
            objExcel.ActiveSheet.Cells(V, H + 1) = Nombre
            objExcel.ActiveSheet.Cells(V, H + 2) = CodigoCuenta
            objExcel.ActiveSheet.Cells(V, H + 3) = QuinLetra
            objExcel.ActiveSheet.Cells(V, H + 4) = Format(Neto, "##,##0.00")
            objExcel.ActiveSheet.Cells(V, H + 5) = "C"
            V = V + 1
            i = i + 1
            TotalNomina = TotalNomina + Neto1
            

   
        End With
     Me.DtaConsulta.Recordset.MoveNext
  Loop
  
  
     
       MDIPrimero.DtaEmpresa.Refresh
       If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("NombreEmpresa")) Then
         NombreEmpresa = MDIPrimero.DtaEmpresa.Recordset("NombreEmpresa")
       End If
       Neto = Format(TotalNomina, "####0.00")
       Longitud = Len(Neto)
       NetoT = Mid(Neto, Longitud - 1, 3)
       NetoT = (Mid(Neto, 1, Longitud - 3)) & NetoT
   

       objExcel.ActiveSheet.Cells(V, 1) = NombreEmpresa
       objExcel.ActiveSheet.Cells(V, 2) = "10013208274380"
       objExcel.ActiveSheet.Cells(V, 3) = QuinLetra
       objExcel.ActiveSheet.Cells(V, 4) = Format(Neto, "##,##0.00")
       objExcel.ActiveSheet.Cells(V, 5) = "D"
       objExcel.ActiveSheet.Cells(V, 1).Font.Bold = True
       objExcel.ActiveSheet.Cells(V, 2).Font.Bold = True
       objExcel.ActiveSheet.Cells(V, 3).Font.Bold = True
       objExcel.ActiveSheet.Cells(V, 4).Font.Bold = True
       objExcel.ActiveSheet.Cells(V, 5).Font.Bold = True
       
        objExcel.ActiveSheet.Columns("A").ColumnWidth = 35
        objExcel.ActiveSheet.Columns("A").Font.Size = 10
        objExcel.ActiveSheet.Columns("B").NumberFormat = "############"
        objExcel.ActiveSheet.Columns("B").ColumnWidth = 17
        objExcel.ActiveSheet.Columns("B").Font.Size = 10
        objExcel.ActiveSheet.Columns("B").HorizontalAlignment = xlHAlignCenter
        objExcel.ActiveSheet.Columns("C").ColumnWidth = 26
        objExcel.ActiveSheet.Columns("C").Font.Size = 10
        objExcel.ActiveSheet.Columns("C").HorizontalAlignment = xlHAlignCenter
        objExcel.ActiveSheet.Columns("D").ColumnWidth = 10
        objExcel.ActiveSheet.Columns("D").Font.Size = 10
        objExcel.ActiveSheet.Columns("D").HorizontalAlignment = xlHAlignCenter
        objExcel.ActiveSheet.Columns("E").ColumnWidth = 4
        objExcel.ActiveSheet.Columns("E").Font.Size = 10
        objExcel.ActiveSheet.Columns("E").HorizontalAlignment = xlHAlignCenter

 
 Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto

Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdExportar_Click()
On Error GoTo TipoErrs
Dim Espacio As String

DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop

Espacio = " "
NumNomina = Me.TxtNumNomVaca.Text
Espacio = " "
NumNomina = Me.TxtNumNom13.Text
Directorio = ""
Me.CommonDialog1.ShowSave

Directorio = Me.CommonDialog1.FileName + ".xls"

      Nom13vo.lbltitulo.Caption = Titulo
      Nom13vo.LblSubtitulo.Caption = SubTitulo
      Nom13vo.ImgLogo.Picture = LoadPicture(RutaLogo)
      
      Nom13vo.LblFecha.Caption = "Desde " + Format(Me.TxtFINI13.Value, "mm/dd/yyyy") + " Hasta " + Format(Me.TxtFFIN13.Value, "mm/dd/yyyy")
      Nom13vo.LblFechaHoy = Format(Now, "dddddd")
      Nom13vo.DataControl1.ConnectionString = ConexionReporte
      
If Me.DBCNominas.Text <> "Administracion" Then
      
 If Me.ChkRestar.Value = 1 Then
       SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                     "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
 Else
      SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado,Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                    "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
 End If
Else

If Me.ChkRestar.Value = 1 Then
    SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
               "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "

Else
      SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                    "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "
End If
'Nom13voMoises.DataControl1.Source = SQLReportes
'Nom13voMoises.ImgLogo.Picture = LoadPicture(RutaLogo)
'Nom13voMoises.Show 1
'Exit Sub

End If
      Nom13vo.Label5.Caption = "Nomina de Vacaciones"
      Nom13vo.DataControl1.Source = SQlReportes
      Nom13vo.ImgLogo.Picture = LoadPicture(RutaLogo)
      Exportar = True
      Nom13vo.Show 1

Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdMonedasvaca_Click()
Quien = "Vacaciones"
FrmMonedas13vo.Show 1
End Sub

Private Sub CmdNominaVaca_Click()
'On Error GoTo TipoErrs
'Dim Espacio As String
'Espacio = " "
'NumNomina = Me.TxtNumNomVaca.Text
'
'DtaTipoNomina.Refresh
'Do While Not DtaTipoNomina.Recordset.EOF
'If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
'   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
'   Exit Do
'End If
'DtaTipoNomina.Recordset.MoveNext
'Loop
'
'
'      Nom13vo.LblTitulo.Caption = Titulo
'      Nom13vo.LblSubtitulo.Caption = SubTitulo
'      Nom13vo.ImgLogo.Picture = LoadPicture(RutaLogo)
'
'      Nom13vo.lblFecha.Caption = "Desde " + Format(Me.TxtFINIVaca.Value, "mm/dd/yyyy") + " Hasta " + Format(Me.TxtFFinVaca.Value, "mm/dd/yyyy")
'      Nom13vo.LblFechaHoy = Format(Now, "dddddd")
'      Nom13vo.DataControl1.ConnectionString = ConexionReporte
'
'If Me.DBCNominas.Text <> "Administracion" Then
'
' If Me.ChkRestar.Value = 1 Then
'       SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
'                     "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
' Else
'      SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado,Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
'                    "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
' End If
'Else
'
'If Me.ChkRestar.Value = 1 Then
'    SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
'               "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "
'
'Else
'      SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
'                    "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "
'End If
''Nom13voMoises.DataControl1.Source = SQLReportes
''Nom13voMoises.ImgLogo.Picture = LoadPicture(RutaLogo)
''Nom13voMoises.Show 1
''Exit Sub
'
'End If
'
'      Nom13vo.Label5.Caption = "Nomina de Vacaciones"
'      Nom13vo.DataControl1.Source = SQlReportes
'      Nom13vo.ImgLogo.Picture = LoadPicture(RutaLogo)
'      Nom13vo.Show 1
'
'Exit Sub
'TipoErrs:
'ControlErrores

On Error GoTo TipoErrs
Dim Espacio As String
Dim rpt As Object
Dim fPreview As New FrmPreview

Espacio = " "
NumNomina = Me.TxtNumNomVaca.Text

DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop


      ArepNomVacaciones.lbltitulo.Caption = Titulo
      ArepNomVacaciones.LblSubtitulo.Caption = SubTitulo
      ArepNomVacaciones.ImgLogo.Picture = LoadPicture(RutaLogo)
      ArepNomVacaciones.LblFecha.Caption = "Desde " + Format(Me.TxtFINIVaca.Value, "mm/dd/yyyy") + " Hasta " + Format(Me.TxtFFinVaca.Value, "mm/dd/yyyy")
      ArepNomVacaciones.LblFechaHoy = Format(Now, "dddddd")
      ArepNomVacaciones.DataControl1.ConnectionString = ConexionReporte
      
If Me.DBCNominas.Text <> "Administracion" Then
      
     If Me.ChkRestar.Value = 1 Then
    '       SQlReportes = "SELECT Empleado.CodEmpleado,NomVaca.NumNomVaca , DetalleNomVaca.Inss, DetalleNomVaca.Ir, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
    '                     "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
            SQlReportes = "SELECT Empleado.CodEmpleado, NomVaca.NumNomVaca, DetalleNomVaca.Inss, DetalleNomVaca.Ir, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo, DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir AS MontoPagar, DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir AS TotalDeducir, NomVaca.CodTipoNomina, Departamento.Departamento, departamento.CodDepartamento FROM NomVaca INNER JOIN Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN  Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN " & _
                        "Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento WHERE  (Empleado.Activo = 1) AND (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (DetalleNomVaca.TotalDevengado + DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss+ DetalleNomVaca.Ir > 0)  AND (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento > 0) ORDER BY Departamento.CodDepartamento, Empleado.CodEmpleado1"
     Else
    '      SQlReportes = "SELECT Empleado.CodEmpleado,NomVaca.NumNomVaca , DetalleNomVaca.Inss, DetalleNomVaca.Ir,Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado,Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
    '                    "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
           SQlReportes = "SELECT  Empleado.CodEmpleado, NomVaca.NumNomVaca, DetalleNomVaca.Inss, DetalleNomVaca.Ir, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo, DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar, DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina, Departamento.CodDepartamento AS Expr1, Departamento.* FROM  NomVaca INNER JOIN  Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN " & _
                         "Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento WHERE (Empleado.Activo = 1) AND (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (DetalleNomVaca.TotalDevengado + DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir > 0) AND (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento > 0) ORDER BY Expr1, Empleado.CodEmpleado1"
     End If
Else
        
        If Me.ChkRestar.Value = 1 Then
            SQlReportes = "SELECT Empleado.CodEmpleado,NomVaca.NumNomVaca , DetalleNomVaca.Inss, DetalleNomVaca.Ir,Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                       "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.TotalDevengado + DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir > 0) AND (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento > 0) ORDER BY Empleado.CodEmpleado1 "
        
        Else
              SQlReportes = "SELECT Empleado.CodEmpleado,NomVaca.NumNomVaca , DetalleNomVaca.Inss, DetalleNomVaca.Ir,Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                            "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.TotalDevengado + DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir > 0) AND (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento > 0)  ORDER BY Empleado.CodEmpleado1 "
        End If


End If
        
      ArepNomVacaciones.Label5.Caption = "Nomina de Vacaciones"
      ArepNomVacaciones.DataControl1.Source = SQlReportes
      ArepNomVacaciones.ImgLogo.Picture = LoadPicture(RutaLogo)
      FechaIniVaca = Me.dtpFPInicio.Value
      FechaFinVaca = Me.dtpFPFinal.Value
'      Nom13vo.Show 1
'        Dim rpt As Object
'        Dim fPreview As New FrmPreview
        
             Set rpt = New ArepNomVacaciones
             rpt.DataControl1.ConnectionString = ConexionReporte
             rpt.DataControl1.Source = SQlReportes
             fPreview.RunReport rpt
        
        
             fPreview.Show 1
           
           
      ArepPersecionesVaca.lbltitulo.Caption = Titulo
      ArepPersecionesVaca.LblSubtitulo.Caption = SubTitulo
      ArepPersecionesVaca.ImgLogo.Picture = LoadPicture(RutaLogo)
      ArepPersecionesVaca.Label5.Caption = "Nomina Vacaciones"
      
      ArepPersecionesVaca.LblFecha.Caption = "Desde " + Format(Me.TxtFINIVaca.Value, "mm/dd/yyyy") + " Hasta " + Format(Me.TxtFFinVaca.Value, "mm/dd/yyyy")
      ArepPersecionesVaca.LblFechaHoy = Format(Now, "dddddd")
      ArepPersecionesVaca.DataControl1.ConnectionString = ConexionReporte
      
        If Me.DBCNominas.Text <> "Administracion" Then
              
         If Me.ChkRestar.Value = 1 Then
               SQlReportes = "SELECT  NomVaca.NumNomVaca AS NumNomVaca, SUM(DetalleNomVaca.Inss) AS Inss,SUM(DetalleNomVaca.Ir) AS Ir, MAX(Empleado.CodEmpleado1) AS CodEmpleado1,MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, SUM(DetalleNomVaca.SalarioMensual) AS SalarioMensual, SUM(DetalleNomVaca.DiasAPagar) AS DiasAPagar, SUM(DetalleNomVaca.DiasDescuento) AS DiasDescuento, SUM(DetalleNomVaca.AdelantoVacaciones) AS Adelanto13vo, SUM(DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir) " & _
                             "AS MontoPagar, SUM(DetalleNomVaca.TotalDevengado) AS TotalDevengado, MAX(Historico.FechaContrato) AS FechaContrato, MAX(Empleado.TarifaHoraria) AS TarifaHoraria, SUM(DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir) AS TotalDeducir, MAX(NomVaca.CodTipoNomina) AS CodTipoNomina  FROM  NomVaca INNER JOIN Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON  " & _
                             "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado Where (Empleado.Activo = 1) GROUP BY NomVaca.NumNomVaca HAVING (SUM(DetalleNomVaca.DiasAPagar) <> 0) AND (MAX(NomVaca.CodTipoNomina) = '" & CodTipoNomina & "') AND (NomVaca.NumNomVaca =" & NumNomVaca & ") ORDER BY MAX(Empleado.CodEmpleado1)"
         Else
               SQlReportes = "SELECT  NomVaca.NumNomVaca AS NumNomVaca, SUM(DetalleNomVaca.Inss) AS Inss, SUM(DetalleNomVaca.Ir) AS Ir, MAX(Empleado.CodEmpleado1) AS CodEmpleado1,MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, SUM(DetalleNomVaca.SalarioMensual) AS SalarioMensual, SUM(DetalleNomVaca.DiasAPagar) AS DiasAPagar, SUM(DetalleNomVaca.DiasDescuento) AS DiasDescuento, SUM(DetalleNomVaca.AdelantoVacaciones) AS Adelanto13vo, SUM(DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss) AS MontoPagar, SUM(DetalleNomVaca.TotalDevengado) AS TotalDevengado, MAX(Historico.FechaContrato) AS FechaContrato, MAX(Empleado.TarifaHoraria) AS TarifaHoraria, SUM(DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss) AS TotalDeducir, MAX(NomVaca.CodTipoNomina) AS CodTipoNomina  FROM  NomVaca INNER JOIN Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON  " & _
                             "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado Where (Empleado.Activo = 1) GROUP BY NomVaca.NumNomVaca HAVING (SUM(DetalleNomVaca.DiasAPagar) <> 0) AND (MAX(NomVaca.CodTipoNomina) = '" & CodTipoNomina & "') AND (NomVaca.NumNomVaca =" & NumNomVaca & ") ORDER BY MAX(Empleado.CodEmpleado1)"
         End If
        Else
        
        If Me.ChkRestar.Value = 1 Then
               SQlReportes = "SELECT  NomVaca.NumNomVaca AS NumNomVaca, SUM(DetalleNomVaca.Inss) AS Inss, SUM(DetalleNomVaca.Ir) AS Ir, MAX(Empleado.CodEmpleado1) AS CodEmpleado1,MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, SUM(DetalleNomVaca.SalarioMensual) AS SalarioMensual, SUM(DetalleNomVaca.DiasAPagar) AS DiasAPagar, SUM(DetalleNomVaca.DiasDescuento) AS DiasDescuento, SUM(DetalleNomVaca.AdelantoVacaciones) AS Adelanto13vo, SUM(DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss -DetalleNomVaca.Ir) " & _
                             " AS MontoPagar, SUM(DetalleNomVaca.TotalDevengado) AS TotalDevengado, MAX(Historico.FechaContrato) AS FechaContrato, MAX(Empleado.TarifaHoraria) AS TarifaHoraria, SUM(DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir) AS TotalDeducir, MAX(NomVaca.CodTipoNomina) AS CodTipoNomina  FROM  NomVaca INNER JOIN Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON  " & _
                             "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado Where (Empleado.Activo = 1) GROUP BY NomVaca.NumNomVaca HAVING (SUM(DetalleNomVaca.DiasAPagar) <> 0) AND (MAX(NomVaca.CodTipoNomina) = '" & CodTipoNomina & "') AND (NomVaca.NumNomVaca =" & NumNomVaca & ") ORDER BY MAX(Empleado.CodEmpleado1)"
        
        Else
               SQlReportes = "SELECT  NomVaca.NumNomVaca AS NumNomVaca, SUM(DetalleNomVaca.Inss) AS Inss, SUM(DetalleNomVaca.Ir) AS Ir, MAX(Empleado.CodEmpleado1) AS CodEmpleado1,MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, SUM(DetalleNomVaca.SalarioMensual) AS SalarioMensual, SUM(DetalleNomVaca.DiasAPagar) AS DiasAPagar, SUM(DetalleNomVaca.DiasDescuento) AS DiasDescuento, SUM(DetalleNomVaca.AdelantoVacaciones) AS Adelanto13vo, SUM(DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss) AS MontoPagar, SUM(DetalleNomVaca.TotalDevengado) AS TotalDevengado, MAX(Historico.FechaContrato) AS FechaContrato, MAX(Empleado.TarifaHoraria) AS TarifaHoraria, SUM(DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss) AS TotalDeducir, MAX(NomVaca.CodTipoNomina) AS CodTipoNomina  FROM  NomVaca INNER JOIN Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON  " & _
                             "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado Where (Empleado.Activo = 1) GROUP BY NomVaca.NumNomVaca HAVING (SUM(DetalleNomVaca.DiasAPagar) <> 0) AND (MAX(NomVaca.CodTipoNomina) = '" & CodTipoNomina & "') AND (NomVaca.NumNomVaca =" & NumNomVaca & ") ORDER BY MAX(Empleado.CodEmpleado1)"
        End If
      End If
      
'      ArepPersecionesVaca.Label5.Caption = "Nomina de Vacaciones"
      ArepPersecionesVaca.DataControl1.Source = SQlReportes
'      ArepPersecionesVaca.ImgLogo.Picture = LoadPicture(RutaLogo)
      ArepPersecionesVaca.Show 1


Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdPrnNomina_Click()
On Error GoTo TipoErrs
Dim Espacio As String, CodTipoNomina As String
Dim CodDepartamento As String, NombreDepartamento As String
Espacio = " "

 If Me.ChkColillaDpto.Value = 1 Then
 
 
       '////////////////////////////////////////////////////////////////////////////
       '////////////////BUSCO EL TIPO DE LA NOMINA //////////////////////////////////
       '//////////////////////////////////////////////////////////////////////////////
       DtaTipoNomina.Refresh
        Do While Not DtaTipoNomina.Recordset.EOF
        If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
           CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
           Exit Do
        End If
        DtaTipoNomina.Recordset.MoveNext
        Loop
        
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////////CONSULTO LOS DEPARTAMENTOS PARA EL TIPO DE NOMINA ////////////////////////
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Me.AdoBusca.RecordSource = "SELECT DISTINCT Departamento.CodDepartamento, Departamento.Departamento, Empleado.CodTipoNomina FROM Departamento INNER JOIN  Empleado ON Departamento.CodDepartamento = Empleado.CodDepartamento  " & _
                               "WHERE  (Empleado.CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Departamento.CodDepartamento"
    Me.AdoBusca.Refresh
    Do While Not Me.AdoBusca.Recordset.EOF
        
        Quien = "Colilla de Pago Aguinaldo"
        NumNomina = Me.TxtNumNom13.Text
        
        CodDepartamento = Me.AdoBusca.Recordset("CodDepartamento")
        NombreDepartamento = Me.AdoBusca.Recordset("Departamento")
        
        MsgBox "SE IMPRIMIRA DEPARTAMENTO: " & NombreDepartamento, vbInformation, "Zeus Nominas"
        
        
       
'        SQlReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, DetalleNom13Mes.MontoPension, DetalleNom13Mes.MontoSuspension, DetalleNom13Mes.DiasSuspension, DetalleNom13Mes.PorcentajePension, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo AS MontoPagar, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Cargo.Cargo, DetalleNom13Mes.SalarioAPagar AS TotalDevengado, Empleado.CodEmpleado1, Historico.FechaContrato, DetalleNom13Mes.MontoPension + DetalleNom13Mes.MontoSuspension + DetalleNom13Mes.Adelanto13vo AS TotalDeducir FROM Nom13Mes INNER JOIN Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN " & _
'                      "DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado  " & _
'                      "WHERE (Nom13Mes.NumNom13Mes = " & NumNomina & ") AND (Empleado.CodDepartamento = '" & CodDepartamento & "') ORDER BY Nombres"
        
        SQlReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, DetalleNom13Mes.MontoPension, DetalleNom13Mes.MontoSuspension, DetalleNom13Mes.DiasSuspension, DetalleNom13Mes.PorcentajePension, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo AS MontoPagar,  Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Cargo.Cargo, DetalleNom13Mes.SalarioAPagar AS TotalDevengado, Empleado.CodEmpleado1, Historico.FechaContrato, DetalleNom13Mes.MontoPension + DetalleNom13Mes.MontoSuspension + DetalleNom13Mes.Adelanto13vo AS TotalDeducir, Departamento.Departamento FROM  Nom13Mes INNER JOIN Cargo INNER JOIN  Empleado ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN " & _
                      "DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento " & _
                      "WHERE (Nom13Mes.NumNom13Mes = " & NumNomina & ") AND (Empleado.CodDepartamento = '" & CodDepartamento & "') ORDER BY Nombres"
        ArepColilla13vo.AdoColillas.Source = SQlReportes
        ArepColilla13vo.LblTipo.Caption = Me.DBCNominas.Text
        ArepColilla13vo.LblPeriodos.Caption = "Desde   " & Me.TxtFINI13.Value & " Hasta    " & Me.TxtFFIN13.Value
        ArepColilla13vo.LblPeriodo.Caption = Format(Me.TxtFINI13.Value, "dddddd") & "   Hasta   " & Format(Me.TxtFFIN13.Value, "dddddd")
        ArepColilla13vo.lbltitulo.Caption = Titulo
        ArepColilla13vo.AdoColillas.ConnectionString = ConexionReporte
        ArepColilla13vo.Show 1
        
        
        
    
      Me.AdoBusca.Recordset.MoveNext
    Loop
    
 
  Else
      
    
    
        Quien = "Colilla de Pago Aguinaldo"
        NumNomina = Me.TxtNumNom13.Text
        
'        SQlReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, DetalleNom13Mes.MontoPension, DetalleNom13Mes.MontoSuspension, DetalleNom13Mes.DiasSuspension, DetalleNom13Mes.PorcentajePension, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo AS MontoPagar, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Cargo.Cargo, DetalleNom13Mes.SalarioAPagar AS TotalDevengado, Empleado.CodEmpleado1, Historico.FechaContrato, DetalleNom13Mes.MontoPension + DetalleNom13Mes.MontoSuspension + DetalleNom13Mes.Adelanto13vo AS TotalDeducir FROM Nom13Mes INNER JOIN Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN " & _
'                      "DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado  " & _
'                      "WHERE (Nom13Mes.NumNom13Mes = " & NumNomina & ") ORDER BY Nombres"
        SQlReportes = "SELECT  Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, DetalleNom13Mes.MontoPension, DetalleNom13Mes.MontoSuspension, DetalleNom13Mes.DiasSuspension, DetalleNom13Mes.PorcentajePension, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo AS MontoPagar, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Cargo.Cargo, DetalleNom13Mes.SalarioAPagar AS TotalDevengado, Empleado.CodEmpleado1, Historico.FechaContrato, DetalleNom13Mes.MontoPension + DetalleNom13Mes.MontoSuspension + DetalleNom13Mes.Adelanto13vo AS TotalDeducir, Departamento.Departamento FROM  Nom13Mes INNER JOIN Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN " & _
                      "DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento  " & _
                      "Where (Nom13Mes.NumNom13Mes = " & NumNomina & ") ORDER BY Nombres"
        ArepColilla13vo.AdoColillas.Source = SQlReportes
        ArepColilla13vo.LblTipo.Caption = Me.DBCNominas.Text
        ArepColilla13vo.LblPeriodos.Caption = "Desde   " & Me.TxtFINI13.Value & " Hasta    " & Me.TxtFFIN13.Value
        ArepColilla13vo.LblPeriodo.Caption = Format(Me.TxtFINI13.Value, "dddddd") & "   Hasta   " & Format(Me.TxtFFIN13.Value, "dddddd")
        ArepColilla13vo.lbltitulo.Caption = Titulo
        ArepColilla13vo.AdoColillas.ConnectionString = ConexionReporte
        ArepColilla13vo.Show 1
 End If


Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdprNomina_Click()
On Error GoTo TipoErrs
Dim Espacio As String
Dim CodTipoNomina As String
Espacio = " "
NumNomina = Me.TxtNumNom13.Text

Dim rpt As New Nom13vo


      rpt.lbltitulo.Caption = Titulo
      rpt.LblSubtitulo.Caption = SubTitulo
      rpt.ImgLogo.Picture = LoadPicture(RutaLogo)
      
      rpt.LblFecha.Caption = "Desde " + Format(Me.TxtFINI13.Value, "dd/mm/yyyy") + " Hasta " + Format(Me.TxtFFIN13.Value, "dd/mm/yyyy")
      rpt.LblFechaHoy = Format(Now, "dddddd")
      rpt.DataControl1.ConnectionString = ConexionReporte
      SQlReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual,     DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo AS MontoPagar,      Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Cargo.Cargo, DetalleNom13Mes.SalarioAPagar AS TotalDevengado, Empleado.CodEmpleado1,     DetalleNom13Mes.TotalDeducciones, Historico.FechaContrato as FechaIngreso  FROM            Nom13Mes INNER JOIN   Cargo INNER JOIN   Empleado ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN  DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes INNER JOIN    Historico ON Empleado.CodEmpleado = Historico.Codempleado    WHERE        (Nom13Mes.NumNom13Mes = " & NumNomina & ")   ORDER BY Nombres"
      SQlReportes = "SELECT  Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo AS MontoPagar, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Cargo.Cargo, DetalleNom13Mes.SalarioAPagar AS TotalDevengado, Empleado.CodEmpleado1, DetalleNom13Mes.TotalDeducciones, Historico.FechaContrato AS FechaIngreso, departamento.CodDepartamento , departamento.departamento  " & _
                    "FROM  Nom13Mes INNER JOIN Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento  " & _
                    "Where (Nom13Mes.NumNom13Mes = " & NumNomina & ") ORDER BY Departamento.Departamento, Nombres"
      rpt.DataControl1.Source = SQlReportes
      rpt.ImgLogo.Picture = LoadPicture(RutaLogo)
     ' Nom13vo.Show 1
     Dim fPreview As New FrmPreview
     fPreview.RunReport rpt
     fPreview.Show 1
     


Exit Sub
TipoErrs:
ControlErrores

End Sub

Private Sub CmdPRVaca_Click()
On Error GoTo TipoErrs
Dim Espacio As String
Espacio = " "
NumNomVaca = DtaNomVaca.Recordset("NumNomVaca")


      ArepNomVacaciones.lbltitulo.Caption = Titulo
      ArepNomVacaciones.LblSubtitulo.Caption = SubTitulo
      ArepNomVacaciones.ImgLogo.Picture = LoadPicture(RutaLogo)
      
      ArepNomVacaciones.LblFecha.Caption = "Desde    " + Format(Me.TxtFFinVaca.Value, "mm/dd/yyyy") + "     Hasta     " + Format(Me.TxtFFinVaca.Value, "mm/dd/yyyy")
      ArepNomVacaciones.LblFechaHoy = Format(Now, "dddddd")
      ArepNomVacaciones.DataControl1.ConnectionString = ConexionReporte
      SQlReportes = "SELECT NomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, ([DetalleNomVaca].[SalarioMensual]*([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento])/" & DiasMes & ")-[DetalleNomVaca].[AdelantoVacaciones] AS MontoAPagar, [DetalleNomVaca].[SalarioMensual]*([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento])/" & DiasMes & " AS TotalDevengado, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, ([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento]) AS TotalDescuento " & vbLf
      SQlReportes = SQlReportes & "FROM NomVaca INNER JOIN (Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado) ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca Where (((NomVaca.NumNomVaca) = " & NumNomVaca & " )) ORDER BY DetalleNomVaca.CodEmpleado"
      ArepNomVacaciones.DataControl1.Source = SQlReportes
      ArepNomVacaciones.ImgLogo.Picture = LoadPicture(RutaLogo)
      ArepNomVacaciones.Show 1
      
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdSalir_Click()
Unload Me

End Sub

Private Sub Command1_Click()
Quien = "Nomina13vo"
FrmBuscaEmpleado.Show 1
End Sub

Private Sub Command2_Click()
Quien = "NominaVaca"
FrmBuscaEmpleado.Show 1
End Sub

Private Sub Command3_Click()
Quien = "CalcularNomina"
FrmExportaBac.Show 1
End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub Command4_Click()

End Sub

Private Sub DBCNominas_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String

DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop



DtaControles.Refresh
DiasMes = DtaControles.Recordset("DiasMes")
DiasSemana = DtaControles.Recordset("DiasSemana")

'/////////////////////////////////////////////////////////////////////////
'///////////CALCULO DE LAS VACACIONES////////////////////////////////////
'///////////////////////////////////////////////////////////////////////


'//////////Busco si existen Nominas de Vacaciones//////////////////////
'//////////Creadas en el sistema///////////////////////////////////////
Me.DtaConsulta.RecordSource = "SELECT NomVaca.NumNomVaca, NomVaca.FechaAplica, NomVaca.FechaIni, NomVaca.FechaFin, Transfereir From NomVaca ORDER BY NomVaca.FechaFin"
Me.DtaConsulta.Refresh

If DtaConsulta.Recordset.EOF Then
  Mes = Month(Me.TxtFFinVaca.Value)
  Año = Year(Me.TxtFFinVaca.Value)
  If Mes <= 6 Then
   FechaIniVaca = CDate("01/01/" & Str(Año))
   FechaFinVaca = CDate("30/06/" & Str(Año))
  Else
   FechaIniVaca = CDate("1/7/" & Str(Año))
   FechaFinVaca = CDate("31/12/" & Str(Año))
 End If

  Me.TxtNumNomVaca.Text = 1
Else

 If Me.DtaConsulta.Recordset("Transfereir") = True Then
   Me.CHKTranferir.Value = 1
 Else
   Me.CHKTranferir.Value = 0
 End If
 
'/////////busco si existen Nominas Activas en el Sistema//////////////////
 Me.DtaConsulta.RecordSource = "SELECT NomVaca.NumNomVaca, NomVaca.FechaAplica, NomVaca.FechaIni, NomVaca.FechaFin, NomVaca.Activa From NomVaca Where (((NomVaca.Activa) = 1))AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY NomVaca.FechaFin"
 Me.DtaConsulta.Refresh
 If DtaConsulta.Recordset.EOF Then
  '///Si no existen nominas Activas me ubica en la ultima nomina Creada//////
  Me.DtaConsulta.RecordSource = "SELECT NomVaca.NumNomVaca, NomVaca.FechaAplica, NomVaca.FechaIni, NomVaca.FechaFin From NomVaca ORDER BY NomVaca.FechaFin"
  Me.DtaConsulta.Refresh
  DtaConsulta.Recordset.MoveLast
  Mes = Month(CDate(DtaConsulta.Recordset("Fechafin")))
  Año = Year(CDate(DtaConsulta.Recordset("Fechafin")))
  If Mes <= 6 Then
   FechaIniVaca = CDate("01/07/" & Str(Año))
   FechaFinVaca = CDate("31/12/" & Str(Año))
  Else
   Año = Año + 1
   FechaIniVaca = CDate("01/01/" & Str(Año))
   FechaFinVaca = CDate("30/06/" & Str(Año))
  End If
 
  Me.DtaConsulta.RecordSource = "SELECT NumNomVaca, FechaAplica, FechaIni, FechaFin From NomVaca ORDER BY NumNomVaca"
  Me.DtaConsulta.Refresh
  If Not Me.DtaConsulta.Recordset.EOF Then
   Me.DtaConsulta.Recordset.MoveLast
   Me.TxtNumNomVaca.Text = DtaConsulta.Recordset("NumNomVaca") + 1
   Me.DtaConsulta.Recordset.MoveFirst
  End If

 Else
  '//Si Existen nominas Activas me ubico en la ultima/////////
  DtaConsulta.Recordset.MoveLast
  Mes = Month(CDate(DtaConsulta.Recordset("Fechafin")))
  Año = Year(CDate(DtaConsulta.Recordset("Fechafin")))
  FechaIniVaca = CDate(DtaConsulta.Recordset("Fechaini"))
  FechaFinVaca = CDate(DtaConsulta.Recordset("Fechafin"))
  Me.TxtNumNomVaca.Text = DtaConsulta.Recordset("NumNomVaca")
 End If
End If

'/////////////////////////////////////////////////////////////////////////////
'///////////////CALCULO DEL 13VO MES//////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////


'/////////Busco si existen Nominas de 13vo//////////////////
 Me.DtaConsulta.RecordSource = "SELECT Nom13Mes.NumNom13Mes, Nom13Mes.FechaIni, Nom13Mes.FechaFin, Nom13Mes.Activa From Nom13Mes ORDER BY Nom13Mes.FechaFin"
 Me.DtaConsulta.Refresh
 If DtaConsulta.Recordset.EOF Then
  FechaIni13 = CDate("01/12/" & Str(Year(Now) - 1))
  FechaFin13 = CDate("30/11/" & Str(Year(Now)))
'  TxtNumNomVaca = 1
  TxtNumNom13 = 1
 Else
'///Busco si Existen nominas de 13vo mes Activas/////////
 Me.DtaConsulta.RecordSource = "SELECT Nom13Mes.NumNom13Mes, Nom13Mes.FechaIni, Nom13Mes.FechaFin, Nom13Mes.Activa From  Nom13Mes WHERE (Activa = 1) AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Nom13Mes.NumNom13Mes"
 Me.DtaConsulta.Refresh
  If DtaConsulta.Recordset.EOF Then
   '///Si no existen nominas Activas me ubica en la ultima nomina Creada//////
   Me.DtaConsulta.RecordSource = "SELECT Nom13Mes.NumNom13Mes, Nom13Mes.FechaIni, Nom13Mes.FechaFin, Nom13Mes.Activa From Nom13Mes ORDER BY Nom13Mes.FechaFin"
   Me.DtaConsulta.Refresh
   DtaConsulta.Recordset.MoveLast
   Mes = Month(CDate(DtaConsulta.Recordset("Fechafin")))
   Año = Year(CDate(DtaConsulta.Recordset("Fechafin")))
   FechaIni13 = CDate("1/12/" & Str(Year(Año)))
   FechaFin13 = CDate("30/11/" & Str(Year(Año) + 1))
   
   Me.DtaConsulta.RecordSource = "SELECT NumNom13Mes, FechaIni, FechaFin, Activa From Nom13Mes ORDER BY NumNom13Mes"
   Me.DtaConsulta.Refresh
   If Not DtaConsulta.Recordset.EOF Then
    Me.DtaConsulta.Recordset.MoveLast
    TxtNumNom13 = DtaConsulta.Recordset("NumNom13Mes") + 1
    Me.DtaConsulta.Recordset.MoveFirst
   End If
  Else
   '//Si Existen nominas Activas me ubico en la ultima/////////

    DtaConsulta.Recordset.MoveLast
    Mes = Month(CDate(DtaConsulta.Recordset("Fechafin")))
    Año = Year(CDate(DtaConsulta.Recordset("Fechafin")))
    FechaIni13 = CDate(DtaConsulta.Recordset("Fechaini"))
    FechaFin13 = CDate(DtaConsulta.Recordset("Fechafin"))
    TxtNumNom13 = DtaConsulta.Recordset("NumNom13Mes")

  End If
 
 
 End If


TxtFINIVaca.Value = Format(FechaIniVaca, "DD/MM/YYYY")
TxtFFinVaca.Value = Format(FechaFinVaca, "DD/MM/YYYY")

TxtFINI13.Value = Format(FechaIni13, "DD/MM/YYYY")
TxtFFIN13.Value = Format(FechaFin13, "DD/MM/YYYY")

DtaConsecutivos.Refresh

NumNomVaca = val(TxtNumNomVaca.Text)
NumNom13Mes = val(TxtNumNom13.Text)


Mes1 = Month(Me.TxtFINIVaca.Value)
Año1 = Year(Me.TxtFINIVaca.Value)
Mes2 = Month(Me.TxtFFinVaca.Value)
Año2 = Year(Me.TxtFFinVaca.Value)

Mes1 = Format(Mes1, "0#")
Mes2 = Format(Mes2, "0#")


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.dtpFPInicio.Value = Me.AdoBusca.Recordset("Inicio")
 End If

Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.dtpFPFinal.Value = Me.AdoBusca.Recordset("Final")
 End If


'SqlVacaciones = "SELECT NomVaca.NumNomVaca, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, ([DetalleNomVaca].[SalarioMensual]*([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento])/" & DiasMes & ")-[DetalleNomVaca].[AdelantoVacaciones] AS MontoAPagar FROM NomVaca INNER JOIN (Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado) ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca Where (((NomVaca.NumNomVaca) = " & NumNomVaca & ")) ORDER BY DetalleNomVaca.CodEmpleado"
SqlVacaciones = "SELECT NomVaca.NumNomVaca, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2,DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, DetalleNomVaca.SalarioMensual * (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento)/ '" & DiasMes & "' - DetalleNomVaca.AdelantoVacaciones AS MontoAPagar, NomVaca.CodTipoNomina FROM NomVaca INNER JOIN Empleado INNER JOIN       DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca WHERE (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1"
DtaVacaciones.RecordSource = SqlVacaciones
DtaVacaciones.Refresh

Me.DbgrVacaciones.Columns(0).Locked = True
Me.DbgrVacaciones.Columns(0).Visible = False
Me.DbgrVacaciones.Columns(1).Locked = True
Me.DbgrVacaciones.Columns(2).Locked = True
Me.DbgrVacaciones.Columns(3).Locked = True
Me.DbgrVacaciones.Columns(4).Locked = True
Me.DbgrVacaciones.Columns(5).Locked = True
Me.DbgrVacaciones.Columns(6).Locked = True
Me.DbgrVacaciones.Columns(7).Locked = True
Me.DbgrVacaciones.Columns(9).Locked = True
Me.DbgrVacaciones.Columns(10).Locked = True
Me.DbgrVacaciones.Columns(11).Locked = True
Me.DbgrVacaciones.Columns(6).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(7).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(8).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(9).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(10).NumberFormat = "##,##0.00"


Mes1 = Month(Me.TxtFINI13.Value)
Año1 = Year(Me.TxtFINI13.Value)
Mes2 = Month(Me.TxtFFIN13.Value)
Año2 = Year(Me.TxtFFIN13.Value)

Mes1 = Format(Mes1, "0#")
Mes2 = Format(Mes2, "0#")

DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.DtpInicio13vo.Value = Me.AdoBusca.Recordset("Inicio")
 End If

Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.DtpFin13vo.Value = Me.AdoBusca.Recordset("Final")
 End If



'Sql13voMes = "SELECT Nom13Mes.NumNom13Mes, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, ([DetalleNom13Mes].[SalarioMensual]*[DetalleNom13Mes].[DiasAPagar]/'" & DiasMes & "')-[DetalleNom13Mes].[Adelanto13vo] AS MontoPagar, Nom13Mes.CodTipoNomina FROM Nom13Mes INNER JOIN (Empleado INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNom13Mes & ")) AND (Nom13Mes.CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Empleado.CodEmpleado1"
Sql13voMes = "SELECT Nom13Mes.NumNom13Mes, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, DetalleNom13Mes.SalarioAPagar-DetalleNom13Mes.Adelanto13vo AS MontoPagar FROM Nom13Mes INNER JOIN (Empleado INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNom13Mes & ")) ORDER BY Empleado.CodEmpleado1"
Dta13voMes.RecordSource = Sql13voMes
Dta13voMes.Refresh

Me.Dbgr13Mes.Columns(0).Locked = True
Me.Dbgr13Mes.Columns(0).Visible = False
Me.Dbgr13Mes.Columns(1).Locked = True
Me.Dbgr13Mes.Columns(2).Locked = True
Me.Dbgr13Mes.Columns(3).Locked = True
Me.Dbgr13Mes.Columns(4).Locked = True
Me.Dbgr13Mes.Columns(5).Locked = True
Me.Dbgr13Mes.Columns(6).Locked = True
Me.Dbgr13Mes.Columns(7).Locked = True
Me.Dbgr13Mes.Columns(9).Locked = True
Me.Dbgr13Mes.Columns(6).NumberFormat = "##,##0.00"
Me.Dbgr13Mes.Columns(7).NumberFormat = "##,##0.00"
Me.Dbgr13Mes.Columns(8).NumberFormat = "##,##0.00"
Me.Dbgr13Mes.Columns(9).NumberFormat = "##,##0.00"
End Sub


Private Sub Form_Load()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String
Dim FechaIniVaca As Date, Mes As Integer, Dia As Integer, Año As Integer
Dim FechaFinVaca As Date
Dim FechaIni13 As Date
Dim FechaFin13 As Date
Dim SqlVacaciones As String
Dim Sql13voMes As String
Dim NumNomVaca As Long
Dim NumNom13Mes As Long

 Me.Dbgr13Mes.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.Dbgr13Mes.OddRowStyle.BackColor = &H80000005
 Me.Dbgr13Mes.AlternatingRowStyle = True
 
  Me.DbgrVacaciones.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DbgrVacaciones.OddRowStyle.BackColor = &H80000005
 Me.DbgrVacaciones.AlternatingRowStyle = True
 
 Me.TxtFechaAplica.Value = Format(Now, "dd/mm/yyyy")
 
 With Me.AdoHistorialSalarial
 '.DatabaseName = Ruta
 .ConnectionString = Conexion

End With

With Me.DtaDetalleNominas
 '.DatabaseName = Ruta
 .ConnectionString = Conexion

End With

With Me.AdoBusca
 .ConnectionString = Conexion
End With

With Me.DtaAdelanto
 '.DatabaseName = Ruta
 .ConnectionString = Conexion
End With

With Me.DtaHistorico
  '.DatabaseName = Ruta
  .ConnectionString = Conexion
End With

With Me.DtaControles
  '.DatabaseName = Ruta
  .ConnectionString = Conexion
  .RecordSource = "Controles"
  .Refresh
End With

With Me.Dta13voMes
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaConsecutivos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Consecutivos"
   .Refresh
End With

With Me.DtaDetalleNom13Mes
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaDetalleNomVaca
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "DetalleNomVaca"
End With


With Me.DtaEmpleados
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With


With Me.DtaNom13Mes
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With


With Me.DtaNominas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaNomVaca
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "NomVaca"
   .Refresh
End With

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaTipoNomina
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "TipoNomina"
   .Refresh
   
End With


With Me.DtaVacaciones
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With


With AdoConsulta
.ConnectionString = Conexion
End With
DtaControles.Refresh
DiasMes = DtaControles.Recordset("DiasMes")
DiasSemana = DtaControles.Recordset("DiasSemana")

'//////////Busco si existen Nominas de Vacaciones//////////////////////
'//////////Creadas en el sistema///////////////////////////////////////
Me.DtaConsulta.RecordSource = "SELECT NomVaca.NumNomVaca, NomVaca.FechaAplica, NomVaca.FechaIni, NomVaca.FechaFin From NomVaca ORDER BY NomVaca.FechaFin"
Me.DtaConsulta.Refresh

If DtaConsulta.Recordset.EOF Then
  Mes = Month(Me.TxtFFinVaca.Value)
  Año = Year(Me.TxtFFinVaca.Value)
  If Mes <= 6 Then
   FechaIniVaca = CDate("01/01/" & Str(Año))
   FechaFinVaca = CDate("30/06/" & Str(Año))
  Else
   FechaIniVaca = CDate("1/7/" & Str(Año))
   FechaFinVaca = CDate("31/12/" & Str(Año))
 End If

  Me.TxtNumNomVaca.Text = 1
Else
'/////////busco si existen Nominas Activas en el Sistema//////////////////
 Me.DtaConsulta.RecordSource = "SELECT NomVaca.NumNomVaca, NomVaca.FechaAplica, NomVaca.FechaIni, NomVaca.FechaFin, NomVaca.Activa From NomVaca Where (((NomVaca.Activa) = 1)) ORDER BY NomVaca.FechaFin"
 Me.DtaConsulta.Refresh
 If DtaConsulta.Recordset.EOF Then
  '///Si no existen nominas Activas me ubica en la ultima nomina Creada//////
  Me.DtaConsulta.RecordSource = "SELECT NomVaca.NumNomVaca, NomVaca.FechaAplica, NomVaca.FechaIni, NomVaca.FechaFin From NomVaca ORDER BY NomVaca.FechaFin"
  Me.DtaConsulta.Refresh
  DtaConsulta.Recordset.MoveLast
  Mes = Month(CDate(DtaConsulta.Recordset("Fechafin")))
  Año = Year(CDate(DtaConsulta.Recordset("Fechafin")))
  If Mes <= 6 Then
   FechaIniVaca = CDate("1/7/" & Str(Año))
   FechaFinVaca = CDate("31/12/" & Str(Año))
  Else
   Año = Año + 1
   FechaIniVaca = CDate("01/01/" & Str(Año))
   FechaFinVaca = CDate("30/06/" & Str(Año))
  End If
 
  Me.TxtNumNomVaca.Text = DtaConsulta.Recordset("NumNomVaca") + 1
 Else
  '//Si Existen nominas Activas me ubico en la ultima/////////
  DtaConsulta.Recordset.MoveLast
  Mes = Month(CDate(DtaConsulta.Recordset("Fechafin")))
  Año = Year(CDate(DtaConsulta.Recordset("Fechafin")))
  FechaIniVaca = CDate(DtaConsulta.Recordset("Fechaini"))
  FechaFinVaca = CDate(DtaConsulta.Recordset("Fechafin"))
  Me.TxtNumNomVaca.Text = DtaConsulta.Recordset("NumNomVaca")
 End If
End If


'/////////Busco si existen Nominas de 13vo//////////////////
 Me.DtaConsulta.RecordSource = "SELECT Nom13Mes.NumNom13Mes, Nom13Mes.FechaIni, Nom13Mes.FechaFin, Nom13Mes.Activa From Nom13Mes ORDER BY Nom13Mes.FechaFin"
 Me.DtaConsulta.Refresh
 If DtaConsulta.Recordset.EOF Then
  FechaIni13 = CDate("01/12/" & Str(Year(Now) - 1))
  FechaFin13 = CDate("30/11/" & Str(Year(Now)))
'  TxtNumNomVaca = 1
  TxtNumNom13 = 1
 Else
'///Busco si Existen nominas de 13vo mes Activas/////////
 Me.DtaConsulta.RecordSource = "SELECT Nom13Mes.NumNom13Mes, Nom13Mes.FechaIni, Nom13Mes.FechaFin, Nom13Mes.Activa From  Nom13Mes WHERE (Activa = 1) ORDER BY Nom13Mes.NumNom13Mes"
 Me.DtaConsulta.Refresh
  If DtaConsulta.Recordset.EOF Then
   '///Si no existen nominas Activas me ubica en la ultima nomina Creada//////
   Me.DtaConsulta.RecordSource = "SELECT Nom13Mes.NumNom13Mes, Nom13Mes.FechaIni, Nom13Mes.FechaFin, Nom13Mes.Activa From Nom13Mes ORDER BY Nom13Mes.FechaFin"
   Me.DtaConsulta.Refresh
   DtaConsulta.Recordset.MoveLast
   Mes = Month(CDate(DtaConsulta.Recordset("Fechafin")))
   Año = Year(CDate(DtaConsulta.Recordset("Fechafin")))
   FechaIni13 = CDate("1/12/" & Str(Year(Año)))
   FechaFin13 = CDate("30/11/" & Str(Year(Año) + 1))
   TxtNumNom13 = DtaConsulta.Recordset("NumNom13Mes") + 1
  Else
   '//Si Existen nominas Activas me ubico en la ultima/////////

    DtaConsulta.Recordset.MoveLast
    Mes = Month(CDate(DtaConsulta.Recordset("Fechafin")))
    Año = Year(CDate(DtaConsulta.Recordset("Fechafin")))
    FechaIni13 = CDate(DtaConsulta.Recordset("Fechaini"))
    FechaFin13 = CDate(DtaConsulta.Recordset("Fechafin"))
    TxtNumNom13 = DtaConsulta.Recordset("NumNom13Mes")

  End If
 
 
 End If


TxtFINIVaca.Value = Format(FechaIniVaca, "DD/MM/YYYY")
TxtFFinVaca.Value = Format(FechaFinVaca, "DD/MM/YYYY")

TxtFINI13.Value = Format(FechaIni13, "DD/MM/YYYY")
TxtFFIN13.Value = Format(FechaFin13, "DD/MM/YYYY")

DtaConsecutivos.Refresh

NumNomVaca = val(TxtNumNomVaca.Text)
NumNom13Mes = val(TxtNumNom13.Text)



Mes1 = Month(Me.TxtFINIVaca.Value)
Año1 = Year(Me.TxtFFinVaca.Value)
Mes2 = Month(Me.TxtFFinVaca.Value)
Año2 = Year(Me.TxtFFinVaca.Value)

Mes1 = Format(Mes1, "0#")
Mes2 = Format(Mes2, "0#")

DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.dtpFPInicio.Value = Me.AdoBusca.Recordset("Inicio")
 End If

Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.dtpFPFinal.Value = Me.AdoBusca.Recordset("Final")
 End If


'SqlVacaciones = "SELECT NomVaca.NumNomVaca, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, ([DetalleNomVaca].[SalarioMensual]*([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento])/" & DiasMes & ")-[DetalleNomVaca].[AdelantoVacaciones] AS MontoAPagar FROM NomVaca INNER JOIN (Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado) ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca Where (((NomVaca.NumNomVaca) = " & NumNomVaca & ")) ORDER BY DetalleNomVaca.CodEmpleado"
SqlVacaciones = "SELECT NomVaca.NumNomVaca, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, ([DetalleNomVaca].[SalarioMensual]*([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento])/" & DiasMes & ")-[DetalleNomVaca].[AdelantoVacaciones] AS MontoAPagar FROM NomVaca INNER JOIN (Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado) ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca Where (((NomVaca.NumNomVaca) = " & NumNomVaca & ")) AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1"
DtaVacaciones.RecordSource = SqlVacaciones

DtaVacaciones.Refresh

Me.DbgrVacaciones.Columns(0).Locked = True
Me.DbgrVacaciones.Columns(0).Visible = False
Me.DbgrVacaciones.Columns(1).Locked = True
Me.DbgrVacaciones.Columns(2).Locked = True
Me.DbgrVacaciones.Columns(3).Locked = True
Me.DbgrVacaciones.Columns(4).Locked = True
Me.DbgrVacaciones.Columns(5).Locked = True
Me.DbgrVacaciones.Columns(6).Locked = True
Me.DbgrVacaciones.Columns(7).Locked = True
Me.DbgrVacaciones.Columns(8).Locked = False
Me.DbgrVacaciones.Columns(6).Caption = "Salario Promedio"
Me.DbgrVacaciones.Columns(10).Locked = True
Me.DbgrVacaciones.Columns(6).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(7).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(8).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(9).NumberFormat = "##,##0.00"
Me.DbgrVacaciones.Columns(10).NumberFormat = "##,##0.00"



'Sql13voMes = "SELECT Nom13Mes.NumNom13Mes, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, ([DetalleNom13Mes].[SalarioMensual]*[DetalleNom13Mes].[DiasAPagar]/" & DiasMes & ")-[DetalleNom13Mes].[Adelanto13vo] AS MontoPagar FROM Nom13Mes INNER JOIN (Empleado INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNom13Mes & ")) ORDER BY Empleado.CodEmpleado1"
Sql13voMes = "SELECT Nom13Mes.NumNom13Mes, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, ([DetalleNom13Mes].[SalarioMensual]*[DetalleNom13Mes].[DiasAPagar]/'" & DiasMes & "')-[DetalleNom13Mes].[Adelanto13vo] AS MontoPagar FROM Nom13Mes INNER JOIN (Empleado INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNom13Mes & ")) ORDER BY Empleado.CodEmpleado1"
Dta13voMes.RecordSource = Sql13voMes
Dta13voMes.Refresh

Me.Dbgr13Mes.Columns(0).Locked = True
Me.Dbgr13Mes.Columns(0).Visible = False
Me.Dbgr13Mes.Columns(1).Locked = True
Me.Dbgr13Mes.Columns(2).Locked = True
Me.Dbgr13Mes.Columns(3).Locked = True
Me.Dbgr13Mes.Columns(4).Locked = True
Me.Dbgr13Mes.Columns(5).Locked = True
Me.Dbgr13Mes.Columns(6).Locked = True
Me.Dbgr13Mes.Columns(7).Locked = True
Me.Dbgr13Mes.Columns(9).Locked = True
Me.Dbgr13Mes.Columns(6).NumberFormat = "##,##0.00"
Me.Dbgr13Mes.Columns(7).NumberFormat = "##,##0.00"
Me.Dbgr13Mes.Columns(8).NumberFormat = "##,##0.00"
Me.Dbgr13Mes.Columns(9).NumberFormat = "##,##0.00"
End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub

Private Sub TDBGrid1_Click()

End Sub

Private Sub PushButton1_Click()
On Error GoTo TipoErrs
Dim SQlReportes As String, V As Integer, H As Integer, i As Integer
Dim Año As String, MesLetra As String, Neto As String, Dias As String
Dim CanDias As String, QuinLetra As String, Nombres As String, Espacio As String
Dim TotalNomina As Double, Neto1 As Double, Cod As String, NetoT As String, Longitud As Integer
Dim CodigoCuenta As String, NombreEmpresa As String, MontoSubsidio As Double

Espacio = " "
Quien = "NominaVacaciones"
Select Case Quien
 Case "CalcularNomina"
       '//////////////////////Cargo la Consulta de la Nomina///////////////////////
  
   NumNomina = Me.TxtNumNomVaca.Text
   SQlReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, (DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo) AS MontoPagar, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, Cargo.Cargo, ([DetalleNom13Mes].SalarioAPagar) AS TotalDevengado, Empleado.CodEmpleado1, Empleado.CuentaBanco FROM Nom13Mes INNER JOIN ((Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNomina & ")) ORDER BY Nombres"



       Me.DtaConsulta.RecordSource = SQlReportes
       Me.DtaConsulta.Refresh


  
   Case "NominaVacaciones"
      NumNomVaca = Frm13VacaMes.TxtNumNomVaca.Text
      '///////////////////////////Cargo la Consulta de Vacaciones////////////////////////////////
      SQlReportes = "SELECT NomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado,Empleado.CuentaBanco, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, ([DetalleNomVaca].[SalarioMensual]*([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento])/" & DiasMes & ")-[DetalleNomVaca].[AdelantoVacaciones]-[DetalleNomVaca].[inss] AS MontoAPagar, [DetalleNomVaca].[SalarioMensual]*([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento])/" & DiasMes & " AS TotalDevengado, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, ([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento]) AS TotalDescuento " & vbLf
      SQlReportes = SQlReportes & "FROM NomVaca INNER JOIN (Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado) ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca Where (((NomVaca.NumNomVaca) = " & NumNomVaca & " )) AND (DetalleNomVaca.SalarioMensual * (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento) / 30 <> 0) ORDER BY Nombres"
'       Me.DtaExporta.Refresh
'        'Me.'Me.DtaExporta.Recordset.Edit
'       Me.DtaExporta.Recordset("CodigoBAC") = val(Me.TxtCod.Text)
'       Me.DtaExporta.Recordset.Update

      Me.DtaConsulta.RecordSource = SQlReportes
      Me.DtaConsulta.Refresh

       Mes = Month(Me.TxtFFinVaca.Value)
       Año = Year(Me.TxtFFinVaca.Value)
       CanDias = Day(Me.TxtFFinVaca.Value)
       Dias = Day(Me.TxtFFinVaca.Value)
'       Cod = Me.TxtCod.Text
End Select

            
   
    Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
    'Call Formato_Excel(8, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
V = 1
H = 0
i = 1

 
  Do While Not Me.DtaConsulta.Recordset.EOF 'esto nos sirve pa leer los datos desde
       
       CodEmpleado = DtaConsulta.Recordset("CodEmpleado")
       
       MontoSubsidio = 0
       
       If Not IsNull(DtaConsulta.Recordset("CuentaBanco")) Then
         CodigoCuenta = DtaConsulta.Recordset("CuentaBanco")
       Else
         CodigoCuenta = ""
       End If
 'la tabla de access para despues colocarlos en las celdas correspondientes
       
       Nombre = Me.DtaConsulta.Recordset("Nombres")
       Neto = Format(Me.DtaConsulta.Recordset("MontoAPagar") + MontoSubsidio, "####0.00")
       Neto1 = Format(Me.DtaConsulta.Recordset("MontoAPagar") + MontoSubsidio, "##,##0.00")
       Longitud = Len(Neto)
       NetoT = Mid(Neto, Longitud - 1, 3)
       NetoT = (Mid(Neto, 1, Longitud - 3)) & NetoT
       With DtaConsulta.Recordset

       
'           If Not (V = 1) Then
'             objExcel.ActiveSheet.Cells(V, H) = "T"
'           End If
            objExcel.ActiveSheet.Cells(V, H + 1) = Nombre
            objExcel.ActiveSheet.Cells(V, H + 2) = CodigoCuenta
            objExcel.ActiveSheet.Cells(V, H + 3) = "Vacaciones"
            objExcel.ActiveSheet.Cells(V, H + 4) = Format(Neto, "##,##0.00")
            objExcel.ActiveSheet.Cells(V, H + 5) = "C"
            V = V + 1
            i = i + 1
            TotalNomina = TotalNomina + Neto1
            .MoveNext

   
        End With
  Loop
  
  '/////////////////////////////SELECCION SOLO LOS EMPLEADOS QUE TIENEN SUBSIDIO Y NO TIENEN SALARIO
  Me.DtaConsulta.RecordSource = "SELECT TOP (200) DetalleNomSubsidio.id, DetalleNomSubsidio.NumNominaSubsidio, DetalleNomSubsidio.CodEmpleado, DetalleNomSubsidio.Subsidio, Empleado.CodEmpleado1, (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS Neto, Empleado.CuentaBanco, Empleado.Dolarizado, Empleado.FechaAntiguedad, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres " & _
                                "FROM DetalleNomSubsidio INNER JOIN Empleado ON DetalleNomSubsidio.CodEmpleado = Empleado.CodEmpleado INNER JOIN Nomina ON DetalleNomSubsidio.NumNominaSubsidio = Nomina.NumNomina INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado AND Nomina.NumNomina = DetalleNomina.NumNomina  " & _
                                "WHERE (DetalleNomSubsidio.NumNominaSubsidio = " & NumNomina & ") AND ((DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) = 0) AND (DetalleNomSubsidio.Subsidio <> 0) Order by Nombres"
  Me.DtaConsulta.Refresh
  Do While Not Me.DtaConsulta.Recordset.EOF
       If Not IsNull(DtaConsulta.Recordset("CuentaBanco")) Then
         CodigoCuenta = DtaConsulta.Recordset("CuentaBanco")
       Else
         CodigoCuenta = ""
       End If
 'la tabla de access para despues colocarlos en las celdas correspondientes
       
       Nombre = Me.DtaConsulta.Recordset("Nombres")
       Neto = Format(Me.DtaConsulta.Recordset("Subsidio"), "####0.00")
       Neto1 = Format(Me.DtaConsulta.Recordset("Subsidio"), "##,##0.00")
       Longitud = Len(Neto)
       NetoT = Mid(Neto, Longitud - 1, 3)
       NetoT = (Mid(Neto, 1, Longitud - 3)) & NetoT
       With DtaConsulta.Recordset

       
'           If Not (V = 1) Then
'             objExcel.ActiveSheet.Cells(V, H) = "T"
'           End If
            objExcel.ActiveSheet.Cells(V, H + 1) = Nombre
            objExcel.ActiveSheet.Cells(V, H + 2) = CodigoCuenta
            objExcel.ActiveSheet.Cells(V, H + 3) = QuinLetra
            objExcel.ActiveSheet.Cells(V, H + 4) = Format(Neto, "##,##0.00")
            objExcel.ActiveSheet.Cells(V, H + 5) = "C"
            V = V + 1
            i = i + 1
            TotalNomina = TotalNomina + Neto1
            

   
        End With
     Me.DtaConsulta.Recordset.MoveNext
  Loop
  
  
     
       MDIPrimero.DtaEmpresa.Refresh
       If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("NombreEmpresa")) Then
         NombreEmpresa = MDIPrimero.DtaEmpresa.Recordset("NombreEmpresa")
       End If
       Neto = Format(TotalNomina, "####0.00")
       Longitud = Len(Neto)
       NetoT = Mid(Neto, Longitud - 1, 3)
       NetoT = (Mid(Neto, 1, Longitud - 3)) & NetoT
   

       objExcel.ActiveSheet.Cells(V, 1) = NombreEmpresa
       objExcel.ActiveSheet.Cells(V, 2) = "10013208274380"
       objExcel.ActiveSheet.Cells(V, 3) = QuinLetra
       objExcel.ActiveSheet.Cells(V, 4) = Format(Neto, "##,##0.00")
       objExcel.ActiveSheet.Cells(V, 5) = "D"
       objExcel.ActiveSheet.Cells(V, 1).Font.Bold = True
       objExcel.ActiveSheet.Cells(V, 2).Font.Bold = True
       objExcel.ActiveSheet.Cells(V, 3).Font.Bold = True
       objExcel.ActiveSheet.Cells(V, 4).Font.Bold = True
       objExcel.ActiveSheet.Cells(V, 5).Font.Bold = True
       
        objExcel.ActiveSheet.Columns("A").ColumnWidth = 35
        objExcel.ActiveSheet.Columns("A").Font.Size = 10
        objExcel.ActiveSheet.Columns("B").NumberFormat = "############"
        objExcel.ActiveSheet.Columns("B").ColumnWidth = 17
        objExcel.ActiveSheet.Columns("B").Font.Size = 10
        objExcel.ActiveSheet.Columns("B").HorizontalAlignment = xlHAlignCenter
        objExcel.ActiveSheet.Columns("C").ColumnWidth = 26
        objExcel.ActiveSheet.Columns("C").Font.Size = 10
        objExcel.ActiveSheet.Columns("C").HorizontalAlignment = xlHAlignCenter
        objExcel.ActiveSheet.Columns("D").ColumnWidth = 10
        objExcel.ActiveSheet.Columns("D").Font.Size = 10
        objExcel.ActiveSheet.Columns("D").HorizontalAlignment = xlHAlignCenter
        objExcel.ActiveSheet.Columns("E").ColumnWidth = 4
        objExcel.ActiveSheet.Columns("E").Font.Size = 10
        objExcel.ActiveSheet.Columns("E").HorizontalAlignment = xlHAlignCenter

 
 Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto

Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub SmartButton1_Click()
Dim sql As String, NumeroNomina As Integer

NumeroNomina = Me.TxtNumNom13.Text
FrmSalarioHistorial.TxtNumNom13.Text = Me.TxtNumNom13.Text
'FrmSalarioHistorial.Numero
sql = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres," & vbLf
sql = sql & "HistorialSalarioMes.Fechaini , HistorialSalarioMes.Fechafin, HistorialSalarioMes.Enero, HistorialSalarioMes.Febrero, HistorialSalarioMes.Marzo, " & vbLf
sql = sql & "HistorialSalarioMes.Abril , HistorialSalarioMes.Mayo, HistorialSalarioMes.Junio, HistorialSalarioMes.Julio, HistorialSalarioMes.Agosto, " & vbLf
sql = sql & "HistorialSalarioMes.Septiembre , HistorialSalarioMes.Octubre, HistorialSalarioMes.Noviembre, HistorialSalarioMes.Diciembre, " & vbLf
sql = sql & "HistorialSalarioMes.NumNomina " & vbLf
sql = sql & "FROM HistorialSalarioMes INNER JOIN" & vbLf
sql = sql & "Empleado ON HistorialSalarioMes.CodEmpleado = Empleado.CodEmpleado" & vbLf
sql = sql & "Where (HistorialSalarioMes.NumNomina = " & NumeroNomina & ") AND (HistorialSalarioMes.Tipo = 'Aguinaldo') ORDER BY Nombres"

FrmSalarioHistorial.AdoSalarios.RecordSource = sql
FrmSalarioHistorial.AdoSalarios.Refresh

FrmSalarioHistorial.TxtNumNom13.Text = Me.TxtNumNom13.Text
FrmSalarioHistorial.Dbgr13Mes.Columns(0).Width = 1100
FrmSalarioHistorial.Dbgr13Mes.Columns(1).Width = 3000
FrmSalarioHistorial.Dbgr13Mes.Columns(4).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(5).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(6).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(7).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(8).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(9).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(10).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(11).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(12).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(13).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(14).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(15).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(16).Visible = False
FrmSalarioHistorial.Dbgr13Mes.Columns(2).Visible = False
FrmSalarioHistorial.Dbgr13Mes.Columns(3).Visible = False
FrmSalarioHistorial.DBCNominas.Text = Me.DBCNominas.Text
FrmSalarioHistorial.TxtFFIN13.Value = Me.TxtFFIN13.Value
FrmSalarioHistorial.TxtFINI13.Value = Me.TxtFINI13.Value

FrmSalarioHistorial.SSTab1.TabEnabled(0) = False
FrmSalarioHistorial.SSTab1.Tab = 1
FrmSalarioHistorial.TxtFFIN13.Visible = True
FrmSalarioHistorial.TxtFINI13.Visible = True
FrmSalarioHistorial.ButtonImprimirAguinaldo.Visible = True
FrmSalarioHistorial.ButtonExcelAguinaldo.Visible = True
FrmSalarioHistorial.CmdPrnNomina.Visible = False


  FrmSalarioHistorial.Dbgr13Mes.Columns(2).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(3).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(4).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(5).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(6).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(7).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(8).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(9).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(10).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(11).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(12).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(13).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(14).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(15).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(16).NumberFormat = "##,##0.00"
FrmSalarioHistorial.Show 1
End Sub

Private Sub SmartButton2_Click()
Dim sql As String, NumeroNomina As Integer

NumeroNomina = Me.TxtNumNomVaca.Text

sql = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres," & vbLf
sql = sql & "HistorialSalarioMes.Fechaini , HistorialSalarioMes.Fechafin, HistorialSalarioMes.Enero, HistorialSalarioMes.Febrero, HistorialSalarioMes.Marzo, " & vbLf
sql = sql & "HistorialSalarioMes.Abril , HistorialSalarioMes.Mayo, HistorialSalarioMes.Junio, HistorialSalarioMes.Julio, HistorialSalarioMes.Agosto, " & vbLf
sql = sql & "HistorialSalarioMes.Septiembre , HistorialSalarioMes.Octubre, HistorialSalarioMes.Noviembre, HistorialSalarioMes.Diciembre, " & vbLf
sql = sql & "HistorialSalarioMes.NumNomina " & vbLf
sql = sql & "FROM HistorialSalarioMes INNER JOIN" & vbLf
sql = sql & "Empleado ON HistorialSalarioMes.CodEmpleado = Empleado.CodEmpleado" & vbLf
sql = sql & "Where (HistorialSalarioMes.NumNomina = " & NumeroNomina & ")AND (HistorialSalarioMes.Tipo = N'Vacaciones') ORDER BY Empleado.CodEmpleado1"

FrmSalarioHistorial.AdoSalarioVacaciones.RecordSource = sql
'InputBox "", "", FrmSalarioHistorial.AdoSalarioVacaciones.RecordSource
FrmSalarioHistorial.AdoSalarioVacaciones.Refresh
FrmSalarioHistorial.SSTab1.TabEnabled(1) = False
FrmSalarioHistorial.SSTab1.Tab = 0

FrmSalarioHistorial.TxtFINIVaca.Value = Me.TxtFINIVaca.Value
FrmSalarioHistorial.TxtFFinVaca.Value = Me.TxtFFinVaca.Value
FrmSalarioHistorial.TxtFINI13.Value = Me.TxtFINIVaca.Value
FrmSalarioHistorial.TxtFFIN13.Value = Me.TxtFFIN13.Value
FrmSalarioHistorial.TxtNumNom13.Text = Me.TxtNumNomVaca.Text
FrmSalarioHistorial.TxtNumNomVaca.Text = Me.TxtNumNomVaca.Text
FrmSalarioHistorial.DbgrVacaciones.Columns(0).Width = 1100
FrmSalarioHistorial.DbgrVacaciones.Columns(1).Width = 3000
FrmSalarioHistorial.DbgrVacaciones.Columns(4).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(5).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(6).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(7).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(8).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(9).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(10).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(11).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(12).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(13).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(14).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(15).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(16).Visible = False
FrmSalarioHistorial.DbgrVacaciones.Columns(2).Visible = False
FrmSalarioHistorial.DbgrVacaciones.Columns(3).Visible = False
FrmSalarioHistorial.DbgrVacaciones.Columns(4).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(5).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(6).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(7).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(8).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(9).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(10).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(11).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(12).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(13).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(14).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(15).NumberFormat = "##,##0.00"


FrmSalarioHistorial.DBCNominas.Text = Me.DBCNominas.Text
FrmSalarioHistorial.TxtFFinVaca.Value = Me.TxtFFinVaca.Value
FrmSalarioHistorial.TxtFINIVaca.Value = Me.TxtFINIVaca.Value
FrmSalarioHistorial.TxtNumNomVaca.Text = Me.TxtNumNomVaca.Text

FrmSalarioHistorial.ButtonImprimirAguinaldo.Visible = False
FrmSalarioHistorial.ButtonExcelAguinaldo.Visible = False
FrmSalarioHistorial.TipoNomina = "Vacaciones"
'FrmSalarioHistorial.CmdPrnNomina.Visible = False
FrmSalarioHistorial.Show 1
End Sub

Private Sub TxtFFIN13_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String

Mes1 = Month(Me.TxtFINIVaca.Value)
Año1 = Year(Me.TxtFINIVaca.Value)
Mes2 = Month(Me.TxtFFinVaca.Value)
Año2 = Year(Me.TxtFFinVaca.Value)

Mes1 = Format(Mes1, "0#")
Mes2 = Format(Mes2, "0#")

DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.DtpFin13vo.Value = Me.AdoBusca.Recordset("Inicio")
 End If

Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.DtpFin13vo.Value = Me.AdoBusca.Recordset("Final")
 End If
End Sub

Private Sub TxtFFinVaca_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String

Mes1 = Month(Me.TxtFINIVaca.Value)
Año1 = Year(Me.TxtFINIVaca.Value)
Mes2 = Month(Me.TxtFFinVaca.Value)
Año2 = Year(Me.TxtFFinVaca.Value)

Mes1 = Format(Mes1, "0#")
Mes2 = Format(Mes2, "0#")

DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.dtpFPInicio.Value = Me.AdoBusca.Recordset("Inicio")
 End If

Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.dtpFPFinal.Value = Me.AdoBusca.Recordset("Final")
 End If
End Sub

Private Sub TxtFINI13_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String

Mes1 = Month(Me.TxtFINI13.Value)
Año1 = Year(Me.TxtFINI13.Value)
Mes2 = Month(Me.TxtFFIN13.Value)
Año2 = Year(Me.TxtFFIN13.Value)

Mes1 = Format(Mes1, "0#")
Mes2 = Format(Mes2, "0#")

DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.DtpInicio13vo.Value = Me.AdoBusca.Recordset("Inicio")
 End If

Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.DtpFin13vo.Value = Me.AdoBusca.Recordset("Final")
 End If
End Sub

Private Sub TxtFINIVaca_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String

Mes1 = Month(Me.TxtFINIVaca.Value)
Año1 = Year(Me.TxtFINIVaca.Value)
Mes2 = Month(Me.TxtFFinVaca.Value)
Año2 = Year(Me.TxtFFinVaca.Value)

Mes1 = Format(Mes1, "0#")
Mes2 = Format(Mes2, "0#")

DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.dtpFPInicio.Value = Me.AdoBusca.Recordset("Inicio")
 End If

Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.dtpFPFinal.Value = Me.AdoBusca.Recordset("Final")
 End If
End Sub
