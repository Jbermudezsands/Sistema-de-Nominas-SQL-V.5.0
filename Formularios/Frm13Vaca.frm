VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form Frm13Vaca 
   Caption         =   "Calculo del 13vo mes y Vacaciones"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   ScaleHeight     =   572
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   839
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   13095
      TabIndex        =   54
      Top             =   0
      Width           =   13095
      Begin VB.Image Image1 
         Height          =   1320
         Left            =   0
         Picture         =   "Frm13Vaca.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1965
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   13080
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Vacaciones Mensuales y 13vo Mes"
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
         Left            =   4200
         TabIndex        =   55
         Top             =   360
         Width           =   5040
      End
   End
   Begin MSAdodcLib.Adodc AdoHistorialSalarial 
      Height          =   375
      Left            =   9600
      Top             =   9600
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
      Left            =   7440
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
      Left            =   7440
      Top             =   10440
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
      Left            =   9720
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
      Left            =   4200
      Top             =   10560
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
      Left            =   7440
      Top             =   10440
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
      Left            =   4200
      Top             =   10200
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
      Left            =   4200
      Top             =   10560
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
      Left            =   1800
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
      Left            =   1800
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
      Left            =   5640
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
      Left            =   360
      Top             =   10440
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
      Left            =   360
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
      Left            =   1320
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
      Left            =   1320
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
      Left            =   840
      Top             =   10440
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
      Left            =   0
      ScaleHeight     =   7035
      ScaleWidth      =   12795
      TabIndex        =   0
      Top             =   1320
      Width           =   12855
      Begin VB.TextBox txtAntiguedad 
         Height          =   285
         Left            =   9720
         TabIndex        =   43
         Text            =   "30"
         Top             =   240
         Width           =   615
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo DBCNominas 
         Bindings        =   "Frm13Vaca.frx":8AB8
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
         DownPicture     =   "Frm13Vaca.frx":8AD4
         Height          =   375
         Left            =   10920
         Picture         =   "Frm13Vaca.frx":A5B6
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
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   12632319
         TabCaption(0)   =   "Vacaciones"
         TabPicture(0)   =   "Frm13Vaca.frx":C098
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "ChkExtraVaca"
         Tab(0).Control(1)=   "CmdCalVacaMes"
         Tab(0).Control(2)=   "ChkEliminar"
         Tab(0).Control(3)=   "ChkRestar"
         Tab(0).Control(4)=   "CHKTranferir"
         Tab(0).Control(5)=   "CmdNominaVaca"
         Tab(0).Control(6)=   "CmdColillaVaca"
         Tab(0).Control(7)=   "CmdMonedasvaca"
         Tab(0).Control(8)=   "CmdExportar"
         Tab(0).Control(9)=   "DbgrVacaciones"
         Tab(0).Control(10)=   "Command2"
         Tab(0).Control(11)=   "TxtFFinVaca"
         Tab(0).Control(12)=   "TxtFINIVaca"
         Tab(0).Control(13)=   "CmdPRVaca"
         Tab(0).Control(14)=   "TxtDiasDescuento"
         Tab(0).Control(15)=   "CmdCerrarVacaciones"
         Tab(0).Control(16)=   "TxtNumNomVaca"
         Tab(0).Control(17)=   "CmdCalVaca"
         Tab(0).Control(18)=   "SmartButton2"
         Tab(0).Control(19)=   "dtpFPFinal"
         Tab(0).Control(20)=   "dtpFPInicio"
         Tab(0).Control(21)=   "TxtFechaAplica"
         Tab(0).Control(22)=   "PBVacaciones"
         Tab(0).Control(23)=   "Label13"
         Tab(0).Control(24)=   "Label7"
         Tab(0).Control(25)=   "Label1"
         Tab(0).Control(26)=   "Label3"
         Tab(0).Control(27)=   "Label4"
         Tab(0).ControlCount=   28
         TabCaption(1)   =   "Trecavo Mes"
         TabPicture(1)   =   "Frm13Vaca.frx":C0B4
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label6"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label5"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label2"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "PB13Mes"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "CmdprNomina"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "CmdPrnNomina"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "CmdCal13"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "TxtNumNom13"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "CmdCerrar13"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "TxtFINI13"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "TxtFFIN13"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Command1"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "CmdExporta2"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "Dbgr13Mes"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "SmartButton1"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "CmdDenominacion"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "Frame1"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "CmdExportaBAC"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "ChkExtra"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "ChkEliminarAguinaldo"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).ControlCount=   20
         Begin VB.CheckBox ChkExtraVaca 
            Caption         =   "Calcular Horas Extra"
            Height          =   255
            Left            =   -66720
            TabIndex        =   60
            Top             =   960
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox ChkEliminarAguinaldo 
            Caption         =   "Eliminar el Calculo Anterior"
            Height          =   375
            Left            =   120
            TabIndex        =   59
            Top             =   1800
            Width           =   1695
         End
         Begin VB.CheckBox ChkExtra 
            Caption         =   "Calcular Horas Extra"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CommandButton CmdCalVacaMes 
            DownPicture     =   "Frm13Vaca.frx":C0D0
            Height          =   375
            Left            =   -67080
            Picture         =   "Frm13Vaca.frx":DBB2
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   4440
            UseMaskColor    =   -1  'True
            Width           =   1455
         End
         Begin VB.CheckBox ChkEliminar 
            Caption         =   "Eliminar el Calculo Anterior"
            Height          =   255
            Left            =   -66720
            TabIndex        =   56
            Top             =   600
            Width           =   2175
         End
         Begin VB.CheckBox ChkRestar 
            Caption         =   "Restar Inss Nomina"
            Height          =   255
            Left            =   -74760
            TabIndex        =   51
            Top             =   3240
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CommandButton CmdExportaBAC 
            DownPicture     =   "Frm13Vaca.frx":F6F4
            Enabled         =   0   'False
            Height          =   375
            Left            =   6480
            Picture         =   "Frm13Vaca.frx":111D6
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   4560
            Width           =   1455
         End
         Begin VB.Frame Frame1 
            Caption         =   "Fechas Segun Periodos"
            Height          =   855
            Left            =   4320
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
               Format          =   81592321
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
               Format          =   81592321
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
            Left            =   -74760
            TabIndex        =   39
            Top             =   2880
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CommandButton CmdNominaVaca 
            DownPicture     =   "Frm13Vaca.frx":12A80
            Enabled         =   0   'False
            Height          =   375
            Left            =   -65640
            Picture         =   "Frm13Vaca.frx":14562
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton CmdColillaVaca 
            DownPicture     =   "Frm13Vaca.frx":16044
            Enabled         =   0   'False
            Height          =   375
            Left            =   -67080
            Picture         =   "Frm13Vaca.frx":17B26
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton CmdMonedasvaca 
            DownPicture     =   "Frm13Vaca.frx":19608
            Enabled         =   0   'False
            Height          =   375
            Left            =   -68520
            Picture         =   "Frm13Vaca.frx":1B0EA
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton CmdExportar 
            DownPicture     =   "Frm13Vaca.frx":1C9EC
            Enabled         =   0   'False
            Height          =   375
            Left            =   -64200
            Picture         =   "Frm13Vaca.frx":1E4CE
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton CmdDenominacion 
            DownPicture     =   "Frm13Vaca.frx":1FD78
            Enabled         =   0   'False
            Height          =   375
            Left            =   6480
            Picture         =   "Frm13Vaca.frx":2185A
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   4920
            Width           =   1455
         End
         Begin SmartButtonProject.SmartButton SmartButton1 
            Height          =   975
            Left            =   480
            TabIndex        =   32
            Top             =   3120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1720
            Caption         =   "Historial Manual"
            Picture         =   "Frm13Vaca.frx":2315C
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
            Bindings        =   "Frm13Vaca.frx":23A36
            Height          =   3015
            Left            =   2040
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
            Bindings        =   "Frm13Vaca.frx":23A4F
            Height          =   2535
            Left            =   -72840
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
            DownPicture     =   "Frm13Vaca.frx":23A6B
            Enabled         =   0   'False
            Height          =   375
            Left            =   10800
            Picture         =   "Frm13Vaca.frx":2554D
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            DownPicture     =   "Frm13Vaca.frx":26DF7
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
            Picture         =   "Frm13Vaca.frx":288D9
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   4560
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            DownPicture     =   "Frm13Vaca.frx":2A3BB
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
            Picture         =   "Frm13Vaca.frx":2BE9D
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   4440
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker TxtFFIN13 
            Height          =   315
            Left            =   2280
            TabIndex        =   25
            Top             =   840
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   81592321
            CurrentDate     =   38305
         End
         Begin MSComCtl2.DTPicker TxtFINI13 
            Height          =   315
            Left            =   240
            TabIndex        =   24
            Top             =   840
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   81592321
            CurrentDate     =   38305
         End
         Begin MSComCtl2.DTPicker TxtFFinVaca 
            Height          =   315
            Left            =   -72600
            TabIndex        =   23
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   81592321
            CurrentDate     =   38305
         End
         Begin MSComCtl2.DTPicker TxtFINIVaca 
            Height          =   315
            Left            =   -74760
            TabIndex        =   22
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   81592321
            CurrentDate     =   38305
         End
         Begin VB.CommandButton CmdPRVaca 
            DownPicture     =   "Frm13Vaca.frx":2D97F
            Enabled         =   0   'False
            Height          =   375
            Left            =   -72840
            Picture         =   "Frm13Vaca.frx":2F461
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   4320
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtDiasDescuento 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   -73440
            TabIndex        =   11
            Text            =   "0"
            Top             =   2520
            Width           =   375
         End
         Begin VB.CommandButton CmdCerrar13 
            DownPicture     =   "Frm13Vaca.frx":30F43
            Height          =   375
            Left            =   9360
            Picture         =   "Frm13Vaca.frx":32A25
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   4560
            Width           =   1455
         End
         Begin VB.CommandButton CmdCerrarVacaciones 
            DownPicture     =   "Frm13Vaca.frx":34327
            Height          =   375
            Left            =   -65640
            Picture         =   "Frm13Vaca.frx":35E09
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
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   2400
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
            Left            =   -74520
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton CmdCal13 
            DownPicture     =   "Frm13Vaca.frx":3770B
            Height          =   375
            Left            =   7920
            Picture         =   "Frm13Vaca.frx":391ED
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   4560
            Width           =   1455
         End
         Begin VB.CommandButton CmdCalVaca 
            DownPicture     =   "Frm13Vaca.frx":3AD2F
            Height          =   375
            Left            =   -67080
            Picture         =   "Frm13Vaca.frx":3C811
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   4440
            UseMaskColor    =   -1  'True
            Width           =   1455
         End
         Begin VB.CommandButton CmdPrnNomina 
            DownPicture     =   "Frm13Vaca.frx":3E353
            Enabled         =   0   'False
            Height          =   375
            Left            =   7920
            Picture         =   "Frm13Vaca.frx":3FE35
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton CmdprNomina 
            DownPicture     =   "Frm13Vaca.frx":41917
            Enabled         =   0   'False
            Height          =   375
            Left            =   9360
            Picture         =   "Frm13Vaca.frx":433F9
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   4920
            Width           =   1455
         End
         Begin SmartButtonProject.SmartButton SmartButton2 
            Height          =   975
            Left            =   -74640
            TabIndex        =   33
            Top             =   3600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1720
            Caption         =   "Historial Manual"
            Picture         =   "Frm13Vaca.frx":44EDB
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
            Left            =   -72600
            TabIndex        =   40
            Top             =   1200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   81592321
            CurrentDate     =   38305
         End
         Begin MSComCtl2.DTPicker dtpFPInicio 
            Height          =   315
            Left            =   -74760
            TabIndex        =   41
            Top             =   1200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   81592321
            CurrentDate     =   38305
         End
         Begin MSComCtl2.DTPicker TxtFechaAplica 
            Height          =   315
            Left            =   -69120
            TabIndex        =   53
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Format          =   81592321
            CurrentDate     =   38305
         End
         Begin XtremeSuiteControls.ProgressBar PBVacaciones 
            Height          =   375
            Left            =   -74880
            TabIndex        =   61
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
            Left            =   120
            TabIndex        =   62
            Top             =   4680
            Width           =   6255
            _Version        =   786432
            _ExtentX        =   11033
            _ExtentY        =   661
            _StockProps     =   93
            BackColor       =   14737632
            Scrolling       =   1
            Appearance      =   6
         End
         Begin VB.Label Label13 
            Caption         =   "Fecha Aplica"
            Height          =   255
            Left            =   -70200
            TabIndex        =   52
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Días de Descuento"
            Height          =   255
            Left            =   -74880
            TabIndex        =   19
            Top             =   2520
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
            Left            =   240
            TabIndex        =   18
            Top             =   2400
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
            Left            =   -74760
            TabIndex        =   17
            Top             =   1920
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
            Left            =   120
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
            Left            =   2160
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
            Left            =   -74880
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
            Left            =   -72720
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
      Left            =   4200
      Top             =   10200
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
Attribute VB_Name = "Frm13Vaca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
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
Dim i As Integer, CantRegistros As Integer
Dim Dias As Double, annos As Double
Dim Adelanto13vo As Double, MontoSubsidio As Double
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

If Me.ChkEliminarAguinaldo.Value = 1 Then
 Fecha1 = Format(Me.TxtFINIVaca.Value, "yyyy-mm-dd")
 Fecha2 = Format(Me.TxtFFinVaca.Value, "yyyy-mm-dd")
 Set Ejecutar = New ADODB.Connection
 Ejecutar.ConnectionString = Conexion
 Ejecutar.Open
 Ejecutar.Execute "DELETE FROM DetalleNom13Mes WHERE (NumNom13Mes =  " & NumNom13Mes & ")"
' Ejecutar.Execute "DELETE FROM HistorialSalarioMes WHERE (FechaIni = CONVERT(DATETIME, '" & Fecha1 & "', 102)) AND (FechaFin = CONVERT(DATETIME, '" & Fecha2 & "', 102))"
End If


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
SqlEmpleados = "SELECT Empleado.SalarioFijo, Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, Empleado.ExentoIr, Empleado.OtrosIngresos,  Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, Empleado.Ausente From Empleado WHERE Empleado.CodTipoNomina= '" & CodTipoNomina & "' AND Empleado.Activo=1"
DtaEmpleados.RecordSource = SqlEmpleados
DtaEmpleados.Refresh
If Not Me.DtaEmpleados.Recordset.EOF Then
DtaEmpleados.Recordset.MoveLast
CantEmpleados = DtaEmpleados.Recordset.RecordCount
End If

With PB13Mes
.Min = 0
.Value = 0
.Max = CantEmpleados

i = 1
DtaEmpleados.Refresh
'recorro la BD empleados y a cada uno le busco su salario mayor (sies destajo) si no solo extraigo su salario
Do While Not DtaEmpleados.Recordset.EOF




Dias = 0
  CodEmpleado = DtaEmpleados.Recordset("CodEmpleado")



  
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
     NumFecha2 = FechaContrato
     NumFecha1 = Me.TxtFFIN13.Value
     annos = (CDbl(NumFecha1) - CDbl(NumFecha2) + 1) / 365
     CantMeses = annos * 12
     
     If CantMeses < 0 Then
      Dias = 0
     ElseIf CantMeses <= 12 Then
'      Dias = Format(CantMeses * (DiasMes / 12), "##,##0.0000")
      Dias = (CDbl(NumFecha1) - CDbl(NumFecha2) + 1) * 0.0833333
     Else
'      NumFecha2 = CDate("01/12/" & Anno - 1)
'      annos = (CDbl(NumFecha1) - CDbl(NumFecha2)) / 365
'      CantMeses = Format(annos * 12, "##,##0.0000")
'      If CantMeses > 12 Then
'       Dias = 12 * 2.5346
'      Else
'       Dias = CantMeses * 2.5346
      Dias = DiasMes
'      End If
     End If
     
     If Dias > 30 Then
        Dias = 30
     End If
     
     
    End If
   End If
     If CodEmpleado = 10329 Then
         CodEmpleado = 10329
     End If
       
       
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

If Dias >= Me.txtAntiguedad.Text Then

 If DtaEmpleados.Recordset("SalarioFijo") = "N" Then
        If MesActual > 6 Then
            Meses = MesActual - 5
        Else
           Meses = 1
        End If
        For Mes = Meses To MesActual
        If Me.ChkExtra.Value = 1 Then
         '////////////SI TIENE MARCADO LAS HORAS SE LAS SUMO EN EL AGUINALDO///////////////////
         SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.BonoProduccion,DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasExtras + DetalleNomina.Incentivos + DetalleNomina.Antiguedad AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where (Nomina.Mes =" & Mes & ") And (Nomina.Ano =" & Anno & ") and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
        Else
         SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.BonoProduccion,DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Incentivos + DetalleNomina.Antiguedad AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where (Nomina.Mes =" & Mes & ") And (Nomina.Ano =" & Anno & ") and DetalleNomina.CodEmpleado = '" & CodEmpleado & "'"
        End If

            DtaNominas.RecordSource = SqlNominas
            DtaNominas.Refresh
               SalTemp = 0
              Do While Not DtaNominas.Recordset.EOF
               SalTemp = SalTemp + DtaNominas.Recordset("Total")
               CantRegistros = CantRegistros + 1
               DtaNominas.Recordset.MoveNext
              Loop
              

              
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
                 Me.AdoHistorialSalarial.RecordSource = "SELECT NumNomina, CodEmpleado, FechaIni, FechaFin, Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, Agosto, Septiembre, Octubre, Noviembre, Diciembre From HistorialSalarioMes WHERE     (CodEmpleado = '" & CodEmpleado & "') AND (FechaIni = CONVERT(DATETIME, '" & Fecha1 & "', 102)) AND (FechaFin = CONVERT(DATETIME, '" & Fecha2 & "',102))"
                 Me.AdoHistorialSalarial.Refresh
                  If Me.AdoHistorialSalarial.Recordset.EOF Then
                        Me.AdoHistorialSalarial.Recordset.AddNew
                        Me.AdoHistorialSalarial.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
                        Me.AdoHistorialSalarial.Recordset("FechaIni") = CDate(Me.TxtFINI13.Value)
                        Me.AdoHistorialSalarial.Recordset("FechaFin") = CDate(Me.TxtFFIN13.Value)
                        Me.AdoHistorialSalarial.Recordset("NumNomina") = val(Me.TxtNumNom13.Text)

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
                 Me.AdoHistorialSalarial.RecordSource = "SELECT NumNomina, CodEmpleado, FechaIni, FechaFin, Enero, Febrero, Marzo, Abril, Mayo, Junio, Julio, Agosto, Septiembre, Octubre, Noviembre, Diciembre From HistorialSalarioMes WHERE  (CodEmpleado = '" & CodEmpleado & "')"
                 Me.AdoHistorialSalarial.Refresh
                  If Not Me.AdoHistorialSalarial.Recordset.EOF Then
 
                        Select Case Mes
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
        If SalMayor < SalTemp Then
           SalMayor = SalTemp
        End If
  
       Next
   Else
   '///////////Si es Salario Fijo se lo Calculo//////////////
       'SqlNominas = "SELECT Nomina.NumNomina, Nomina.FechaNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[Comisiones]+[DetalleNomina].[Incentivos] AS Total, Month([Nomina].[FechaNomina]) AS Mes, Year([Nomina].[FechaNomina]) AS Anno FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina Where (((DetalleNomina.CodEmpleado) = '" & CodEmpleado & "'))"
       'DtaNominas.RecordSource = SqlNominas
       'DtaNominas.Refresh
       'If Not DtaNominas.Recordset.EOF Then
       'DtaNominas.Recordset.MoveLast
     SalMayor = DtaEmpleados.Recordset("SueldoPeriodo")
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
     End If
    
    
   End If
        
    
       
DtaNom13Mes.Refresh
Do While Not DtaNom13Mes.Recordset.EOF
    If DtaNom13Mes.Recordset("NumNom13Mes") = val(TxtNumNom13.Text) And DtaNom13Mes.Recordset("Activa") = True Then
       'DtaNom13Mes.Recordset.Edit
       DtaNom13Mes.Recordset("montopagado") = DtaNom13Mes.Recordset("montopagado") + SalMayor
       DtaNom13Mes.Recordset.Update
   Exit Do
    End If
DtaNom13Mes.Recordset.MoveNext
Loop

       Me.DtaDetalleNom13Mes.RecordSource = "SELECT Id, NumNom13Mes, CodEmpleado, SalarioMensual, SalarioAPagar, DiasAPagar, Adelanto13vo From DetalleNom13Mes"
       Me.DtaDetalleNom13Mes.Refresh
       If Me.DtaDetalleNom13Mes.Recordset.EOF Then
       Id = 1
       Else
       Me.DtaDetalleNom13Mes.Recordset.MoveLast
       Id = Me.DtaDetalleNom13Mes.Recordset("id") + 1
       End If
       
        DtaDetalleNom13Mes.Recordset.AddNew
        Me.DtaDetalleNom13Mes.Recordset("id") = Id
        DtaDetalleNom13Mes.Recordset("Adelanto13vo") = Adelanto13vo
        DtaDetalleNom13Mes.Recordset("NumNom13Mes") = val(TxtNumNom13.Text)
        DtaDetalleNom13Mes.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
        DtaDetalleNom13Mes.Recordset("SalarioMensual") = SalMayor
        If (Not Dias = 0) And Dias < 30 Then
         DtaDetalleNom13Mes.Recordset("SalarioAPagar") = ((SalMayor * Dias) / DiasMes)
        ElseIf Dias >= 30 Then
         DtaDetalleNom13Mes.Recordset("SalarioAPagar") = (SalMayor)
        Else
         DtaDetalleNom13Mes.Recordset("SalarioAPagar") = 0
        End If
        DtaDetalleNom13Mes.Recordset("DiasAPagar") = Dias
        DtaDetalleNom13Mes.Recordset.Update
        


       



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

Private Sub CmdCalVaca_Click()
Dim SalarioBasico As Double, DiasDescuento As Double
Dim SqlEmpleados As String, MontoSubsidio As Double
Dim CodTipoNomina As String, TotalSubsidio As Double
Dim SqlNominas As String, NumVaca As Integer
Dim SalMayor As Double, TarifaHoraria As Double
Dim SalTemp As Double, Dias As Double
Dim CodEmpleado As String, AdelantoVaca As Double
Dim Edicion As Boolean
Dim Anno As Integer
Dim Mes As Integer, Fecha As Date
Dim CantEmpleados As Long
Dim i As Integer, CantMeses As Integer, CantRegistros As Integer
Dim DiasMes As Double
Dim DiasSemana As Double, DiasPagar As Double
Dim FechaHoy As Date
Dim rsDB As New ADODB.Recordset
Dim cnDB As New ADODB.Connection
Dim iMes As Integer

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

NumVaca = val(TxtNumNomVaca.Text)

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

SqlEmpleados = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, " & _
               "Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, " & _
               "Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, " & _
               "Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, " & _
               "Empleado.ExentoIr, Empleado.OtrosIngresos, Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, " & _
               "Empleado.Ausente, Empleado.SalarioFijo, Empleado.CodEmpleado1, Historico.FechaContrato " & _
               "FROM  Empleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado " & _
               "WHERE  (Empleado.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (Empleado.Ausente = 0)"
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
            'tengo que hacer un SQL de Solo los que esten en el rango de fechas
            'solo se veran las nominas de cada mes
            'se deben de hacer ciclos por cada mes, seis ciclos por los seis meses.
         SalMayor = 0
         CantMeses = 0
         CantRegistros = 0
         
        
          
         
         
          '/////////////////ESTE ES EL SALARIO BASICO, SIN PRODUCCION////////////////
             TarifaHoraria = DtaEmpleados.Recordset("TarifaHoraria")
             SalarioBasico = Format(DiasMes * 8 * TarifaHoraria, "##,##0.00")
        
        
          DtaHistorico.RecordSource = "SELECT Historico.Codempleado, Historico.FechaBaja, Historico.FechaContrato From Historico Where (((Historico.CodEmpleado) = '" & CodEmpleado & "'))"
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
                 
             Dias = (CDate(Me.TxtFFinVaca.Value) - CDate(FechaContrato)) + 1
        '     Dias = (CDate(Me.dtpFPFinal.Value) - CDate(FechaContrato)) + 1
                 
             If Dias < 0 Then
               Dias = 0
             ElseIf Dias > 182 Then
               Dias = (CDate(Me.TxtFFinVaca.Value) - CDate(Me.TxtFINIVaca.Value)) + 1
        
             End If
            End If
           End If
           
           
         
         
         If Dias > CInt(Me.txtAntiguedad.Text) Then
         
             If DtaEmpleados.Recordset("SalarioFijo") = "N" Then
                 Año = Year(Me.TxtFFinVaca.Value)
         '///////////Si el Salario es Variable Busco el Salario Mayor/////////
              
                '/////////////CAlculo las vacaciones Enero - Junio ////////////////////////////
                SalMayor = 0
                
                
                
                
                 Do While Not rsDB.EOF And Dias > CInt(Me.txtAntiguedad.Text)  '30   '    For Mes = 1 To 6
              
                       SqlNominas = "SELECT DISTINCT " & _
                                    "DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, " & _
                                    "SUM(DetalleNomina.SeptimoDia) AS Septimo, SUM(DetalleNomina.OtrosIngresos) AS Otros, SUM(DetalleNomina.Incentivos) AS Incentivos, " & _
                                    "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.SeptimoDia + DetalleNomina.OtrosIngresos + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion) AS TotalIngresos, " & _
                                    "MIN(Nomina.FechaNominaINI) AS FechaInicio, MAX(Nomina.FechaNomina) AS FechaFin, Nomina.Mes, Nomina.Ano AS AÑO " & _
                                    "FROM  DetalleNomina INNER JOIN " & _
                                    "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina " & _
                                    "GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano " & _
                                    "HAVING      (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo) <> 0) AND (DetalleNomina.CodEmpleado = " & CodEmpleado & ") AND (Nomina.Mes = " & CInt(rsDB.Fields("mes")) & ") AND " & _
                                    "(Nomina.Ano = " & rsDB.Fields("año") & ") ORDER BY Nomina.Ano, Nomina.Mes "
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
                
            If SalarioBasico > SalMayor Then
             SalMayor = SalarioBasico
           End If
                
                DiasDescuento = 0
          If Edicion = True Then
          
          
        
           
           Me.DtaDetalleNomVaca.RecordSource = "SELECT DetalleNomVaca.Id, DetalleNomVaca.TotalDevengado, DetalleNomVaca.Inss, DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones From DetalleNomVaca Where (((DetalleNomVaca.NumNomVaca) = " & val(TxtNumNomVaca.Text) & ") And ((DetalleNomVaca.CodEmpleado) = '" & CodEmpleado & "'))"
           Me.DtaDetalleNomVaca.Refresh
             
           '///////Busco si el Empleado ya Existe en la Nomina de Vacaciones/////
             If Not Me.DtaDetalleNomVaca.Recordset.EOF Then
                DiasDescuento = DtaDetalleNomVaca.Recordset("DiasDescuento") + val(TxtDiasDescuento)
                'DtaDetalleNomVaca.Recordset.Edit
                DtaDetalleNomVaca.Recordset("NumNomVaca") = val(TxtNumNomVaca.Text)
                DtaDetalleNomVaca.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
        '        If Me.DBCNominas.Text = "Administracion" Then
                
                  If DtaEmpleados.Recordset("SalarioFijo") = "S" Then
                   
                         
        '         Else
                    
                    If Dias = 182 Then
                       DtaDetalleNomVaca.Recordset("Inss") = SalMayor * 0.0625
                       DtaDetalleNomVaca.Recordset("TotalDevengado") = SalMayor
                    Else
                       DtaDetalleNomVaca.Recordset("Inss") = (Dias * (SalMayor * 2) / DiasMes) * 0.08333333 * 0.0625
                       DtaDetalleNomVaca.Recordset("TotalDevengado") = (Dias * (SalMayor * 2) / DiasMes) * 0.08333333
                    End If
                      
                     DtaDetalleNomVaca.Recordset("SalarioMensual") = SalMayor * 2
                      
        '          End If
                  
                Else
                    
                    DtaDetalleNomVaca.Recordset("SalarioMensual") = SalMayor
                    DtaDetalleNomVaca.Recordset("Inss") = ((SalMayor / DiasMes) * ((Dias - DiasDescuento) * 0.0833333) - AdelantoVaca) * 0.0625
                    DtaDetalleNomVaca.Recordset("TotalDevengado") = ((SalMayor / DiasMes) * (Dias - DiasDescuento) * 0.0833333)
                    
                    
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
               Me.DtaDetalleNomVaca.RecordSource = "SELECT DetalleNomVaca.Id, DetalleNomVaca.Inss, DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones From DetalleNomVaca"
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
                DtaDetalleNomVaca.Recordset("Inss") = (SalMayor * (((Dias - DiasDescuento) * 0.0833333) / DiasMes) - AdelantoVaca) * 0.0625
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
           
               Me.DtaDetalleNomVaca.RecordSource = "SELECT DetalleNomVaca.Id, DetalleNomVaca.Inss, DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones From DetalleNomVaca"
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
                DtaDetalleNomVaca.Recordset("Inss") = (SalMayor * (((Dias - DiasDescuento) * 0.0833333) / DiasMes) - AdelantoVaca) * 0.0625
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

Private Sub CmdCalVacaMes_Click()
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
Dim DiasMenos As Double


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
                  
                  For i = 1 To 6

                   FechaInicioVaca = DateSerial(Year(Me.TxtFFinVaca.Value), Month(Me.TxtFFinVaca.Value) - i, 1)
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
                             
                           
                                                     
                rs.Open "DELETE FROM HistorialSalarioMes WHERE (CodEmpleado = " & CodEmpleado & ") AND (FechaIni = CONVERT(DATETIME, '" & Format(CDate(Me.dtpFPInicio), "yyyy-mm-dd") & "', 102)) AND (FechaFin = CONVERT(DATETIME, '" & Format(CDate(Me.dtpFPFinal), "yyyy-mm-dd") & "', 102)), conexion"

                             
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
                                 SalarioAcumulado = Format(Me.AdoBusca.Recordset("TotalIngresos"), "####0.00") + ((SalMayor / DiasMes) * (Dias / 12)) + Monto - InssAcumulado - (((SalMayor / DiasMes) * (Dias / 12)) + Monto) * 0.0625
                                 
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
                                       DtaDetalleNomVaca.Recordset("Inss") = SalMayor * 0.0625
                                       DtaDetalleNomVaca.Recordset("TotalDevengado") = SalMayor
                                    Else
                                       DtaDetalleNomVaca.Recordset("Inss") = (Dias * (SalMayor * 2) / DiasMes) * 0.08333333 * 0.0625
                                       DtaDetalleNomVaca.Recordset("TotalDevengado") = (Dias * (SalMayor * 2) / DiasMes) * 0.08333333
                                    End If
                                      
                                     DtaDetalleNomVaca.Recordset("SalarioMensual") = SalMayor * 2
                                      
                        '          End If
                                  
                                Else
                                    
                                    DtaDetalleNomVaca.Recordset("SalarioMensual") = SalMayor
'                                    DtaDetalleNomVaca.Recordset("Inss") = ((SalMayor / DiasMes) * ((Dias - DiasDescuento) * 0.0833333) - AdelantoVaca) * 0.0625
                                    DtaDetalleNomVaca.Recordset("Inss") = ((SalMayor / DiasMes) * ((Dias * 0.0833333) - DiasDescuento)) * 0.0625
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
                               
                               
                               '/////////////////////////////////////////////////////////////////////////////////////////////////
                               '///////////////////////BUSCO LOS DIAS DE VACACIONES /////////////////////////////////////////////
                               '/////////////////////////////////////////////////////////////////////////////////////////////////
                               MDIPrimero.AdoConsulta.RecordSource = "SELECT CodigoEmpleado, SUM(DiasDisfrutar) AS Dias From SolicitudVacaciones WHERE (TipoSolicitud = 'Vacaciones') AND (Anulado = 0) AND (FechaInicio >= CONVERT(DATETIME, '" & Format(Me.dtpFPInicio.Value, "yyyy-MM-dd") & "', 102)) AND (FechaFin <= CONVERT(DATETIME,'" & Format(Me.dtpFPFinal.Value, "yyyy-MM-dd") & "', 102)) GROUP BY CodigoEmpleado, TipoSolicitud HAVING  (CodigoEmpleado = " & DtaEmpleados.Recordset("CodEmpleado") & ")"
                               MDIPrimero.AdoConsulta.Refresh
                               If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
                                    DiasMenos = MDIPrimero.AdoConsulta.Recordset("Dias")
                               Else
                                    DiasMenos = 0
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
                                'DtaDetalleNomVaca.Recordset("Inss") = (SalMayor * (((Dias - DiasDescuento) * 0.0833333) / DiasMes) - AdelantoVaca) * 0.0625
                                DtaDetalleNomVaca.Recordset("Ir") = IrAcumulado
                                DtaDetalleNomVaca.Recordset("Inss") = (SalMayor * (((Dias * 0.0833333) - DiasDescuento) / DiasMes)) * 0.0625
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
'                                DtaDetalleNomVaca.Recordset("Inss") = (SalMayor * (((Dias - DiasDescuento) * 0.0833333) / DiasMes) - AdelantoVaca) * 0.0625
                                DtaDetalleNomVaca.Recordset("Ir") = IrAcumulado
                                DtaDetalleNomVaca.Recordset("Inss") = (SalMayor * (((Dias * 0.0833333) - DiasDescuento) / DiasMes)) * 0.0625
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
                                                                     DtaConsulta.Recordset("Inss") = Monto * 0.0625
        '                                                             DtaConsulta.Recordset("Ir") = 0
                                                                     DtaConsulta.Recordset("TotalDevengado") = Monto
                                                                     DtaConsulta.Recordset.Update
                                                                 ElseIf DtaConsulta.Recordset("DiasAPagar") = 30 Then
                                                                     DtaConsulta.Recordset("SalarioMensual") = Monto
                                                                     DtaConsulta.Recordset("DiasAPagar") = 30
                                                                     DtaConsulta.Recordset("AdelantoVacaciones") = 0
                                                                     DtaConsulta.Recordset("DiasDescuento") = 0
                                                                     DtaConsulta.Recordset("Inss") = Monto * 0.0625
        '                                                             DtaConsulta.Recordset("Ir") = 0
                                                                     DtaConsulta.Recordset("TotalDevengado") = Monto
                                                                     DtaConsulta.Recordset.Update
                                                                 Else
                                                                     
                                                                     DtaConsulta.Recordset("SalarioMensual") = DtaConsulta.Recordset("SalarioMensual") + Monto
                                                                     DtaConsulta.Recordset("Inss") = DtaConsulta.Recordset("Inss") + Monto * 0.0625
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
'On Error GoTo TipoErrs
Dim SqlEmpleados As String, CantEmpleados As Integer
Dim Respuesta As Integer, Cadena As String, FechaVaca As String

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
 
     DtaTipoNomina.Refresh
    Do While Not DtaTipoNomina.Recordset.EOF
    If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
       CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
       Exit Do
    End If
    DtaTipoNomina.Recordset.MoveNext
    Loop
 
  SqlEmpleados = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, " & _
               "Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, " & _
               "Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, " & _
               "Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, " & _
               "Empleado.ExentoIr, Empleado.OtrosIngresos, Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, " & _
               "Empleado.Ausente, Empleado.SalarioFijo, Empleado.CodEmpleado1, Historico.FechaContrato,Historico.FechaContratoVac " & _
               "FROM  Empleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado " & _
               "WHERE  (Empleado.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (Empleado.Ausente = 0)"
'    SqlEmpleados = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, Empleado.ExentoIr, Empleado.OtrosIngresos,  Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, Empleado.Ausente, Empleado.SalarioFijo From Empleado WHERE Empleado.CodTipoNomina= '" & CodTipoNomina & "' AND Empleado.Activo=1 AND Empleado.Ausente=0"
    DtaEmpleados.RecordSource = SqlEmpleados
    DtaEmpleados.Refresh

    DtaEmpleados.Recordset.MoveLast
    
    CantEmpleados = DtaEmpleados.Recordset.RecordCount

    With PBVacaciones
        .Min = 0
        .Max = CantEmpleados
        .Value = 0
         i = 1

       
        NumNomina = val(TxtNumNomVaca.Text)
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
            
'            FechaVaca = Me.TxtFFinVaca.Value + 1
             FechaVaca = Me.TxtFINIVaca.Value
             
             '///////////////////////BUSCO SI EXISTE EN LA TABLA DE REEMBOLSOS /////////////////////////////////////
             Me.AdoBusca.RecordSource = "SELECT  * From Reembolso WHERE (NumNomina = " & NumNomina & ") AND (CodEmpleado = " & CodEmpleado & ")"
             Me.AdoBusca.Refresh
             If Me.AdoBusca.Recordset.EOF Then
                '//////////////////////SI NO EXISTE EN LA TABLA REEMBOLSO LO ACTUALIZO ///////////////////////////////////
                '//////////////////////ACTUALIZO EL EMPLEADO/////////////////////////////////////////////////////////77
                 Set Ejecutar = New ADODB.Connection
                 Ejecutar.ConnectionString = Conexion
                 Ejecutar.Open
                 Ejecutar.Execute "UPDATE Historico SET FechaContratoVac = '" & FechaVaca & "' Where (CodEmpleado = " & CodEmpleado & ")"
             End If
             
             
          If Me.CHKTranferir.Value = 0 Then
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
          End If
          
             DtaConsulta.Recordset.MoveNext
             i = i + 1
        Loop
  
    End With

  
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

 If Me.DBCNominas.Text <> "Administracion" Then
 
      If Me.ChkRestar.Value = 1 Then
           SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, DetalleNomVaca.Ir, Empleado.CodEmpleado,Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir AS TotalDeducir, NomVaca.CodTipoNomina, DetalleNomVaca.Ir FROM  NomVaca INNER JOIN  Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                         "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "
     Else
          SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, DetalleNomVaca.Ir, Empleado.CodEmpleado,Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado,Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir AS TotalDeducir, NomVaca.CodTipoNomina,DetalleNomVaca.Ir FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                        "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "
     End If
     
    '' SQLReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoAPagar,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                   "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "
        
        ArepColillaVaca.AdoColillas.Source = SQlReportes
        ArepColillaVaca.LblTipo.Caption = Me.DBCNominas.Text
        ArepColillaVaca.LblPeriodos.Caption = "Desde   " & Me.TxtFINIVaca.Value & " Hasta    " & Me.TxtFFinVaca.Value
        ArepColillaVaca.LblPeriodo.Caption = Format(Me.TxtFINIVaca.Value, "dd/mm/yyyy") & "   Hasta   " & Format(Me.TxtFFinVaca.Value, "dd/mm/yyyy")
        ArepColillaVaca.lbltitulo.Caption = Titulo
        ArepColillaVaca.AdoColillas.ConnectionString = ConexionReporte
        ArepColillaVaca.Show 1
    '           fPreview.arv.ReportSource = ArepColillaVaca
    '           fPreview.Show 1


 Else

        If Me.ChkRestar.Value = 1 Then
            SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina, DetalleNomVaca.Ir FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                       "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
        
        Else
              SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina, DetalleNomVaca.Ir FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                            "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
        End If
        
        
        
        'SQLReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoAPagar,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                       "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "
         
        ArepColillaVaca.AdoColillas.Source = SQlReportes
        ArepColillaVaca.LblTipo.Caption = Me.DBCNominas.Text
        ArepColillaVaca.LblPeriodos.Caption = "Desde   " & Me.TxtFINIVaca.Value & " Hasta    " & Me.TxtFFinVaca.Value
        ArepColillaVaca.LblPeriodo.Caption = Format(Me.TxtFINIVaca.Value, "dd/mm/yyyy") & "   Hasta   " & Format(Me.TxtFFinVaca.Value, "dd/mm/yyyy")
        ArepColillaVaca.lbltitulo.Caption = Titulo
        ArepColillaVaca.AdoColillas.ConnectionString = ConexionReporte
        ArepColillaVaca.Show 1
        '           fPreview.arv.ReportSource = ArepColillaVaca
        '           fPreview.Show 1

        Exit Sub



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

      SQlReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, (DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo) AS MontoPagar, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, Cargo.Cargo, ([DetalleNom13Mes].SalarioAPagar) AS TotalDevengado, Empleado.CodEmpleado1 FROM Nom13Mes INNER JOIN ((Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNomina & ")) ORDER BY Empleado.CodEmpleado1"

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
       SQlReportes = "SELECT Empleado.CodEmpleado,NomVaca.NumNomVaca , DetalleNomVaca.Inss, DetalleNomVaca.Ir, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                     "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
 Else
      SQlReportes = "SELECT Empleado.CodEmpleado,NomVaca.NumNomVaca , DetalleNomVaca.Inss, DetalleNomVaca.Ir,Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado,Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                    "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
 End If
Else

If Me.ChkRestar.Value = 1 Then
    SQlReportes = "SELECT Empleado.CodEmpleado,NomVaca.NumNomVaca , DetalleNomVaca.Inss, DetalleNomVaca.Ir,Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
               "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "

Else
      SQlReportes = "SELECT Empleado.CodEmpleado,NomVaca.NumNomVaca , DetalleNomVaca.Inss, DetalleNomVaca.Ir,Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                    "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "
End If
'Nom13voMoises.DataControl1.Source = SQLReportes
'Nom13voMoises.ImgLogo.Picture = LoadPicture(RutaLogo)
'Nom13voMoises.Show 1
'Exit Sub

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
Dim Espacio As String
Espacio = " "

NumNomina = Me.TxtNumNom13.Text
SQlReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, (DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo) AS MontoPagar, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, Cargo.Cargo, ([DetalleNom13Mes].SalarioAPagar) AS TotalDevengado, Empleado.CodEmpleado1 FROM Nom13Mes INNER JOIN ((Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNomina & ")) ORDER BY Empleado.CodEmpleado1"
'SQLReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, (DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo) AS MontoPagar, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, Cargo.Cargo, ([DetalleNom13Mes].SalarioAPagar) AS TotalDevengado, Empleado.CodEmpleado1 FROM Nom13Mes INNER JOIN ((Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNomina & ")) ORDER BY Empleado.CodEmpleado1"
ArepColilla13vo.AdoColillas.Source = SQlReportes
ArepColilla13vo.LblTipo.Caption = Me.DBCNominas.Text
ArepColilla13vo.LblPeriodos.Caption = "Desde   " & Me.TxtFINI13.Value & " Hasta    " & Me.TxtFFIN13.Value
ArepColilla13vo.LblPeriodo.Caption = Format(Me.TxtFINI13.Value, "dddddd") & "   Hasta   " & Format(Me.TxtFFIN13.Value, "dddddd")
ArepColilla13vo.lbltitulo.Caption = Titulo
ArepColilla13vo.AdoColillas.ConnectionString = ConexionReporte
ArepColilla13vo.Show 1
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdprNomina_Click()
On Error GoTo TipoErrs
Dim Espacio As String
Espacio = " "
NumNomina = Me.TxtNumNom13.Text


      Nom13vo.lbltitulo.Caption = Titulo
      Nom13vo.LblSubtitulo.Caption = SubTitulo
      Nom13vo.ImgLogo.Picture = LoadPicture(RutaLogo)
      
      Nom13vo.LblFecha.Caption = "Desde " + Format(Me.TxtFINI13.Value, "mm/dd/yyyy") + " Hasta " + Format(Me.TxtFFIN13.Value, "mm/dd/yyyy")
      Nom13vo.LblFechaHoy = Format(Now, "dddddd")
      Nom13vo.DataControl1.ConnectionString = ConexionReporte
      SQlReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, (DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo) AS MontoPagar, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, Cargo.Cargo, ([DetalleNom13Mes].SalarioAPagar) AS TotalDevengado, Empleado.CodEmpleado1 FROM Nom13Mes INNER JOIN ((Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNomina & ")) ORDER BY Empleado.CodEmpleado1"
      Nom13vo.DataControl1.Source = SQlReportes
      Nom13vo.ImgLogo.Picture = LoadPicture(RutaLogo)
      Nom13vo.Show 1

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
Dim FechaIni13 As Date, MetodoVacaciones As String
Dim FechaFin13 As Date
Dim SqlVacaciones As String
Dim Sql13voMes As String
Dim NumNomVaca As Long
Dim NumNom13Mes As Long

 Me.Picture2.BackColor = RGB(173, 199, 236)
 Me.Dbgr13Mes.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.Dbgr13Mes.OddRowStyle.BackColor = &H80000005
 Me.Dbgr13Mes.AlternatingRowStyle = True
 
  Me.DbgrVacaciones.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DbgrVacaciones.OddRowStyle.BackColor = &H80000005
 Me.DbgrVacaciones.AlternatingRowStyle = True
 
 Me.TxtFechaAplica.Value = Format(Now, "dd/mm/yyyy")
 MDIPrimero.DtaEmpresa.Refresh
 
 If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("MetodoVacaciones")) Then
  MetodoVacaciones = MDIPrimero.DtaEmpresa.Recordset("MetodoVacaciones")
 Else
  MsgBox "El Metodo de Vacaciones es Nulo", vbCritical, "Sistema de Nominas"
 End If
 
 If MetodoVacaciones = "Vacaciones Semestrales" Then
   Me.CmdCalVaca.Visible = True
   Me.CmdCalVacaMes.Visible = False
   Me.lbltitulo.Caption = "Vacaciones Semestrales y 13vo Mes"
 ElseIf MetodoVacaciones = "Vacaciones Mensuales" Then
   Me.CmdCalVaca.Visible = False
   Me.CmdCalVacaMes.Visible = True
   Me.lbltitulo.Caption = "Vacaciones Mensuales y 13vo Mes"
 End If
 
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
'    Mes = Month(CDate(DtaConsulta.Recordset("Fechafin")))
'    Año = Year(CDate(DtaConsulta.Recordset("Fechafin")))
'    FechaIni13 = CDate(DtaConsulta.Recordset("Fechaini"))
'    FechaFin13 = CDate(DtaConsulta.Recordset("Fechafin"))
'    TxtNumNom13 = DtaConsulta.Recordset("NumNom13Mes")
     Mes = Month(Frm13VacaMes.TxtFFinVaca.Value)
     Año = Year(Frm13VacaMes.TxtFFinVaca.Value)
    FechaIni13 = Frm13VacaMes.TxtFINIVaca.Value
    FechaFin13 = Frm13VacaMes.TxtFFinVaca.Value
    TxtNumNom13 = Frm13VacaMes.TxtNumNomVaca.Text
     

  End If
 
 
 End If


TxtFINIVaca.Value = Format(FechaIniVaca, "DD/MM/YYYY")
TxtFFinVaca.Value = Format(FechaFinVaca, "DD/MM/YYYY")

TxtFINI13.Value = Format(FechaIni13, "DD/MM/YYYY")
TxtFFIN13.Value = Format(FechaFin13, "DD/MM/YYYY")

DtaConsecutivos.Refresh

NumNomVaca = val(Frm13VacaMes.TxtNumNomVaca.Text)
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

Private Sub SmartButton1_Click()
Dim sql As String, NumeroNomina As Integer

NumeroNomina = Me.TxtNumNom13.Text
sql = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres," & vbLf
sql = sql & "HistorialSalarioMes.Fechaini , HistorialSalarioMes.Fechafin, HistorialSalarioMes.Enero, HistorialSalarioMes.Febrero, HistorialSalarioMes.Marzo, " & vbLf
sql = sql & "HistorialSalarioMes.Abril , HistorialSalarioMes.Mayo, HistorialSalarioMes.Junio, HistorialSalarioMes.Julio, HistorialSalarioMes.Agosto, " & vbLf
sql = sql & "HistorialSalarioMes.Septiembre , HistorialSalarioMes.Octubre, HistorialSalarioMes.Noviembre, HistorialSalarioMes.Diciembre, " & vbLf
sql = sql & "HistorialSalarioMes.NumNomina " & vbLf
sql = sql & "FROM HistorialSalarioMes INNER JOIN" & vbLf
sql = sql & "Empleado ON HistorialSalarioMes.CodEmpleado = Empleado.CodEmpleado" & vbLf
sql = sql & "Where (HistorialSalarioMes.NumNomina = " & NumeroNomina & ")  ORDER BY Empleado.CodEmpleado1"

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
  
  FrmSalarioHistorial.SSTab1.Tab = 1
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

FrmSalarioHistorial.TxtNumNom13.Text = Me.TxtNumNom13.Text
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

 FrmSalarioHistorial.SSTab1.Tab = 0
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
