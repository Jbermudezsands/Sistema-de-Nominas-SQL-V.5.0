VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmCalcularNomina 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calcular Nomina"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9855
   Icon            =   "FrmCalcularNomina.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc DtaAuxiliar 
      Height          =   375
      Left            =   7080
      Top             =   6360
      Width           =   2295
      _ExtentX        =   4048
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
   Begin XtremeSuiteControls.ProgressBar Barra 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   3735
      _Version        =   786432
      _ExtentX        =   6588
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin XtremeSuiteControls.ProgressBar PBCalcNomina 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   7575
      _Version        =   786432
      _ExtentX        =   13361
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin MSAdodcLib.Adodc AdoDetalleProduccionManual 
      Height          =   375
      Left            =   240
      Top             =   6240
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "AdoDetalleProduccionManual"
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
   Begin MSAdodcLib.Adodc AdoDepartamento 
      Height          =   375
      Left            =   3360
      Top             =   6600
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
      Caption         =   "AdoDepartamento"
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
   Begin MSAdodcLib.Adodc AdoDetalleViaticos 
      Height          =   375
      Left            =   240
      Top             =   6840
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
      Caption         =   "AdoDetalleViaticos"
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
   Begin MSAdodcLib.Adodc AdoViaticos 
      Height          =   375
      Left            =   240
      Top             =   6600
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
      Caption         =   "AdoViaticos"
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
   Begin VB.TextBox TxtFechaIni 
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   6600
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc AdoPeriodoFiscal 
      Height          =   375
      Left            =   6600
      Top             =   8760
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
      Caption         =   "AdoPeriodoFiscal"
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
   Begin MSAdodcLib.Adodc AdoSuspendido 
      Height          =   375
      Left            =   6720
      Top             =   6360
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
      Caption         =   "AdoSuspendido"
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
      BackColor       =   &H00E0E0E0&
      Height          =   5850
      Left            =   0
      ScaleHeight     =   5790
      ScaleWidth      =   9795
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin TrueOleDBGrid70.TDBGrid DbgrNominas 
         Bindings        =   "FrmCalcularNomina.frx":5C12
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3625
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
         AllowUpdate     =   0   'False
         Appearance      =   2
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
      Begin VB.Frame Frame1 
         Height          =   5055
         Left            =   7800
         TabIndex        =   1
         Top             =   720
         Width           =   1935
         Begin XtremeSuiteControls.PushButton CmdSalir 
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   4560
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Salir"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmCalcularNomina.frx":5C2E
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton CmdCerrarNomina 
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   4200
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Cerrar"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmCalcularNomina.frx":6132
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton CmdExportaCsv 
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   3840
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Exportar CSV"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmCalcularNomina.frx":84AB
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton CmdExportar 
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   2760
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Exportar"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmCalcularNomina.frx":A883
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton CmdDenominacion 
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   1680
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Monedas"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmCalcularNomina.frx":B007
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton CmdMonedasDpto 
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   2040
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Mond Dep"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmCalcularNomina.frx":E0B1
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton FrmReportes 
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   1320
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Reportes"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmCalcularNomina.frx":103D8
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton CmdPrnNomina 
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   600
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "COLILLAS"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "FrmCalcularNomina.frx":126A8
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton CmdprNomina 
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Top             =   960
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "NOMINAS"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "FrmCalcularNomina.frx":14994
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton CmdCalcular 
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "CALCULAR"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
            Picture         =   "FrmCalcularNomina.frx":16C80
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton CmdExportaBanpro 
            Height          =   375
            Left            =   240
            TabIndex        =   22
            Top             =   3120
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Exportar"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmCalcularNomina.frx":1903B
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Top             =   3480
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   " Exportar"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmCalcularNomina.frx":19917
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   2400
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Exportar"
            ForeColor       =   0
            Appearance      =   6
            Picture         =   "FrmCalcularNomina.frx":19FE1
            ImageAlignment  =   0
         End
      End
      Begin VB.Label LblTotal 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   7215
      End
      Begin VB.Label LblFecha2 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   6720
         TabIndex        =   6
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label LblFecha1 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   6720
         TabIndex        =   5
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando la Nomina"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc AdoIncentivoPro 
      Height          =   450
      Left            =   3240
      Top             =   6600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   794
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
      Caption         =   "AdoIncentivoPro"
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
      Left            =   6360
      Top             =   6720
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
   Begin MSAdodcLib.Adodc AdoHistorico 
      Height          =   375
      Left            =   240
      Top             =   6840
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
      Caption         =   "AdoHistorico"
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
   Begin MSAdodcLib.Adodc DtaHorasProducidas 
      Height          =   375
      Left            =   3360
      Top             =   8760
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
      Caption         =   "DtaHorasProducidas"
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
      Left            =   240
      Top             =   8760
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
   Begin MSAdodcLib.Adodc DtaDeduccion 
      Height          =   375
      Left            =   3360
      Top             =   8400
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Deduccion"
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
   Begin MSAdodcLib.Adodc DtaHrsExtras 
      Height          =   375
      Left            =   5400
      Top             =   8040
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
   Begin MSAdodcLib.Adodc DtaDestajo 
      Height          =   375
      Left            =   6480
      Top             =   7680
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
      Caption         =   "DtaDestajo"
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
   Begin MSAdodcLib.Adodc DtaNewNomina 
      Height          =   375
      Left            =   240
      Top             =   8400
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
      Caption         =   "DtaNewNomina"
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
   Begin MSAdodcLib.Adodc Dtaprestamo 
      Height          =   375
      Left            =   3000
      Top             =   8040
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "Dtaprestamo"
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
      Left            =   240
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
   Begin MSAdodcLib.Adodc DtaIncentivos 
      Height          =   375
      Left            =   3360
      Top             =   7680
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
      Caption         =   "DtaIncentivos"
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
   Begin MSAdodcLib.Adodc DtaDetalleDeduccion 
      Height          =   375
      Left            =   5880
      Top             =   7320
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "DetalleDeduccion"
      Caption         =   "DtaDetalleDeduccion"
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
   Begin MSAdodcLib.Adodc DtaMovprestamo 
      Height          =   375
      Left            =   240
      Top             =   7680
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "MovPrestamo"
      Caption         =   "DtaMovprestamo"
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
   Begin MSAdodcLib.Adodc DtaComisiones 
      Height          =   375
      Left            =   3120
      Top             =   7320
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
      Caption         =   "DtaComisiones"
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
   Begin MSAdodcLib.Adodc DtaDetalleNomina 
      Height          =   375
      Left            =   240
      Top             =   7320
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
      Caption         =   "DtaDetalleNomina"
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
   Begin MSAdodcLib.Adodc DtaNomina 
      Height          =   375
      Left            =   6480
      Top             =   7080
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
      Caption         =   "DtaNomina"
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
      Left            =   6480
      Top             =   7440
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
   Begin MSAdodcLib.Adodc DtaInss 
      Height          =   375
      Left            =   3120
      Top             =   8160
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "INSS"
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
   Begin MSAdodcLib.Adodc DtaPagosMensuales 
      Height          =   375
      Left            =   6000
      Top             =   8160
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
      Caption         =   "DtaPagosMensuales"
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
      Left            =   6000
      Top             =   7800
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
   Begin MSAdodcLib.Adodc DtaIR 
      Height          =   375
      Left            =   240
      Top             =   8160
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "IR"
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
   Begin MSAdodcLib.Adodc DtaNomSubsidios 
      Height          =   375
      Left            =   3240
      Top             =   7800
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
      Caption         =   "DtaNomSubsidios"
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
      Left            =   3240
      Top             =   7440
      Width           =   3255
      _ExtentX        =   5741
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Controles"
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
   Begin MSAdodcLib.Adodc DtaDetalleNominaAnterior 
      Height          =   375
      Left            =   3240
      Top             =   7080
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "DtaDetalleNominaAnterior"
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
   Begin MSAdodcLib.Adodc DtaDetalleIncentivo 
      Height          =   375
      Left            =   240
      Top             =   7800
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "DetalleIncentivo"
      Caption         =   "DtaDetalleIncentivo"
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
      Left            =   240
      Top             =   7440
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Consecutivos"
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
   Begin MSAdodcLib.Adodc DtaNominaMes 
      Height          =   375
      Left            =   240
      Top             =   7080
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
      Caption         =   "DtaNominaMes"
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
   Begin vbskfree.Skinner Skinner1 
      Left            =   360
      Top             =   7440
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
      ChangeControlsBackColor=   0   'False
   End
   Begin MSAdodcLib.Adodc DtaNominas 
      Height          =   375
      Left            =   6240
      Top             =   8400
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Nomina"
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
   Begin MSAdodcLib.Adodc AdoConfiguracion 
      Height          =   375
      Left            =   240
      Top             =   6720
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
      Caption         =   "AdoConfiguracion"
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
   Begin MSAdodcLib.Adodc AdoBusca 
      Height          =   375
      Left            =   3240
      Top             =   6840
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
   Begin MSAdodcLib.Adodc DtaExporta 
      Height          =   375
      Left            =   3480
      Top             =   6240
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
      Caption         =   "DtaExporta"
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
   Begin MSAdodcLib.Adodc AdoIncentivo 
      Height          =   375
      Left            =   8400
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
      Caption         =   "AdoIncentivo"
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
   Begin MSAdodcLib.Adodc AdoDetalleIncentivo 
      Height          =   375
      Left            =   7680
      Top             =   8280
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
      Caption         =   "AdoDetalleIncentivo"
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
Attribute VB_Name = "FrmCalcularNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NumeroNominas As Double
Public Pmes As String
Public PAno As Integer
Public PFechaNomina As Date
Public PTipoNomina As String
Public Periodo As Integer
Public MontoExento As Double

Public Function AgregarDiasAdicionales(Dias As Double, MontoHora As Double, DiaMes As Double, CodigoEmpleado As String, CodTipoNomina As String, NumNomina As String)
  Dim DiasAdicionales As Double, NumeroIncentivo As Double, MontoIncentivo As Double
  Dim MontoDeduccion As Double, IdIncentivo As Double
  Dim Horas As Double
        
        MontoIncentivo = 0
        
        MDIPrimero.DtaConsulta.RecordSource = "SELECT  * From TipoNomina WHERE (CodTipoNomina = '" & CodTipoNomina & "')"
        MDIPrimero.DtaConsulta.Refresh
        If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
          Horas = MDIPrimero.DtaConsulta.Recordset("Horas")
        End If


        '////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////// Deduccion por dias de descuento /////////////////////////////////
        'calculo la deduccion por los dias de descuento
       
            If Not IsNull(DtaEmpleados.Recordset("CantPts")) And DtaEmpleados.Recordset("CantPts") > 0 Then
                    MontoIncentivo = (MontoHora * Horas) * Dias
                    
                    '/////////////////////////////////////////////////////////////////////
                    '/////GRABO LA DEDUCCION POR FALTAS/////////////////////////////////
                    '///////////////////////////////////////////////////////////////////
                    
'                    Me.DtaConsulta.RecordSource = "SELECT Deduccion.NumDeduccion, Deduccion.CodEmpleado, Deduccion.CodTipoDeduccion, DetalleDeduccion.NumNomina " & _
'                                                  "FROM Deduccion INNER JOIN " & _
'                                                  "DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion " & _
'                                                  "WHERE (Deduccion.CodTipoDeduccion = '01') AND (DetalleDeduccion.NumNomina = " & CDbl(NumNomina) & ") AND (Deduccion.CodEmpleado = " & CDbl(CodEmpleado) & ") "
                    Me.DtaConsulta.RecordSource = "SELECT DetalleIncentivo.*, Incentivo.CodEmpleado, DetalleIncentivo.NumNomina, Incentivo.CodTipoIncentivo FROM Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo  " & _
                                                  "WHERE (Incentivo.CodEmpleado = " & CDbl(CodigoEmpleado) & ") AND (DetalleIncentivo.NumNomina = " & CDbl(NumNomina) & ") AND (Incentivo.CodTipoIncentivo = '01')"
                    Me.DtaConsulta.Refresh
                    
                    If Me.DtaConsulta.Recordset.EOF Then
                    
                         'creo la nueva deducción
                        DtaConsecutivos.Refresh
                        DtaConsecutivos.Recordset("Incentivos") = DtaConsecutivos.Recordset("Incentivos") + 1
                        DtaConsecutivos.Recordset.Update
                        
                        Me.adoIncentivo.RecordSource = "SELECT * From Incentivo ORDER BY NumIncentivo"
                        Me.adoIncentivo.Refresh
                        If Me.adoIncentivo.Recordset.EOF Then
                            NumeroIncentivo = 1
                        Else
                            Me.adoIncentivo.Recordset.MoveLast
                            NumeroIncentivo = Me.adoIncentivo.Recordset("NumIncentivo") + 1
                        End If
                    
                        Me.adoIncentivo.Recordset.AddNew
                        adoIncentivo.Recordset("NumIncentivo") = NumeroIncentivo
                        adoIncentivo.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
                        adoIncentivo.Recordset("CodTipoIncentivo") = "01"
                        adoIncentivo.Recordset("NumVeces") = "1"
                        adoIncentivo.Recordset("Pagado") = 0
                        adoIncentivo.Recordset.Update
                    
                        '////////////////////////////////////////////////////////////////////
                        '//GRABO EL DETALLE DE LA DEDUCCION POR FALTAS//////////////////////
                        '///////////////////////////////////////////////////////////////////
                        
                        Me.DtaConsulta.RecordSource = "SELECT id,Deduccion.CodTipoDeduccion, DetalleDeduccion.NumDeduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.Pagado, DetalleDeduccion.NumNomina, Deduccion.CodEmpleado FROM Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion "
                        Me.DtaConsulta.Refresh
                        
                        If Me.DtaConsulta.Recordset.EOF Then
                          IdIncentivo = 1
                        Else
                          Me.DtaConsulta.Recordset.MoveLast
                          IdIncentivo = Me.DtaConsulta.Recordset("Id") + 1
                        End If
                        
                        Me.AdoDetalleIncentivo.RecordSource = "SELECT  * From DetalleIncentivo"
                        Me.AdoDetalleIncentivo.Refresh
                        AdoDetalleIncentivo.Recordset.AddNew
                        AdoDetalleIncentivo.Recordset("Id") = IdIncentivo
                        AdoDetalleIncentivo.Recordset("NumIncentivo") = NumeroIncentivo
                        AdoDetalleIncentivo.Recordset("valor") = Format(MontoIncentivo, "##,##0.00")
                        AdoDetalleIncentivo.Recordset("NumVez") = 1
                        AdoDetalleIncentivo.Recordset("pagado") = 0
                        AdoDetalleIncentivo.Recordset("NumNomina") = NumNomina
                        AdoDetalleIncentivo.Recordset.Update
                        
                    Else
'                        Me.DtaConsulta.RecordSource = "SELECT Deduccion.CodEmpleado, Deduccion.CodTipoDeduccion, DetalleDeduccion.NumNomina, DetalleDeduccion.Valor " & _
'                                                      "FROM  DetalleDeduccion INNER JOIN Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion " & _
'                                                      "WHERE  (Deduccion.CodEmpleado = " & CDbl(CodEmpleado) & ") AND (Deduccion.CodTipoDeduccion = '01') AND (DetalleDeduccion.NumNomina = " & CDbl(NumNomina) & ") "
                        Me.DtaConsulta.RecordSource = "SELECT  DetalleIncentivo.NumNomina, DetalleIncentivo.NumIncentivo AS Expr1, DetalleIncentivo.Valor FROM DetalleIncentivo INNER JOIN Incentivo ON DetalleIncentivo.NumIncentivo = Incentivo.NumIncentivo " & _
                                                      "Where (DetalleIncentivo.NumNomina = " & CDbl(NumNomina) & ") And (Incentivo.CodTipoIncentivo = '01') And (Incentivo.CodEmpleado = " & CDbl(CodigoEmpleado) & ")"
                        Me.DtaConsulta.Refresh
                        If Not Me.DtaConsulta.Recordset.EOF Then
                           '////////////////////////////////////////////////////////////////////////////////
                           '///////////EDITO LA DEDUCCION SI EXISTE////////////////////////////////////////
                           '///////////////////////////////////////////////////////////////////////////////
                            DtaConsulta.Recordset("valor") = Format(MontoIncentivo, "##,##0.00")
                            DtaConsulta.Recordset.Update
                        End If
                    End If

            Else  '////////////////////////SI NO ES MAYOR DE CERO... LO HAGO CERO ///////////////
                        Me.DtaConsulta.RecordSource = "SELECT  DetalleIncentivo.NumNomina, DetalleIncentivo.NumIncentivo AS Expr1, DetalleIncentivo.Valor FROM DetalleIncentivo INNER JOIN Incentivo ON DetalleIncentivo.NumIncentivo = Incentivo.NumIncentivo " & _
                                                      "Where (DetalleIncentivo.NumNomina = " & CDbl(NumNomina) & ") And (Incentivo.CodTipoIncentivo = '01') And (Incentivo.CodEmpleado = " & CDbl(CodigoEmpleado) & ")"
                        Me.DtaConsulta.Refresh
                        If Not Me.DtaConsulta.Recordset.EOF Then
                           '////////////////////////////////////////////////////////////////////////////////
                           '///////////EDITO LA DEDUCCION SI EXISTE////////////////////////////////////////
                           '///////////////////////////////////////////////////////////////////////////////
                           DtaConsulta.Recordset("valor") = Format(0, "##,##0.00")
                           DtaConsulta.Recordset.Update
                        End If
            
            
            End If


End Function
Public Function CalcularMontoIr(CodTipoNomina As String, TotalDevengado As Double, CodEmpleado As String, TipoCalculoIr As String, IrUltimaSemana As Boolean, MontoInss As Double, NumNomina As String) As Double
Dim i As Integer, sql As String, TarifaHorariaBasico As Double, FechaIni As String, TotalSalarioxHora As Double, TasaCambio As Double, SQLNominaEmpleado As String, TarifaHoraria As Double, SQLNomina As String, TotalHoras As Double, CodDepartamento As String, MontoIncentivoHoras As Double, PorcientoIncentivo As Double, TotalPuntualidad As Double, SQlIncentivos As String, Septimo As Double, SQlDeducciones As String, MontoIrAcumulado As Double, SQlPrestamo As String, SQlComisiones As String, SQlDestajo As String, SqlHrsExtras As String, SueldoPeriodo As Double, TasaInss As Double, TasaInssPatronal As Double, MontoIncentivos As Double, TasaIr As Double, MontoDeduccion As Double, MontoPrestamo As Double, MontoComisiones As Double, FechaContrato As Date, MontoDestajos As Double, annos As Date, MontoHRSExtras As Double, Antiguedad As Double
Dim MontoOtrosIngresos As Double, PorcientoAntiguedad As Double, DescripOtrIngre As String, NumFecha1 As Date, MontoHora As Double, CodProceso As String, CantEmpleados As Long, CodReferencia As String, MontoIr As Double, UnidadesProducidas As Double, Rango As Double, Monto As Double, MontoIRPatronal As Double, NumeroDeduccion As Double, MontoInssPatronal As Double, MontoBrutoAnual As Double, MontoVacaciones As Double, Nombres As String, MontoMes13 As Double, FechaNomina As Date, DeduccionPorFalta As Double, SeptimoAnterior As Double, MinIR As Double, AñoFiscal As Double, SalarioMensual As Double, RentaGravable As Double, DiasMes As Double, TotalDevengadoAcumulado As Double, DiasSemana As Double, IncentivoProduccion As Double, CantSabados As Byte, IdDeduccion As Double, PagoProduccion As Double, SalHora As Double, NumeroPeriodo As Double, PeriodoFiscal As Double, Factor As Double, NQuincenas As Double, INATEC As Double, FechaInicialIr As Date
Dim FechaFinalIr As Date, DevengadoSinHrsExtras As Double, VacacionesAcumuladas As Double, HE As Single, HoraPuntualidad As Double, MontoPuntualidad As Double, DD As Single, HoraSeptimo As Double, HoraBasico As Double, FormatoNomina As String, Adelantos As Double, Anos As Double, Moneda As String, MontoDolares As Double, MontoProduccion As Double, agregar As Boolean, FechaIngreso As Date, PeriodoIngreso As Double, BonoProduccion As Double, MontoViaticos As Double, NumIncentivo As Double, FechaInicio As String, FechaFin As String, Mes As Double, Fecha As String, Calcular7mo As Boolean, Dolarizado As Boolean, cn As New ADODB.Connection, ValorPunto As Double, SalarioMinimo As Double, SalarioPorciento As Double, rs As New ADODB.Recordset, TotalPuntos As Double, SalarioBasico As Double, CalcularPuntos As Boolean, MontoInssBasico As Double, AjusteINSS As Double, cmd As New ADODB.Command, HT As Double, MontoHorasTurno As Double
Dim TipoVacaciones As Boolean, MontoTipoVacaciones As Double, CalcularHorasTurno As Boolean
Dim AnoIni As Double, Viaticos As Double, NumeroNominaAnt As Double

AnoIni = Year(Me.DtaNomina.Recordset("FechaNomina"))
Mes = (Me.DtaNomina.Recordset("Mes"))

        '-------------------------------------------------------------------------------------------------------
        '------------------------------BUSCO EL SALARIO BASICO DEL EMPLEADO-------------------------------------
        '-------------------------------------------------------------------------------------------------------
             
        
        
        Me.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, PorcientoInss, TasaInss, PorcientoIr, TasaIr From TipoNomina WHERE (PorcientoIr = 1) AND (CodTipoNomina = '" & CodTipoNomina & "' )"
        Me.DtaConsulta.Refresh
        If DtaConsulta.Recordset.EOF Then
        
        End If

        '//////////////////////////////////////////////////
        '///PRIMERO BUSCO EL NUMERO DEL PERIODO PARA CALCULAR IR
        '////////////////////////////////////////////////////////
        '///////////////////////Verifico si Tiene Ir Porcentual//////////////////////////////
'        CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
        Me.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, PorcientoInss, TasaInss, PorcientoIr, TasaIr From TipoNomina WHERE (PorcientoIr = 1) AND (CodTipoNomina = '" & CodTipoNomina & "' )"
        Me.DtaConsulta.Refresh
        If DtaConsulta.Recordset.EOF Then
        
         Select Case DtaTipoNomina.Recordset("Periodo")
              Case "Catorcenal los Sabados"
                Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (CodTipoNomina = " & CodTipoNomina & ") AND (Inicio = '" & Format(CDate(Me.TxtFechaIni.Text), "DD/MM/YYYY") & "') ORDER BY Periodo"
                Me.AdoPeriodoFiscal.Refresh
                If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
                   PeriodoFiscal = Me.AdoPeriodoFiscal.Recordset("Periodo") ' formula = n
                   NumeroPeriodo = 24 - (PeriodoFiscal - 1) 'formula = 24-(n-1)
                   AñoFiscal = Me.AdoPeriodoFiscal.Recordset("Año")
                End If
              Case "Quincenal"
                Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (CodTipoNomina = " & CodTipoNomina & ") AND (Inicio = '" & Format(CDate(Me.TxtFechaIni.Text), "DD/MM/YYYY") & "') ORDER BY Periodo"
                Me.AdoPeriodoFiscal.Refresh
                If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
                   PeriodoFiscal = Me.AdoPeriodoFiscal.Recordset("Periodo") ' formula = n
                   NumeroPeriodo = 24 - (PeriodoFiscal - 1) 'formula = 24-(n-1)
                   AñoFiscal = Me.AdoPeriodoFiscal.Recordset("Año")
                End If
        
               Case "Mensual"
                Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (CodTipoNomina = " & CodTipoNomina & ") AND (Inicio = '" & Format(CDate(Me.TxtFechaIni.Text), "DD/MM/YYYY") & "') ORDER BY Periodo"
                Me.AdoPeriodoFiscal.Refresh
                If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
                   PeriodoFiscal = Me.AdoPeriodoFiscal.Recordset("Periodo") ' formula = n
                   NumeroPeriodo = 12 - (PeriodoFiscal - 1) 'formula = 12-(n-1)
                   AñoFiscal = Me.AdoPeriodoFiscal.Recordset("Año")
                End If
        
        
          End Select
        End If
        
        '//////////////////////////////////////////////////
        '///BUSCO LA FECHA INICIAL DEL AÑO FISCAL
        '////////////////////////////////////////////////////////
                Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (Año = " & AñoFiscal & ") AND (CodTipoNomina = " & CodTipoNomina & ") AND (Periodo = 1)ORDER BY Periodo"
                Me.AdoPeriodoFiscal.Refresh
                If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
                  FechaInicialIr = Me.AdoPeriodoFiscal.Recordset("Inicio")
                End If
        
        '//////////////////////////////////////////////////
        '///BUSCO LA FECHA DE LA ULTIMA NOMINA DEL AÑO FISCAL CALCULADA
        '////////////////////////////////////////////////////////
                PeriodoFiscal = PeriodoFiscal - 1
                Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (Año = " & AñoFiscal & ") AND (CodTipoNomina = " & CodTipoNomina & ") AND (Periodo = " & PeriodoFiscal & ") ORDER BY Periodo"
                Me.AdoPeriodoFiscal.Refresh
                PeriodoFiscal = PeriodoFiscal + 1
                If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
                  FechaFinalIr = Me.AdoPeriodoFiscal.Recordset("Final")
                End If
        
        '/////////////////////////////////////////////////////////////////
        '///BUSCO LAS NOMINAS ACUMULADAS/////////////////////////////////
        '///////////////////////////////////////////////////////////////////
        
        sql = "SELECT     DetalleNomina.CodEmpleado AS CodEmpleado, SUM(DetalleNomina.MontoINSS) AS MontoINSS, SUM(DetalleNomina.MontoIR) AS MontoIR, " & _
             "SUM(DetalleNomina.VacacionesPagadas) AS Vacaciones, SUM(DetalleNomina.INSSPatronal) AS INSSPatronal, SUM(DetalleNomina.IRPatronal) AS IRPatronal, " & _
             "SUM(DetalleNomina.INATEC) AS INATEC, " & _
             "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos " & _
             " + DetalleNomina.VacacionesPagadas + DetalleNomina.AdelantosVacaciones) AS TotalDevengado, COUNT(DetalleNomina.NumNomina) AS NQuincenas, MIN(Nomina.FechaNominaINI) AS FechaIngreso FROM DetalleNomina INNER JOIN " & _
             "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina " & _
             "WHERE     (Nomina.FechaNomina <= CONVERT(DATETIME, '" & Format(FechaFinalIr, "yyyy/mm/dd") & "', 102)) AND (Nomina.FechaNominaINI >= CONVERT(DATETIME, '" & Format(FechaInicialIr, "yyyy/mm/dd") & "', 102)) " & _
             "GROUP BY DetalleNomina.CodEmpleado " & _
             "Having (DetalleNomina.CodEmpleado = " & CodEmpleado & ") "
        
        '+ DetalleNomina.Incentivos   PANAM LO QUITE
        
            Me.DtaConsulta.RecordSource = sql
            Me.DtaConsulta.Refresh
            TotalDevengadoAcumulado = 0
            MontoIrAcumulado = 0
            VacacionesPagadas = 0
            If Not Me.DtaConsulta.Recordset.EOF Then
               MontoIrAcumulado = Me.DtaConsulta.Recordset("MontoIR")
               TotalDevengadoAcumulado = Me.DtaConsulta.Recordset("TotalDevengado") - Me.DtaConsulta.Recordset("MontoINSS")
               NQuincenas = Me.DtaConsulta.Recordset("NQuincenas") + 1
               FechaIngreso = Me.DtaConsulta.Recordset("FechaIngreso")
               VacacionesAcumuladas = Me.DtaConsulta.Recordset("Vacaciones")
            Else
               NQuincenas = 1
               FechaIngreso = Me.TxtFechaIni.Text
        
            End If
        
        
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////BUSCO SI EXISTE NOMINA ACUMULADA //////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
         sql = "SELECT id, NumNomina, CodEmpleado, SalarioBasico, Destajo, HE, DD, HorasExtras, Comisiones, OtrosIngresos, DescripOtrIngre, Incentivos, Deducciones, Prestamo, MontoINSS, MontoIR, Vacaciones, INSSPatronal, IRPatronal, INATEC, Mes13, DiasDescuento, Adelantos, TotalSubsidio, VacacionesPagadas, DiasVacaciones, AdelantosVacaciones, HTrabajada, SeptimoDia, IncetivoProduccion, TarifaHoraria, produjo, BonoProduccion, Viaticos, Ajuste, TIngresos, TGastos, SalarioBasico + Destajo + HorasExtras + Comisiones + OtrosIngresos + Incentivos + SeptimoDia + IncetivoProduccion + BonoProduccion AS TotalDevengado,NQuincenaAcumulada From DetalleNominaAcumulada Where (NumNomina = 0) And (CodEmpleado = " & CodEmpleado & ")"
         Me.DtaConsulta.RecordSource = sql
         Me.DtaConsulta.Refresh
             If Not Me.DtaConsulta.Recordset.EOF Then
               MontoIrAcumulado = Me.DtaConsulta.Recordset("MontoIR") + MontoIrAcumulado
               TotalDevengadoAcumulado = TotalDevengadoAcumulado + Me.DtaConsulta.Recordset("TotalDevengado") - Me.DtaConsulta.Recordset("MontoINSS")
               VacacionesAcumuladas = Me.DtaConsulta.Recordset("Vacaciones")
               If Not IsNull(Me.DtaConsulta.Recordset("NQuincenaAcumulada")) Then
                 NQuincenas = Me.DtaConsulta.Recordset("NQuincenaAcumulada") + NQuincenas
               End If
             End If
        
        
        
        
        '//////////////////////////////////////////////////
        '///BUSCO EL PERIODO DE INGRESO DEL EMPLEADO
        '////////////////////////////////////////////////////////
                Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (CodTipoNomina = " & CodTipoNomina & ") AND (Inicio = '" & Format(CDate(FechaIngreso), "DD/MM/YYYY") & "') ORDER BY Periodo"
                Me.AdoPeriodoFiscal.Refresh
                If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
                   PeriodoIngreso = Me.AdoPeriodoFiscal.Recordset("Periodo")
                End If
        
        
        
        '///////////////////////Verifico si Tiene Ir Porcentual//////////////////////////////
'        CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
        Me.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, PorcientoInss, TasaInss, PorcientoIr, TasaIr From TipoNomina WHERE (PorcientoIr = 1) AND (CodTipoNomina = '" & CodTipoNomina & "' )"
        Me.DtaConsulta.Refresh
        If DtaConsulta.Recordset.EOF Then
         'Hago el Calcul del nuevo Techo para el Ir
         Select Case DtaTipoNomina.Recordset("Periodo")
                        Case "Semanal Viernes"
        
                            If BuscaUltimaSemana(CDbl(CantSabados), CDbl(NumNomina), Format(Mes, "0#"), CDbl(AnoIni)) = True Then
                             MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
                             MontoBrutoMensual = MontoBruto + TotalSueldoAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes)) - TotalInssAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes))
                            ElseIf IrUltimaSemana = False Then
                                MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
                                MontoBrutoMensual = MontoBruto * CantSabados
                            Else
                                MontoBrutoMensual = 0
                            End If
        
                        Case "Semanal Sabado"
        
                            If BuscaUltimaSemana(CDbl(CantSabados), CDbl(NumNomina), Format(Mes, "0#"), CDbl(AnoIni)) = True Then
                             MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
                             MontoBrutoMensual = MontoBruto + TotalSueldoAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes)) - TotalInssAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes))
                            ElseIf IrUltimaSemana = False Then
                                MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
                                MontoBrutoMensual = MontoBruto * CantSabados
                            Else
                                MontoBrutoMensual = 0
                            End If
        
                        Case "Catorcenal los Viernes"
                            If DiaFin < 28 Then
                             MontoBruto = (TotalDevengado + MontoOtrosIngresos + MontoTipoVacaciones) - MontoInss
                             MontoBrutoMensual = ((MontoBruto * 15) / 14) * 2
                            Else
                             MontoBrutoMensual = SalarioMensual - MontoInssMensual
                            End If
                        Case "Catorcenal los Sabados"
                        'EMPIEZO A BUSCAR SI EN EL PERIODO EN EL QUE ESTOY ES LA ULTIMA SEMANA, SI LO ES ENTONCES CALCULO
                        'SI NO EXISTEN FILAS/ROWS ENTONCES SE CALCULA IR
                        'SELECT     Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina   FROM         Fecha_Planilla  WHERE     (año = 2017) AND (mes = N'04') AND (CodTipoNomina = N'04') AND (Inicio > CONVERT(DATETIME, '04-17-2017 00:00:00', 102))
                        
                        
                        
                        'CodTipoNomina
                        'DtaNomina.Recordset ("FechaNomina")
                        'Mes = (DtaNomina.Recordset("Mes"))
                        
                        
                        If IrUltimaSemana = False Then
                              MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
                              MontoBrutoMensual = (MontoBruto * 26) / 12
                        Else
                          
                             Me.DtaConsulta.RecordSource = "SELECT     Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina   FROM         Fecha_Planilla  WHERE     (año = " & PAno & ") AND (mes ='" & Format(Pmes, "00") & "') AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Inicio > CONVERT(DATETIME, '" & Format(PFechaNomina, "MM-dd-yyyy") & " 00:00:00', 102))"
                             Me.DtaConsulta.Refresh
                             If Me.DtaConsulta.Recordset.EOF Then
                             
                             Dim pPeriodo As Integer
                             pPeriodo = Periodo
                             
                             'Periodo Actual
                             'MontoComisiones  no sumo comisiones, guardo viaticos
                                 MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
                                 MontoBrutoMensual = (MontoBruto * 26) / 12
                                 
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
                                End If
                                '/////////////////////////////////////////////////////////////////////////////////////
                                '////////////////////////BUSCO LOS INCENTIVOS EXCENTEOS PARA RESTAR EL DEVENGADO ///
                                '///////////////////////////////////////////////////////////////////////////////////
                                  '/////////////////////////////BUSCO LOS INCENTIVOS /////////////////////////////////////////////
'                                Me.DtaConsulta.RecordSource = "SELECT  Nomina.* From Nomina WHERE (Mes = " & Pmes & ") AND (Ano = " & PAno & ") AND (Periodo = " & pPeriodo & ") AND (CodTipoNomina = '" & CodTipoNomina & "')"
'                                Me.DtaConsulta.Refresh
'                                If Not Me.DtaConsulta.Recordset.EOF Then
'                                  NumeroNominaAnt = Me.DtaConsulta.Recordset("NumNomina")
'                                End If
                                

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
                                  MontoBrutoMensual = MontoBruto + (Tdevenga - tInss - Viaticos)
                                 
                             Else
                                 MontoBrutoMensual = 0
                                 
                                 
                             End If
 
                             
                        End If
                        
                        
                        
                        
   
                        
                              
                        Case "Quincenal"
                          If TipoCalculoIr = "Calcular IR x 12" Then
                            If DiaFin < 28 Then
                              If IrUltimaSemana = False Then
                                MontoBruto = (TotalDevengado) - MontoInss
                                MontoBrutoMensual = MontoBruto * 2
                                MontoBrutoAnual = MontoBrutoMensual * 12
                              Else
                                MontoBruto = 0
                                MontoBrutoMensual = 0
                                MontoBrutoAnual = 0
                              End If
                            ElseIf IrUltimaSemana = False Then
                                MontoBruto = (TotalDevengado) - MontoInss
                                MontoBrutoMensual = MontoBruto * 2
                                MontoBrutoAnual = MontoBrutoMensual * 12
                                '                        If TotalDevengadoAnterior = 0 Then
        '                           MontoBrutoMensual = (SalarioMensual - MontoInssMensual) * 2
        '                           MontoBrutoAnual = MontoBrutoMensual * 12
        '                        Else
        '                           MontoBrutoMensual = SalarioMensual - MontoInssMensual
        '                           MontoBrutoAnual = MontoBrutoMensual * 12
        '                        End If
                            ElseIf IrUltimaSemana = True Then
                             MontoBruto = (TotalDevengado) - MontoInss
                             MontoBrutoMensual = MontoBruto + TotalSueldoAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes)) - TotalInssAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes))
                             MontoBrutoAnual = MontoBrutoMensual * 12
                            End If
                          Else
                           MontoBruto = (TotalDevengado) - MontoInss '+ MontoOtrosIngresos
                           RentaGravable = ((TotalDevengadoAcumulado + MontoBruto) / NQuincenas) * 24
        
                           MontoBrutoAnual = RentaGravable '+ MontoVacaciones + VacacionesAcumuladas
                           MontoBrutoMensual = MontoBruto * 2
                          End If
        
                        Case "Mensual"
        
                           MontoBruto = (TotalDevengado) - MontoInss
                           RentaGravable = ((TotalDevengadoAcumulado + MontoBruto) * (12 - (PeriodoIngreso - 1))) / NQuincenas
        '                   MontoBrutoAnual = RentaGravable + MontoVacaciones + VacacionesAcumuladas
                           MontoBrutoMensual = MontoBruto
                           MontoBrutoAnual = MontoBrutoMensual * 12
        '                    MontoBruto = SalarioMensual - MontoInssMensual
        '                    MontoBrutoMensual = MontoBruto
                        Case "Trimestral"
        
                            MontoBruto = SalarioMensual - MontoInssMensual
                            MontoBrutoMensual = MontoBruto / 3
                        Case "Semestral"
        
                            MontoBruto = SalarioMensual - MontoInssMensual
                            MontoBrutoMensual = MontoBruto / 6
        End Select
        
        
          '//////////////////////////////////////////////////////////////////////////
          '///////////////////BUSCO EL TIPO DE MONEDA DE LA NOMINA///////////////////
          '//////////////////////////////////////////////////////////////////////////
           Me.AdoBusca.RecordSource = "SELECT CodTipoNomina, Nomina, Periodo, UltFecha, TipoPago, Moneda, MantValor, Activa, PorcientoInss, TasaInss, PorcientoIr, TasaIr,TasaInssPatronal From TipoNomina WHERE (CodTipoNomina = '" & CodTipoNomina & "')"
           Me.AdoBusca.Refresh
           If Not Me.AdoBusca.Recordset.EOF Then
              Moneda = Me.AdoBusca.Recordset("Moneda")
           Else
              Moneda = "C$"
           End If
        
        If DtaEmpleados.Recordset("ExentoIr") = False Then
                'agregar IR laboral y patronal
        
                MontoIr = 0
                MontoIRPatronal = 0
                MontoDolares = 0
                If Moneda = "US" Then
                 MontoDolares = MontoBrutoMensual
                 MontoBrutoMensual = MontoBrutoMensual * TasaCambio
                End If
        
        
                DtaIr.Refresh
                DtaIr.Recordset.MoveNext
                MinIR = DtaIr.Recordset("desde")
                MinIR = MinIR - 1
                MinIR = (MinIR / 12)
                Do While Not DtaIr.Recordset.EOF
        
                   'ubicar la linea
                 If DtaTipoNomina.Recordset("Periodo") = "Semanal Viernes" Then
                    If (MontoBrutoMensual) >= MinIR Then
                    If DtaIr.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIr.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
                       MontoIr = ((MontoBrutoMensual * 12) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
                       MontoIr = Format(MontoIr / 12, "###,##0.00")  'MontoIr = Format(MontoIr / CantSabados / 12, "###,##0.00")
                       MontoIRPatronal = MontoIr
                       Exit Do
                    End If
                    End If
        
                 ElseIf DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
                    If (MontoBrutoMensual) >= MinIR Then
                    If DtaIr.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIr.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
                       MontoIr = ((MontoBrutoMensual * 12) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
                       MontoIr = Format(MontoIr / 12, "###,##0.00")
                       MontoIRPatronal = MontoIr
                       Exit Do
        
                    End If
                    End If
        
                ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Viernes" Then
                    If (MontoBrutoMensual) >= MinIR Then
                    If DtaIr.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIr.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
                       MontoIr = ((MontoBrutoMensual * 12) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
          '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
                       If DiaFin < 28 Then
                        MontoIr = Format(MontoIr / 2 / 12, "###,##0.00")
                        MontoIRPatronal = MontoIr
                        Exit Do
                       Else
                        MontoIrMensual = Format(MontoIr / 1 / 12, "###,##0.00")
                        MontoIr = MontoIrMensual - MontoIrAnterior
                        MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
                       End If
                    End If
                    Else
                       MontoIrMensual = 0
                       MontoIr = MontoIrMensual - MontoIrAnterior
                       MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
                    End If
                 ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
                    If (MontoBrutoMensual) >= MinIR Then
                    If DtaIr.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIr.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
                       MontoIr = ((MontoBrutoMensual * 12) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
          '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
                       If DiaFin < 20 Then
                            If IrUltimaSemana = False Then
                                MontoIr = Format(MontoIr / 26, "###,##0.00")
                                MontoIRPatronal = MontoIr
                            Else
                             MontoIr = Format(MontoIr / 12, "###,##0.00")
                             MontoIRPatronal = MontoIr
                            End If
                       Else
                            If IrUltimaSemana = False Then
                                MontoIr = Format(MontoIr / 26, "###,##0.00")
                                MontoIRPatronal = MontoIr
                            Else
                             MontoIr = Format(MontoIr / 12, "###,##0.00")
                             MontoIRPatronal = MontoIr
                            End If
                       End If
                    End If
                    Else
                       MontoIrMensual = 0
                        MontoIr = MontoIrMensual - MontoIrAnterior
                        MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
                    End If
        
        
                 ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
                     If DtaIr.Recordset("desde") <= (MontoBrutoAnual) And DtaIr.Recordset("Hasta") >= (MontoBrutoAnual) Then
                       MontoIr = ((MontoBrutoAnual) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
        '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
        
                        If TipoCalculoIr = "Calcular IR x 12" Then
                            If DiaFin < 28 Then
                                MontoIr = Format(MontoIr / 2 / 12, "###,##0.00")
                                MontoIRPatronal = MontoIr
                                Exit Do
                            ElseIf IrUltimaSemana = False Then
'                                MontoIrAcumulado = TotalIrAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes))
'                                MontoIr = Format((MontoIr / 12) - MontoIrAcumulado, "###,##0.00")
'                                MontoIRPatronal = MontoIr
                                MontoIr = Format(MontoIr / 24, "###,##0.00")
                                MontoIRPatronal = MontoIr
                            ElseIf IrUltimaSemana = True Then
                                MontoIrAcumulado = TotalIrAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes))
                                MontoIr = Format((MontoIr / 12) - MontoIrAcumulado, "###,##0.00")
                                MontoIRPatronal = MontoIr
                            End If
                        Else
                        If Not NumeroPeriodo = 0 Then
                          'NumeroPeriodo = 24-(NQuincenas-1)
                         MontoIr = (MontoIr - MontoIrAcumulado) / NumeroPeriodo
                         ' MontoIr = ((MontoIr / 24) * NQuincenas) - MontoIrAcumulado
                        Else
                         MontoIr = 0
                        End If
                        End If
        
                        MontoIRPatronal = MontoIr - MontoIrPatronalAnterior
                        Exit Do
        '               End If
                     End If
        '            Else
        '               MontoIrMensual = 0
        
        '                MontoIR = MontoIrMensual - MontoIrAnterior
        '                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
        '            End If
        
        
        
                 ElseIf DtaTipoNomina.Recordset("Periodo") = "Mensual" Then
        '           If (MontoBrutoAnual) >= MinIR Then
                    If DtaIr.Recordset("desde") <= (MontoBrutoAnual) And DtaIr.Recordset("Hasta") >= (MontoBrutoAnual) Then
        
                       MontoIr = ((MontoBrutoAnual) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
        
                        MontoIr = (MontoIr - MontoIrAcumulado) / 12
                        MontoIRPatronal = MontoIr - MontoIrPatronalAnterior
                        Exit Do
        
                       Exit Do
                    End If
        '         End If
                 ElseIf DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
                   If (MontoBrutoMensual) >= MinIR Then
                    If DtaIr.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIr.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
                       MontoIr = ((MontoBrutoMensual * 12) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
                       MontoIr = Format(MontoIr / 4, "###,##0.00")
                       MontoIRPatronal = MontoIr
                       Exit Do
                    End If
                   End If
                 ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
                     If (MontoBrutoMensual) >= MinIR Then
                    If DtaIr.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIr.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
                       MontoIr = ((MontoBrutoMensual * 12) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
                       MontoIr = Format(MontoIr / 2, "###,##0.00")
                       MontoIRPatronal = MontoIr
                       Exit Do
                    End If
                    End If
                 End If
          DtaIr.Recordset.MoveNext
          Loop
        
            If Moneda = "US" Then
               MontoBrutoMensual = MontoDolares
               If TasaCambio <> 0 Then
                MontoIr = MontoIr / TasaCambio
                MontoIRPatronal = MontoIRPatronal / TasaCambio
               End If
            End If
        
          End If 'del if que pregunta si esta excento de IR
                'TotalDevengado = TotalDevengado + MontoDestajo + MontoHRSExtras + MontoComisiones + MontoIncentivos
        Else
        
        
        
        End If


CalcularMontoIr = MontoIr


End Function



Public Function TotalSueldoAnterior(NumeronNomina As Double, CodigoEmpleado As String, Año As Double, Mes As String)
 Dim SqlString As String
 
 TotalSueldoAnterior = 0
 
 SqlString = "SELECT  DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.HE) AS HE, SUM(DetalleNomina.DD) AS DD, SUM(DetalleNomina.HorasExtras) AS HorasExtras, SUM(DetalleNomina.Comisiones) AS Comisiones, SUM(DetalleNomina.OtrosIngresos) AS OtrosIngresos, SUM(DetalleNomina.Incentivos) AS Incentivos, SUM(DetalleNomina.Deducciones) AS Deducciones, SUM(DetalleNomina.Prestamo) AS Prestamo, SUM(DetalleNomina.MontoINSS) AS MontoINSS, SUM(DetalleNomina.MontoIR) AS MontoIR, SUM(DetalleNomina.Vacaciones) AS Vacaciones, SUM(DetalleNomina.INSSPatronal) AS INSSPatronal, SUM(DetalleNomina.IRPatronal) AS IRPatronal, SUM(DetalleNomina.INATEC) AS INATEC, SUM(DetalleNomina.Mes13) AS Mes13, SUM(DetalleNomina.DiasDescuento) AS DiasDescuento, SUM(DetalleNomina.Adelantos) AS Adelantos, SUM(DetalleNomina.TotalSubsidio) AS TotalSubsidio, " & _
                      "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos + DetalleNomina.VacacionesPagadas + DetalleNomina.AdelantosVacaciones + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) AS SueldoAnterior, Nomina.Mes, Nomina.Ano, SUM(DetalleNomina.SeptimoDia) AS SeptimoDia, SUM(DetalleNomina.IncetivoProduccion) AS IncetivoProduccion, SUM(DetalleNomina.BonoProduccion) As BonoProduccion FROM DetalleNomina INNER JOIN  Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina Where (Nomina.NumNomina <> " & NumeronNomina & ") GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano HAVING (DetalleNomina.CodEmpleado = '" & CodigoEmpleado & "') AND (Nomina.Ano = " & Año & ") AND (Nomina.Mes = " & Mes & ")"
 Me.DtaConsulta.RecordSource = SqlString
 Me.DtaConsulta.Refresh
 If Not Me.DtaConsulta.Recordset.EOF Then
   TotalSueldoAnterior = Me.DtaConsulta.Recordset("SueldoAnterior")
 End If
 
  

End Function
Public Function TotalInssAnterior(NumeronNomina As Double, CodigoEmpleado As String, Año As Double, Mes As String)
 Dim SqlString As String
 
 TotalInssAnterior = 0
 
  SqlString = "SELECT  DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.HE) AS HE, SUM(DetalleNomina.DD) AS DD, SUM(DetalleNomina.HorasExtras) AS HorasExtras, SUM(DetalleNomina.Comisiones) AS Comisiones, SUM(DetalleNomina.OtrosIngresos) AS OtrosIngresos, SUM(DetalleNomina.Incentivos) AS Incentivos, SUM(DetalleNomina.Deducciones) AS Deducciones, SUM(DetalleNomina.Prestamo) AS Prestamo, SUM(DetalleNomina.MontoINSS) AS MontoINSS, SUM(DetalleNomina.MontoIR) AS MontoIR, SUM(DetalleNomina.Vacaciones) AS Vacaciones, SUM(DetalleNomina.INSSPatronal) AS INSSPatronal, SUM(DetalleNomina.IRPatronal) AS IRPatronal, SUM(DetalleNomina.INATEC) AS INATEC, SUM(DetalleNomina.Mes13) AS Mes13, SUM(DetalleNomina.DiasDescuento) AS DiasDescuento, SUM(DetalleNomina.Adelantos) AS Adelantos, SUM(DetalleNomina.TotalSubsidio) AS TotalSubsidio, " & _
                      "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos + DetalleNomina.VacacionesPagadas + DetalleNomina.AdelantosVacaciones + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) AS SueldoAnterior, Nomina.Mes, Nomina.Ano, SUM(DetalleNomina.SeptimoDia) AS SeptimoDia, SUM(DetalleNomina.IncetivoProduccion) AS IncetivoProduccion, SUM(DetalleNomina.BonoProduccion) As BonoProduccion FROM DetalleNomina INNER JOIN  Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina Where (Nomina.NumNomina <> " & NumeronNomina & ") GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano HAVING (DetalleNomina.CodEmpleado = '" & CodigoEmpleado & "') AND (Nomina.Ano = " & Año & ") AND (Nomina.Mes = " & Mes & ")"

 Me.DtaConsulta.RecordSource = SqlString
 Me.DtaConsulta.Refresh
 If Not Me.DtaConsulta.Recordset.EOF Then
   TotalInssAnterior = Me.DtaConsulta.Recordset("MontoINSS")
 End If
 
  

End Function

Public Function TotalIrAnterior(NumeronNomina As Double, CodigoEmpleado As String, Año As Double, Mes As String)
 Dim SqlString As String
 
 TotalIrAnterior = 0
 
  SqlString = "SELECT  DetalleNomina.CodEmpleado, SUM(DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.HE) AS HE, SUM(DetalleNomina.DD) AS DD, SUM(DetalleNomina.HorasExtras) AS HorasExtras, SUM(DetalleNomina.Comisiones) AS Comisiones, SUM(DetalleNomina.OtrosIngresos) AS OtrosIngresos, SUM(DetalleNomina.Incentivos) AS Incentivos, SUM(DetalleNomina.Deducciones) AS Deducciones, SUM(DetalleNomina.Prestamo) AS Prestamo, SUM(DetalleNomina.MontoINSS) AS MontoINSS, SUM(DetalleNomina.MontoIR) AS MontoIR, SUM(DetalleNomina.Vacaciones) AS Vacaciones, SUM(DetalleNomina.INSSPatronal) AS INSSPatronal, SUM(DetalleNomina.IRPatronal) AS IRPatronal, SUM(DetalleNomina.INATEC) AS INATEC, SUM(DetalleNomina.Mes13) AS Mes13, SUM(DetalleNomina.DiasDescuento) AS DiasDescuento, SUM(DetalleNomina.Adelantos) AS Adelantos, SUM(DetalleNomina.TotalSubsidio) AS TotalSubsidio, " & _
                      "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos + DetalleNomina.VacacionesPagadas + DetalleNomina.AdelantosVacaciones + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) AS SueldoAnterior, Nomina.Mes, Nomina.Ano, SUM(DetalleNomina.SeptimoDia) AS SeptimoDia, SUM(DetalleNomina.IncetivoProduccion) AS IncetivoProduccion, SUM(DetalleNomina.BonoProduccion) As BonoProduccion FROM DetalleNomina INNER JOIN  Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina Where (Nomina.NumNomina <> " & NumeronNomina & ") GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano HAVING (DetalleNomina.CodEmpleado = '" & CodigoEmpleado & "') AND (Nomina.Ano = " & Año & ") AND (Nomina.Mes = " & Mes & ")"

 Me.DtaConsulta.RecordSource = SqlString
 Me.DtaConsulta.Refresh
 If Not Me.DtaConsulta.Recordset.EOF Then
   TotalIrAnterior = Me.DtaConsulta.Recordset("MontoIR")
 End If
 
  

End Function


Public Sub CodigoRetirado()
'             'DATOS DEL SALARIO MINIMO Y EL VALOR X PUNTOS
'             Me.DtaConsulta.RecordSource = "SELECT [SalarioMinimo], [ValorPts] FROM [dbo].[DatosEmpresa]"
'             Me.DtaConsulta.Refresh
'             If Not Me.DtaConsulta.Recordset.EOF Then
'              If Not IsNull(Me.DtaConsulta.Recordset("salariominimo")) Then
'               SalarioMinimo = Me.DtaConsulta.Recordset("salariominimo")
'              Else
'               SalarioMinimo = 0
'              End If
'                If Not IsNull(Me.DtaConsulta.Recordset("valorpts")) Then
'                  ValorPunto = Me.DtaConsulta.Recordset("valorpts")
'                Else
'                  ValorPunto = 0
'                End If
'             Else
'               SalarioMinimo = 0
'               ValorPunto = 0
'             End If
             
             'PORCIENTO DEL SALARIO
'             Me.DtaConsulta.RecordSource = "SELECT  * From Empleado Where (CodEmpleado = " & CodEmpleado & ")"
'             Me.DtaConsulta.Refresh
'             If Not Me.DtaConsulta.Recordset.EOF Then
'              If Not IsNull(Me.DtaConsulta.Recordset("salporcentaje")) Then
'                 SalarioPorciento = Me.DtaConsulta.Recordset("salporcentaje")
'              Else
'                 SalarioPorciento = 0
'              End If
'             End If
        
   
            'DATOS DE TODOS LOS PUNTOS ADQUIRIDOS X EMPLEADO
'            Me.DtaConsulta.RecordSource = "SELECT  SUM(Puntos.CantPts) AS CantPts FROM  PuntosEmpleado INNER JOIN  Puntos ON PuntosEmpleado.Puntos = Puntos.Id  Where (PuntosEmpleado.Empleado = " & CodEmpleado & ") And (PuntosEmpleado.Aprobado = 1)"
'            Me.DtaConsulta.Refresh
'            If Not Me.DtaConsulta.Recordset.EOF Then
'              If Not IsNull(Me.DtaConsulta.Recordset("CantPts")) Then
'               TotalPuntos = Me.DtaConsulta.Recordset("CantPts")
'              Else
'               TotalPuntos = 0
'              End If
'            Else
'               TotalPuntos = 0
'            End If
End Sub




Public Sub ActualizaTipoNomina()
    'DtaNomina.Recordset.Edit
    DtaNomina.Recordset("TotalSalarioBasico") = 0
    DtaNomina.Recordset("TotalDestajo") = 0
    DtaNomina.Recordset("TotalHorasExtras") = 0
    DtaNomina.Recordset("TotalComisiones") = 0
    DtaNomina.Recordset("TotalIncentivos") = 0
    DtaNomina.Recordset("TotalDeducciones") = 0
    DtaNomina.Recordset("TotalPrestamo") = 0
    DtaNomina.Recordset("TotalMontoInss") = 0
    DtaNomina.Recordset("TotalMontoIR") = 0
    DtaNomina.Recordset("TotalVacaciones") = 0
    DtaNomina.Recordset("TotalINSSPatronal") = 0
    DtaNomina.Recordset("TotalIRPatronal") = 0
    DtaNomina.Recordset("Totalmes13") = 0
    DtaNomina.Recordset("TotalInatec") = 0
    DtaNomina.Recordset("Procesada") = 1
    Me.DtaNomina.Recordset.Update
End Sub



Public Function BuscaDeduccionPorFalta(TipoPeriodo As String, TotalDevengado As Double, MontoVacaciones As Double, DiasDescuento As Double, DiasMes As Double) As Double
            Select Case TipoPeriodo
                Case "Semanal Viernes"
                    BuscaDeduccionPorFalta = TotalDevengado - ((TotalDevengado - MontoVacaciones) * (7 - DiasDescuento) / 7)
                Case "Semanal Sabado"
                    BuscaDeduccionPorFalta = TotalDevengado - ((TotalDevengado - MontoVacaciones) * (7 - DiasDescuento) / 7)
                Case "Catorcenal los Viernes"
                    BuscaDeduccionPorFalta = TotalDevengado - ((TotalDevengado - MontoVacaciones) * (14 - DiasDescuento) / 14)
                Case "Catorcenal los Sabados"
                    BuscaDeduccionPorFalta = TotalDevengado - ((TotalDevengado - MontoVacaciones) * (14 - DiasDescuento) / 14)
                Case "Quincenal"
                    BuscaDeduccionPorFalta = (TotalDevengado / 15) * DtaEmpleados.Recordset("DiasDescuento")
                Case "Mensual"
                    BuscaDeduccionPorFalta = TotalDevengado - ((TotalDevengado - MontoVacaciones) * (DiasMes - DiasDescuento) / DiasMes)
                Case "Trimestral"
                    BuscaDeduccionPorFalta = TotalDevengado - ((TotalDevengado - MontoVacaciones) * ((DiasMes * 3) - DiasDescuento) / (DiasMes * 3))
                Case "Semestral"
                    BuscaDeduccionPorFalta = TotalDevengado - ((TotalDevengado - MontoVacaciones) * ((DiasMes * 6) - DiasDescuento) / (DiasMes * 6))
           End Select
End Function
 
 Public Function BuscaTarifa(CodigoEmpleado As String, NumeroNomina As String) As Double
                    Me.DtaHorasProducidas.RecordSource = "SELECT dbo.DetalleHorasProduccion.CodEmpleado, dbo.DetalleHorasProduccion.NumNomina, dbo.DetalleHorasProduccion.NumLinea, dbo.DetalleHorasProduccion.Lunes + dbo.DetalleHorasProduccion.Martes + dbo.DetalleHorasProduccion.Miercoles + dbo.DetalleHorasProduccion.Jueves + dbo.DetalleHorasProduccion.Viernes AS TotalDias,dbo.Empleado.TarifaHoraria,(dbo.DetalleHorasProduccion.Lunes + dbo.DetalleHorasProduccion.Martes + dbo.DetalleHorasProduccion.Miercoles + dbo.DetalleHorasProduccion.Jueves + dbo.DetalleHorasProduccion.Viernes)* dbo.Empleado.TarifaHoraria AS TotalSalario FROM dbo.DetalleHorasProduccion INNER JOIN dbo.Empleado ON dbo.DetalleHorasProduccion.CodEmpleado = dbo.Empleado.CodEmpleado WHERE (dbo.DetalleHorasProduccion.CodEmpleado = '" & CodigoEmpleado & "')  AND (dbo.DetalleHorasProduccion.Pagado = 0)AND (DetalleHorasProduccion.NumNomina = " & NumeroNomina & ")"
                    Me.DtaHorasProducidas.Refresh
                    TarifaHoraria = 0
                    TotalHoras = 0
                    If Not DtaHorasProducidas.Recordset.EOF Then
                       BuscaTarifa = Me.DtaHorasProducidas.Recordset("TarifaHoraria")
                    End If
 End Function
 Function BuscaSeptimo(CodigoEmpleado As String, NumeroNomina As String) As Double
                    Me.DtaHorasProducidas.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, MAX(DetalleHorasProduccion.NumLinea) AS NumLinea, SUM(DetalleHorasProduccion.Lunes + DetalleHorasProduccion.Martes + DetalleHorasProduccion.Miercoles + DetalleHorasProduccion.Jueves + DetalleHorasProduccion.Viernes + DetalleHorasProduccion.Sabado + DetalleHorasProduccion.Domingo) AS TotalDias, SUM(Empleado.TarifaHoraria) AS TarifaHoraria, SUM((DetalleHorasProduccion.Lunes + DetalleHorasProduccion.Martes + DetalleHorasProduccion.Miercoles + DetalleHorasProduccion.Jueves + DetalleHorasProduccion.Viernes + DetalleHorasProduccion.Sabado + DetalleHorasProduccion.Domingo) * Empleado.TarifaHoraria) AS TotalSalario FROM DetalleHorasProduccion INNER JOIN Empleado ON DetalleHorasProduccion.CodEmpleado = Empleado.CodEmpleado Where (DetalleHorasProduccion.Pagado = 0) GROUP BY DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina " & _
                                                         "HAVING (DetalleHorasProduccion.CodEmpleado = '" & CodigoEmpleado & "') AND (DetalleHorasProduccion.NumNomina = " & NumeroNomina & ")"
                    Me.DtaHorasProducidas.Refresh
                    TarifaHoraria = 0
                    TotalHoras = 0
                    If Not DtaHorasProducidas.Recordset.EOF Then
                        TarifaHoraria = Me.DtaHorasProducidas.Recordset("TarifaHoraria")
                        If Not IsNull(Me.DtaHorasProducidas.Recordset("TotalDias")) Then
                          TotalHoras = Me.DtaHorasProducidas.Recordset("TotalDias")
                        End If
                        If TotalHoras < HoraSeptimo Then
                          If Not IsNull(Me.DtaHorasProducidas.Recordset("TotalSalario")) Then
                            TotalSalarioxHora = Me.DtaHorasProducidas.Recordset("TotalSalario")
                          End If
                        Else
                          '///////////Calculo el salario sumando 48 Horas + 7mo dia 8 Horas
                          TotalSalarioxHora = TarifaHoraria * TotalHoras
                          Septimo = TarifaHoraria * 8
                        End If
                    End If
                    
                    BuscaSeptimo = Septimo
 End Function
 Function BuscaTotalHoras(CodigoEmpleado As String, NumeroNomina As String) As Double
                    Me.DtaHorasProducidas.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, MAX(DetalleHorasProduccion.NumLinea) AS NumLinea, SUM(DetalleHorasProduccion.Lunes + DetalleHorasProduccion.Martes + DetalleHorasProduccion.Miercoles + DetalleHorasProduccion.Jueves + DetalleHorasProduccion.Viernes + DetalleHorasProduccion.Sabado + DetalleHorasProduccion.Domingo) AS TotalDias, SUM(Empleado.TarifaHoraria) AS TarifaHoraria, SUM((DetalleHorasProduccion.Lunes + DetalleHorasProduccion.Martes + DetalleHorasProduccion.Miercoles + DetalleHorasProduccion.Jueves + DetalleHorasProduccion.Viernes + DetalleHorasProduccion.Sabado + DetalleHorasProduccion.Domingo) * Empleado.TarifaHoraria) AS TotalSalario FROM DetalleHorasProduccion INNER JOIN Empleado ON DetalleHorasProduccion.CodEmpleado = Empleado.CodEmpleado Where (DetalleHorasProduccion.Pagado = 0) GROUP BY DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina " & _
                                                         "HAVING (DetalleHorasProduccion.CodEmpleado = '" & CodigoEmpleado & "') AND (DetalleHorasProduccion.NumNomina = " & NumeroNomina & ")"
                    Me.DtaHorasProducidas.Refresh
                    TarifaHoraria = 0
                    TotalHoras = 0
                    If Not DtaHorasProducidas.Recordset.EOF Then
                        TarifaHoraria = Me.DtaHorasProducidas.Recordset("TarifaHoraria")
                        If Not IsNull(Me.DtaHorasProducidas.Recordset("TotalDias")) Then
                          TotalHoras = Me.DtaHorasProducidas.Recordset("TotalDias")
                        End If
                        
                        '////////////SI ES MAYOR DE 48 SEMANAL //////////////////
                        If TotalHoras > 48 Then
                         TotalHoras = 48
                        End If
                        
                        
                        If TotalHoras < HoraSeptimo Then
                          If Not IsNull(Me.DtaHorasProducidas.Recordset("TotalSalario")) Then
                            TotalSalarioxHora = Me.DtaHorasProducidas.Recordset("TotalSalario")
                          End If
                        Else
                          '///////////Calculo el salario sumando 48 Horas + 7mo dia 8 Horas
                          TotalSalarioxHora = TarifaHoraria * TotalHoras
                          Septimo = TarifaHoraria * 8
                        End If
                    End If
                    
                    BuscaTotalHoras = TotalHoras
 End Function
Function BuscaTotalSeptimoSemana(CodigoEmpleado As String, NumeroNomina As String, HorasSeptimo As Double, Semana As Integer) As Double
        Dim Septimo As Double
                    Me.DtaHorasProducidas.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, DetalleHorasProduccion.NumLinea, DetalleHorasProduccion.Lunes + DetalleHorasProduccion.Martes + DetalleHorasProduccion.Miercoles + DetalleHorasProduccion.Jueves + DetalleHorasProduccion.Viernes + DetalleHorasProduccion.Sabado + DetalleHorasProduccion.Domingo AS TotalDias, Empleado.TarifaHoraria, (DetalleHorasProduccion.Lunes + DetalleHorasProduccion.Martes + DetalleHorasProduccion.Miercoles + DetalleHorasProduccion.Jueves + DetalleHorasProduccion.Viernes + DetalleHorasProduccion.Sabado + DetalleHorasProduccion.Domingo) * Empleado.TarifaHoraria AS TotalSalario FROM  DetalleHorasProduccion INNER JOIN  Empleado ON DetalleHorasProduccion.CodEmpleado = Empleado.CodEmpleado WHERE (DetalleHorasProduccion.Pagado = 0) AND (DetalleHorasProduccion.CodEmpleado = '" & CodigoEmpleado & "') AND (DetalleHorasProduccion.NumNomina = " & NumeroNomina & ")"
                    Me.DtaHorasProducidas.Refresh
                    TarifaHoraria = 0
                    TotalHoras = 0
                    If Not DtaHorasProducidas.Recordset.EOF Then
                      If Semana = 1 Then
                        Me.DtaHorasProducidas.Recordset.MoveFirst
                      ElseIf Semana = 2 Then
                        Me.DtaHorasProducidas.Recordset.MoveFirst
                        Me.DtaHorasProducidas.Recordset.MoveLast
                      End If
                        TarifaHoraria = Me.DtaHorasProducidas.Recordset("TarifaHoraria")
                        If Not IsNull(Me.DtaHorasProducidas.Recordset("TotalDias")) Then
                          TotalHoras = Me.DtaHorasProducidas.Recordset("TotalDias")
                        End If
                        If TotalHoras < HorasSeptimo Then
                          If Not IsNull(Me.DtaHorasProducidas.Recordset("TotalSalario")) Then
                            TotalSalarioxHora = Me.DtaHorasProducidas.Recordset("TotalSalario")
                          End If
                          Septimo = 0
                        Else
                          '///////////Calculo el salario sumando 48 Horas + 7mo dia 8 Horas
                          TotalSalarioxHora = TarifaHoraria * TotalHoras
                          Septimo = TarifaHoraria * 8
                        End If
                    End If
                    
                      BuscaTotalSeptimoSemana = Septimo

 End Function
 
 
 Function BuscaTotalSalarioxHora(CodigoEmpleado As String, NumeroNomina As String) As Double
                    Me.DtaHorasProducidas.RecordSource = "SELECT DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina, MAX(DetalleHorasProduccion.NumLinea) AS NumLinea, SUM(DetalleHorasProduccion.Lunes + DetalleHorasProduccion.Martes + DetalleHorasProduccion.Miercoles + DetalleHorasProduccion.Jueves + DetalleHorasProduccion.Viernes + DetalleHorasProduccion.Sabado + DetalleHorasProduccion.Domingo) AS TotalDias, SUM(Empleado.TarifaHoraria) AS TarifaHoraria, SUM((DetalleHorasProduccion.Lunes + DetalleHorasProduccion.Martes + DetalleHorasProduccion.Miercoles + DetalleHorasProduccion.Jueves + DetalleHorasProduccion.Viernes + DetalleHorasProduccion.Sabado + DetalleHorasProduccion.Domingo) * Empleado.TarifaHoraria) AS TotalSalario FROM DetalleHorasProduccion INNER JOIN Empleado ON DetalleHorasProduccion.CodEmpleado = Empleado.CodEmpleado Where (DetalleHorasProduccion.Pagado = 0) GROUP BY DetalleHorasProduccion.CodEmpleado, DetalleHorasProduccion.NumNomina " & _
                                                         "HAVING (DetalleHorasProduccion.CodEmpleado = '" & CodigoEmpleado & "') AND (DetalleHorasProduccion.NumNomina = " & NumeroNomina & ")"
                    Me.DtaHorasProducidas.Refresh
                    TarifaHoraria = 0
                    TotalHoras = 0
                    If Not DtaHorasProducidas.Recordset.EOF Then
                        TarifaHoraria = Me.DtaHorasProducidas.Recordset("TarifaHoraria")
                        If Not IsNull(Me.DtaHorasProducidas.Recordset("TotalDias")) Then
                          TotalHoras = Me.DtaHorasProducidas.Recordset("TotalDias")
                        End If
                        If TotalHoras < HoraSeptimo Then
                          If Not IsNull(Me.DtaHorasProducidas.Recordset("TotalSalario")) Then
                            TotalSalarioxHora = Me.DtaHorasProducidas.Recordset("TotalSalario")
                          End If
                        Else
                          '///////////Calculo el salario sumando 48 Horas + 7mo dia 8 Horas
                          TotalSalarioxHora = TarifaHoraria * TotalHoras
                          Septimo = TarifaHoraria * 8
                        End If
                    End If
                    
                    BuscaTotalSalarioxHora = TotalSalarioxHora
 End Function
 Function BuscaHoraNomina(CodTipoNomina As String) As Double
        MDIPrimero.DtaConsulta.RecordSource = "SELECT  * From TipoNomina WHERE (CodTipoNomina = '" & CodTipoNomina & "')"
        MDIPrimero.DtaConsulta.Refresh
        If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
          BuscaHoraNomina = MDIPrimero.DtaConsulta.Recordset("Horas")
        Else
          BuscaHoraNomina = 0
        End If
 End Function
 
 
 
 Function BuscaMontoHora(TipoPeriodo As String, TarifaHoraria As Double, DiaSemana, SueldoPeriodo, Salario, DiasMes, CodTipoNomina As String) As Double
        Dim Horas As Double
        
        MDIPrimero.DtaConsulta.RecordSource = "SELECT  * From TipoNomina WHERE (CodTipoNomina = '" & CodTipoNomina & "')"
        MDIPrimero.DtaConsulta.Refresh
        If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
          Horas = MDIPrimero.DtaConsulta.Recordset("Horas")
        End If
        
   
        
        
        
        Select Case TipoPeriodo
        
        Case "Semanal Viernes"
            Dim SalarioMes As Double, SalarioDia As Double, SalarioHora As Double
            SalarioMes = Salario * 4.33 '(52 / 12)  Factor se obtiene dividiendo
            SalarioDia = SalarioMes / DiasMes
            SalarioHora = SalarioDia / Horas
            BuscaMontoHora = Format(SalarioHora, "###,##0.000000")
        Case "Semanal Sabado"
            BuscaMontoHora = Format(TarifaHoraria, "###,##0.000000")
        Case "Catorcenal los Viernes"
           If SueldoPeriodo <> 0 Then
             BuscaMontoHora = Format(SueldoPeriodo / (DiasMes * Horas), "###,##0.0000")
           Else
            BuscaMontoHora = Format(TarifaHoraria, "###,##0.0000")
           End If
        Case "Catorcenal los Sabados"
           If SueldoPeriodo <> 0 Then
             BuscaMontoHora = Format(SueldoPeriodo / 14 / Horas, "###,##0.0000")
           ElseIf Salario = 0 Then
              BuscaMontoHora = Format(TarifaHoraria, "###,##0.0000")
            Else
              BuscaMontoHora = Format(Salario / 14 / Horas, "###,##0.0000")
            End If
           
        Case "Quincenal"
            BuscaMontoHora = Format((Salario / 15) / Horas, "###,##0.000000")
            'BuscaMontoHora = Format(((Salario) * 2) / (DiasMes * Horas), "###,##0.000000")
'            MontoHora = Format(DtaEmpleados.Recordset("SueldoPeriodo") / ((DiasMes * 8) / 2), "###,##0.000000")
        Case "Mensual"
            BuscaMontoHora = Format(DtaEmpleados.Recordset("SueldoPeriodo") / (DiasMes * Horas), "###,##0.00")
        Case "Trimestral"
            BuscaMontoHora = Format(DtaEmpleados.Recordset("SueldoPeriodo") / (DiasMes * Horas * 3), "###,##0.00")
        Case "Semestral"
            BuscaMontoHora = Format(DtaEmpleados.Recordset("SueldoPeriodo") / (DiasMes * Horas * 6), "###,##0.00")
        End Select


End Function
 Function BuscaMontoDia(TipoPeriodo As String, TarifaHoraria As Double, DiaSemana, SueldoPeriodo, Salario, DiasMes, CodTipoNomina As String) As Double
        Dim Horas As Double
        
        MDIPrimero.DtaConsulta.RecordSource = "SELECT  * From TipoNomina WHERE (CodTipoNomina = '" & CodTipoNomina & "')"
        MDIPrimero.DtaConsulta.Refresh
        If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
          Horas = MDIPrimero.DtaConsulta.Recordset("Horas")
        End If
        
      
        
        
        
        Select Case TipoPeriodo
        
        Case "Semanal Viernes"
            Dim SalarioMes As Double, SalarioDia As Double, SalarioHora As Double
            SalarioMes = Salario * 4.33 '(52 / 12)  Factor se obtiene dividiendo
            SalarioDia = SalarioMes / DiasMes
            SalarioHora = SalarioDia / Horas
            BuscaMontoDia = Format(SalarioHora, "###,##0.000000")
        Case "Semanal Sabado"
            BuscaMontoDia = Format(TarifaHoraria, "###,##0.000000")
        Case "Catorcenal los Viernes"
          If SueldoPeriodo <> 0 Then
             BuscaMontoDia = Format(DtaEmpleados.Recordset("SueldoPeriodo") / 14, "###,##0.00")
          Else
             BuscaMontoDia = Format(TarifaHoraria * Horas, "###,##0.0000")
          End If
        Case "Catorcenal los Sabados"
          If SueldoPeriodo <> 0 Then
             BuscaMontoDia = Format(DtaEmpleados.Recordset("SueldoPeriodo") / 14, "###,##0.00")
          Else
             BuscaMontoDia = Format(TarifaHoraria * Horas, "###,##0.0000")
          End If
            
        Case "Quincenal"
            BuscaMontoDia = Format((Salario * 2) / DiasMes, "###,##0.000000")
'            MontoHora = Format(DtaEmpleados.Recordset("SueldoPeriodo") / ((DiasMes * 8) / 2), "###,##0.000000")
        Case "Mensual"
            BuscaMontoDia = Format(DtaEmpleados.Recordset("SueldoPeriodo") / (DiasMes), "###,##0.00")
        Case "Trimestral"
            BuscaMontoDia = Format(DtaEmpleados.Recordset("SueldoPeriodo") / (DiasMes), "###,##0.00")
        Case "Semestral"
            BuscaMontoDia = Format(DtaEmpleados.Recordset("SueldoPeriodo") / (DiasMes), "###,##0.00")
        End Select


End Function

Sub CreateTaskPanel()


    Dim Group As TaskPanelGroup
    Dim item As TaskPanelGroupItem
    
    Set Group = wndTaskPanel.Groups.Add(100, "Procesos")
    Group.Tooltip = "Registro de Laboratorio"
    Group.Special = True
    Group.Items.Add 1, "Estatus Examenes", xtpTaskItemTypeLink, 2
    Group.Items.Add 2, "Solicitud Examenes", xtpTaskItemTypeLink, 8

    
    Set Group = wndTaskPanel.Groups.Add(100, "Catalogos")
    Group.Tooltip = "Catalogo del sistema Contable"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 6, "Doctores", xtpTaskItemTypeLink, 6
    Group.Items.Add 7, "Pacientes", xtpTaskItemTypeLink, 7
    Group.Items.Add 8, "Usuarios", xtpTaskItemTypeLink, 10

    

    
    Set Group = wndTaskPanel.Groups.Add(100, "Opciones")
    Group.Tooltip = "Procesos del Sistema Contable"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 13, "Calculadora", xtpTaskItemTypeLink, 18
    Group.Items.Add 13, "Informacion de Usuarios", xtpTaskItemTypeLink, 19
    Group.Items.Add 13, "Configuracion", xtpTaskItemTypeLink, 28
    Group.Items.Add 13, "Respaldar", xtpTaskItemTypeLink, 29
    
    Set Group = wndTaskPanel.Groups.Add(100, "Reportes")
    Group.Tooltip = "Procesos del Sistema Contable"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 13, "Reportes Generales", xtpTaskItemTypeLink, 20
    Group.Items.Add 13, "Reportes de Movimientos", xtpTaskItemTypeLink, 21

    
   
    
     
'    wndTaskPanel.SetImageList Me.ImageList1
End Sub







Private Sub CmdCalcular_Click()
'On Error GoTo TipoErrs
Dim CodIncentivos As String, IrUltimaSemana As Boolean

k% = MsgBox("Desea Procesar la Nómina " + DtaTipoNomina.Recordset("nomina") + " " + DtaTipoNomina.Recordset("Periodo") + "?", vbYesNo)
If k <> 6 Then Exit Sub
Me.MousePointer = 11
Dim i As Integer, sql As String, TarifaHorariaBasico As Double, FechaIni As String, TotalSalarioxHora As Double, TasaCambio As Double, SQLNominaEmpleado As String, TarifaHoraria As Double, SQLNomina As String, TotalHoras As Double, CodDepartamento As String, NumNomina As String, MontoIncentivoHoras As Double, CodTipoNomina As String, PorcientoIncentivo As Double, CodEmpleado As String, TotalPuntualidad As Double, SQlIncentivos As String, Septimo As Double, SQlDeducciones As String, MontoIrAcumulado As Double, SQlPrestamo As String, SQlComisiones As String, SQlDestajo As String, SqlHrsExtras As String, SueldoPeriodo As Double, TasaInss As Double, TasaInssPatronal As Double, MontoIncentivos As Double, TasaIr As Double, MontoDeduccion As Double, MontoPrestamo As Double, MontoComisiones As Double, FechaContrato As Date, MontoDestajos As Double, annos As Date, MontoHRSExtras As Double, Antiguedad As Double
Dim MontoOtrosIngresos As Double, PorcientoAntiguedad As Double, DescripOtrIngre As String, NumFecha1 As Date, MontoHora As Double, CodProceso As String, CantEmpleados As Long, CodReferencia As String, MontoIr As Double, UnidadesProducidas As Double, Rango As Double, MontoInss As Double, Monto As Double, MontoIRPatronal As Double, NumeroDeduccion As Double, MontoInssPatronal As Double, MontoBrutoAnual As Double, MontoVacaciones As Double, Nombres As String, MontoMes13 As Double, FechaNomina As Date, DeduccionPorFalta As Double, SeptimoAnterior As Double, MinIR As Double, AñoFiscal As Double, SalarioMensual As Double, RentaGravable As Double, DiasMes As Double, TotalDevengadoAcumulado As Double, DiasSemana As Double, IncentivoProduccion As Double, CantSabados As Byte, IdDeduccion As Double, TotalDevengado As Double, PagoProduccion As Double, SalHora As Double, NumeroPeriodo As Double, PeriodoFiscal As Double, Factor As Double, NQuincenas As Double, INATEC As Double, FechaInicialIr As Date
Dim FechaFinalIr As Date, DevengadoSinHrsExtras As Double, VacacionesAcumuladas As Double, HE As Single, HoraPuntualidad As Double, MontoPuntualidad As Double, DD As Single, HoraSeptimo As Double, HoraBasico As Double, FormatoNomina As String, Adelantos As Double, Anos As Double, Moneda As String, MontoDolares As Double, MontoProduccion As Double, agregar As Boolean, FechaIngreso As Date, PeriodoIngreso As Double, BonoProduccion As Double, MontoViaticos As Double, NumIncentivo As Double, FechaInicio As String, FechaFin As String, Mes As Double, Fecha As String, TipoCalculoIr As String, Calcular7mo As Boolean, Dolarizado As Boolean, cn As New ADODB.Connection, ValorPunto As Double, SalarioMinimo As Double, SalarioPorciento As Double, rs As New ADODB.Recordset, TotalPuntos As Double, SalarioBasico As Double, CalcularPuntos As Boolean, MontoInssBasico As Double, AjusteINSS As Double, cmd As New ADODB.Command, HT As Double, MontoHorasTurno As Double, UltimaSemana As Boolean
Dim TipoVacaciones As Boolean, MontoTipoVacaciones As Double, CalcularHorasTurno As Boolean, MontoIncentivoExcento As Double
Dim DiasBasico As Double, MontoBasico As Double, MontoDia As Double, AumentoBasico As Double, MontoViaticoEmpleado As Double
Dim PAntiguedad As Double, DiasAdicionales As Double, MontoDiasAdicionales As Double, MontoSubsidio As Double, DiasSubsidio As Double
Dim DiasVacaciones As Double, ValorViaticoxDia As Double, Reembolso As Double, DiasHorasExtra As Double



    CalcularHorasTurno = False
    MT = 0
    MontoHorasTurno = 0
    HoraSeptimo = 0
    HoraPuntualidad = 0
    HoraBasico = 0
    MontoPuntualidad = 0
    MontoViaticos = 0
    MontoViaticoEmpleado = 0
    Calcular7mo = True

 MDIPrimero.DtaEmpresa.Refresh
 If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
   TipoCalculoIr = MDIPrimero.DtaEmpresa.Recordset("TipoCalculoIR")
   If MDIPrimero.DtaEmpresa.Recordset("Calcular7mo") = True Then
     Calcular7mo = True
   Else
     Calcular7mo = False
   End If
   
   If MDIPrimero.DtaEmpresa.Recordset("CalcularPuntos") = True Then
     CalcularPuntos = True
   Else
     CalcularPuntos = False
   End If
   
   If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("HorasExtra")) Then
    DiasHorasExtra = MDIPrimero.DtaEmpresa.Recordset("HorasExtra")
   Else
     DiasHorasExtra = 2
   End If
   
 End If
 
 

Me.CmdExportar.Enabled = True
Me.CmdExportaCSV.Enabled = True
'///////////////Busco la Configuracion de Incentivo de Puntualidad / Spetimo Dia / Basico
 If Not Me.AdoConfiguracion.Recordset.EOF Then
   If Not IsNull(Me.AdoConfiguracion.Recordset("HorasPuntualidad")) Then
      HoraPuntualidad = Me.AdoConfiguracion.Recordset("HorasPuntualidad")
   End If
   If Not IsNull(Me.AdoConfiguracion.Recordset("HorasSeptimo")) Then
      HoraSeptimo = Me.AdoConfiguracion.Recordset("HorasSeptimo")
   End If
   
    If Not IsNull(Me.AdoConfiguracion.Recordset("HorasBasico")) Then
      HoraBasico = Me.AdoConfiguracion.Recordset("HorasBasico")
   End If
   
    If Not IsNull(Me.AdoConfiguracion.Recordset("MontoPuntualidad")) Then
      MontoPuntualidad = Me.AdoConfiguracion.Recordset("MontoPuntualidad")
   End If
   
  If Not IsNull(Me.AdoConfiguracion.Recordset("MontoViaticos")) Then
      MontoViaticos = Me.AdoConfiguracion.Recordset("MontoViaticos")
  Else
      MontoViaticos = 0
  End If
   
 End If


Dim FechaInicioNomina As String, dFechaInicioNomina As Date
Dim FechaFinNomina As String, dFechaFinNomina As Date

'//////////////// Cargo el Tipo de Nomina a la Consulta////////////////////
'/////////Ubico la nomina Actual Activa/////////////////////////
CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
IrUltimaSemana = DtaTipoNomina.Recordset("IrUltimaSemana")
SQLNomina = "SELECT Nomina.* From Nomina WHERE Nomina.Activa=1 AND Nomina.CodTipoNomina= '" & CodTipoNomina & "'"

DtaNomina.RecordSource = SQLNomina
DtaNomina.Refresh

Periodo = DtaNomina.Recordset("Periodo")
   FechaInicioNomina = Format(DtaNomina.Recordset("FechaNominaINI"), "yyyy") & Format(DtaNomina.Recordset("FechaNominaINI"), "MM") & Format(DtaNomina.Recordset("FechaNominaINI"), "dd")
   FechaFinNomina = Format(DtaNomina.Recordset("FechaNomina"), "yyyy") & Format(DtaNomina.Recordset("FechaNomina"), "MM") & Format(DtaNomina.Recordset("FechaNomina"), "dd")
   dFechaInicioNomina = DtaNomina.Recordset("FechaNominaINI")
   dFechaFinNomina = DtaNomina.Recordset("FechaNomina")

    Me.DtaConsulta.RecordSource = "SELECT * From TipoNomina WHERE (CodTipoNomina = '" & CodTipoNomina & "')"
    Me.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
     If Not IsNull(Me.DtaConsulta.Recordset("TarifaHoraria")) Then
      TarifaHorariaBasico = Me.DtaConsulta.Recordset("TarifaHoraria")
      TarifaHoraria = Me.DtaConsulta.Recordset("TarifaHoraria")
     End If
    Else
      TarifaHorariaBasico = 0
      TarifaHoraria = 0
    End If
    
    
    Dim EmpleadoConstruccion As Boolean
    If Not IsNull(Me.DtaConsulta.Recordset("EmpleadoConstruccion")) Then
      EmpleadoConstruccion = Me.DtaConsulta.Recordset("EmpleadoConstruccion")
    Else
      EmpleadoConstruccion = False
    End If

    PTipoNomina = CodTipoNomina
    PFechaNomina = DtaNomina.Recordset("FechaNominaIni")
    Pmes = CStr(Month(DtaNomina.Recordset("FechaNomina")))
    PAno = CInt(Year(DtaNomina.Recordset("FechaNomina")))

'///////////////Verifico si el ultimo dia del mes//////////////////
CodTipoNomina = DtaNomina.Recordset("CodTipoNomina")
NumNomina = DtaNomina.Recordset("NumNomina")

res = Bitacora(Now, NombreUsuario, "Calcular Nomina", "Se Calculo la Nomina: " & NumNomina)


DiaFin = Day(DtaNomina.Recordset("FechaNomina"))
MesIni = Month(DtaNomina.Recordset("FechaNomina"))
AnoIni = Year(DtaNomina.Recordset("FechaNomina"))


Mes = (DtaNomina.Recordset("Mes"))
Fecha = DateSerial(AnoIni, Mes + 1, 0)

'NumFecha2 = CDate(DtaNomina.Recordset("FechaNomina"))
NumFecha2 = DateSerial(AnoIni, Mes + 1, 0)

FechaInicio = Format(DtaNomina.Recordset("FechaNominaINI"), "yyyy-mm-dd")
FechaFin = Format(DtaNomina.Recordset("FechaNomina"), "yyyy-mm-dd")

    If DiaFin >= 28 Then
         '///////Ubico la Fecha de la Quincena Anterior//////////////?////
         '/////////////Cargo la Nomina de la Quincena anterior/////////////
         Me.DtaNominaMes.RecordSource = "SELECT Nomina.* From Nomina Where (((Nomina.CodTipoNomina) = '" & CodTipoNomina & "'))"
         Me.DtaNominaMes.Refresh
         Me.DtaNominaMes.Recordset.MoveLast
        If Not DtaNominaMes.Recordset.EOF Then
         If DtaNominaMes.Recordset("Activa") = True Then
           If DtaNominaMes.Recordset.RecordCount > 1 Then
           DtaNominaMes.Recordset.MovePrevious
           NumNominaAnterior = Me.DtaNominaMes.Recordset("NumNomina")
           NumFecha1 = CDate(DtaNominaMes.Recordset("FechaNomina"))
           End If
         End If
        End If
     
    
    End If




If DtaTipoNomina.Recordset("MantValor") = True Then
   Factor = Tasa
Else
   Factor = 1
End If

'/////////Borro toda la Informacio Anterior/////////////////////

'rs.Open "DELETE FROM [DetalleNomina] Where (NumNomina = " & NumNomina & ")", Conexion

ActualizaTipoNomina

'/////////// Sql Tipo de Nomina////////////////////////////
'SQLNominaEmpleado = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, Empleado.ExentoIr, Empleado.OtrosIngresos,  Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, Empleado.Ausente, PorcientoIncentivo From Empleado WHERE Empleado.CodTipoNomina= '" & CodTipoNomina & "' AND Empleado.Activo= 1 AND Empleado.Ausente=0"
SQLNominaEmpleado = "SELECT Empleado.CodEmpleado1,Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, Empleado.ExentoIr, Empleado.OtrosIngresos,  Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, Empleado.Ausente, PorcientoIncentivo,Empleado.Dolarizado,HorasTurno,CantPts,DiasBasico, AumentoBasico, ViaticoxDia, Reembolso From Empleado WHERE Empleado.CodTipoNomina= '" & CodTipoNomina & "' AND Empleado.Activo= 1 ORDER BY CodEmpleado1"
DtaEmpleados.RecordSource = SQLNominaEmpleado
DtaEmpleados.Refresh
If Me.DtaEmpleados.Recordset.EOF Then
 MsgBox "No Existe Ningun Empleado Asignado a Esta Nomina", vbCritical, "Sistema de Nominas"
 Exit Sub
End If

DtaEmpleados.Recordset.MoveLast
CantEmpleados = DtaEmpleados.Recordset.RecordCount
DtaEmpleados.Recordset.MoveFirst

MsgBox ("Se Procesarán " & CantEmpleados & " Empleados")

DtaControles.Refresh
DiasMes = DtaControles.Recordset("DiasMes")
DiasSemana = DtaControles.Recordset("DiasSemana")


            MDIPrimero.DtaEmpresa.Refresh
            If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
              FormatoNomina = MDIPrimero.DtaEmpresa.Recordset("FormatoNomina")
            End If

'////////////// Inicio el Calculo de la Nomina ////////////////////
'////////////// Actualizo el Control Progress Bar /////////////////
With PBCalcNomina
 .Min = 0
 .Max = CantEmpleados
 .Value = 0
 i = 1
 

Do While Not DtaEmpleados.Recordset.EOF


MontoOtrosIngresos = 0
TotalHoras = 0
TotalSalarioxHora = 0
TarifaHoraria = 0
PagoProduccion = 0
DiasBasico = 0
        AumentoBasico = 0
        If Not IsNull(DtaEmpleados.Recordset("AumentoBasico")) Then
            AumentoBasico = DtaEmpleados.Recordset("AumentoBasico")
        Else
            AumentoBasico = 0
        End If
        
        Reembolso = 0
        If Not IsNull(DtaEmpleados.Recordset("Reembolso")) Then
            Reembolso = DtaEmpleados.Recordset("Reembolso")
        Else
            Reembolso = 0
        End If
        
        
        
        
        Nombres = DtaEmpleados.Recordset("Nombre1") + " " + DtaEmpleados.Recordset("Nombre2") + " " + DtaEmpleados.Recordset("Apellido1") + "  " + DtaEmpleados.Recordset("Apellido2")
        TotalEmpleado = TotalEmpleado + 1
        'Me.xp_canvas1.Caption = "Procesando:  " & i & " de " & CantEmpleados & " Empleados "
        Me.Caption = "Procesando:  " & i & " de " & CantEmpleados & " Empleados "
        Me.LblTotal.Caption = "Procesando:  " & i & " de " & CantEmpleados & " Empleados "
        .Value = i
        DoEvents
        If Me.Label1.Caption = "Procesando la Nomina" Then
         Me.Label1.Caption = "Procesando la Nomina."
        ElseIf Me.Label1.Caption = "Procesando la Nomina." Then
            Me.Label1.Caption = "Procesando la Nomina.."
         ElseIf Me.Label1.Caption = "Procesando la Nomina.." Then
            Me.Label1.Caption = "Procesando la Nomina..."
           ElseIf Me.Label1.Caption = "Procesando la Nomina..." Then
             Me.Label1.Caption = "Procesando la Nomina...."
          ElseIf Me.Label1.Caption = "Procesando la Nomina...." Then
           Me.Label1.Caption = "Procesando la Nomina"
        End If
        
        MDIPrimero.PopupControl1.RemoveAllItems
        
        Set item = MDIPrimero.PopupControl1.AddItem(20, 15, 270, 45, Titulo)
          item.TextColor = RGB(0, 61, 178)
          item.Bold = True
        Set item = MDIPrimero.PopupControl1.AddItem(20, 29, 400, 100, "Calculando:" & DtaEmpleados.Recordset("CodEmpleado1"))
        item.TextColor = RGB(0, 61, 178)
          item.Bold = True
        Set item = MDIPrimero.PopupControl1.AddItem(20, 60, 400, 100, Nombres)
        item.Bold = True
        MDIPrimero.PopupControl1.VisualTheme = xtpPopupThemeOffice2003
        MDIPrimero.PopupControl1.SetSize 300, 110
        MDIPrimero.PopupControl1.Show
        MDIPrimero.PopupControl1.Show
        
        Dim CodempleadoSoli As String
          CodempleadoSoli = DtaEmpleados.Recordset("CodEmpleado1")
          CodEmpleado = DtaEmpleados.Recordset("CodEmpleado")
          

       
        If Not IsNull(Me.DtaEmpleados.Recordset("CodDepartamento")) Then
          CodDepartamento = Me.DtaEmpleados.Recordset("CodDepartamento")
        End If
        PorcientoIncentivo = DtaEmpleados.Recordset("PorcientoIncentivo")
        
        FechaNomina = CDate(DtaNomina.Recordset("FechaNomina"))
        TasaCambio = BuscaTasaCambio(FechaNomina)
'        Me.AdoBusca.RecordSource = "SELECT FechaDia, MontoDia From CambioMoneda WHERE (FechaDia = '" & Format(FechaNomina, "yyyymmdd") & "')"
'        Me.AdoBusca.Refresh
'        If Not Me.AdoBusca.Recordset.EOF Then
'           TasaCambio = Me.AdoBusca.Recordset("MontoDia")
'        Else
'           TasaCambio = 1
'        End If

        If Not IsNull(DtaEmpleados.Recordset("HorasTurno")) Then
          CalcularHorasTurno = DtaEmpleados.Recordset("HorasTurno")
        Else
          CalcularHorasTurno = False
        End If
        
        If Not IsNull(DtaEmpleados.Recordset("Dolarizado")) Then
           Dolarizado = DtaEmpleados.Recordset("Dolarizado")
        Else
           Dolarizado = False
        End If
        
        If Dolarizado = False Then
          TasaCambio = 1
        End If
        
        ValorViaticoxDia = 0
        If Not IsNull(DtaEmpleados.Recordset("ViaticoxDia")) Then
          ValorViaticoxDia = DtaEmpleados.Recordset("ViaticoxDia")
        End If


 
   Select Case DtaTipoNomina.Recordset("Periodo")
       Case "Quincenal"
        If DiaFin >= 28 Then

         Cadena = "SELECT DetalleNomina.id, DetalleNomina.NumNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HE, DetalleNomina.DD, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.INATEC, DetalleNomina.Mes13, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.TotalSubsidio, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[HorasExtras]+[DetalleNomina].[Comisiones]+[DetalleNomina].[OtrosIngresos]+[DetalleNomina].[Incentivos]+[DetalleNomina].[VacacionesPagadas] + [DetalleNomina].[AdelantosVacaciones]" & vbLf
         Cadena = Cadena & "AS SueldoAnterior, DetalleNomina.ValorDiasAdicionales, DetalleNomina.DiasAdicionales From DetalleNomina Where (((DetalleNomina.NumNomina) = " & NumNominaAnterior & " ) And ((DetalleNomina.CodEmpleado) = '" & CodEmpleado & "'))"
         Me.DtaDetalleNominaAnterior.RecordSource = Cadena
         Me.DtaDetalleNominaAnterior.Refresh
         
          If Not Me.DtaDetalleNominaAnterior.Recordset.EOF Then
            TotalDevengadoAnterior = Me.DtaDetalleNominaAnterior.Recordset("SueldoAnterior")
            MontoInssAnterior = Me.DtaDetalleNominaAnterior.Recordset("MontoInss")
            MontoIrAnterior = Me.DtaDetalleNominaAnterior.Recordset("MontoIR")
            MontoInssPatronalAnterior = Me.DtaDetalleNominaAnterior.Recordset("INSSPatronal")
            MontoIrPatronalAnterior = Me.DtaDetalleNominaAnterior.Recordset("IRPatronal")
          Else
            TotalDevengadoAnterior = 0
            MontoInssAnterior = 0
            MontoIrAnterior = 0
            MontoInssPatronalAnterior = 0
            MontoIrPatronalAnterior = 0
          End If
         End If
        Case "Catorcenal los Sabados"
        If DiaFin >= 28 Then

         Cadena = "SELECT DetalleNomina.id, DetalleNomina.NumNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HE, DetalleNomina.DD, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.INATEC, DetalleNomina.Mes13, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.TotalSubsidio, [DetalleNomina].[SalarioBasico]+[DetalleNomina].[Destajo]+[DetalleNomina].[HorasExtras]+[DetalleNomina].[Comisiones]+[DetalleNomina].[OtrosIngresos]+[DetalleNomina].[Incentivos]+[DetalleNomina].[VacacionesPagadas] + [DetalleNomina].[AdelantosVacaciones]" & vbLf
         Cadena = Cadena & "AS SueldoAnterior, DetalleNomina.ValorDiasAdicionales, DetalleNomina.DiasAdicionales From DetalleNomina Where (((DetalleNomina.NumNomina) = " & NumNominaAnterior & " ) And ((DetalleNomina.CodEmpleado) = '" & CodEmpleado & "'))"
         Me.DtaDetalleNominaAnterior.RecordSource = Cadena
         Me.DtaDetalleNominaAnterior.Refresh
         
          If Not Me.DtaDetalleNominaAnterior.Recordset.EOF Then
            TotalDevengadoAnterior = Me.DtaDetalleNominaAnterior.Recordset("SueldoAnterior")
            MontoInssAnterior = Me.DtaDetalleNominaAnterior.Recordset("MontoInss")
            MontoIrAnterior = Me.DtaDetalleNominaAnterior.Recordset("MontoIR")
            MontoInssPatronalAnterior = Me.DtaDetalleNominaAnterior.Recordset("INSSPatronal")
            MontoIrPatronalAnterior = Me.DtaDetalleNominaAnterior.Recordset("IRPatronal")
          Else
            TotalDevengadoAnterior = 0
            MontoInssAnterior = 0
            MontoIrAnterior = 0
            MontoInssPatronalAnterior = 0
            MontoIrPatronalAnterior = 0
          End If
         End If
        
        
        End Select
        
        'creo todos los SQL's
        '/////////////////////////////
'        SQlIncentivos = "SELECT DetalleIncentivo.NumIncentivo, DetalleIncentivo.Valor, DetalleIncentivo.NumVez, DetalleIncentivo.Pagado,DetalleIncentivo.NumNomina, Incentivo.CodEmpleado FROM Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo WHERE DetalleIncentivo.Pagado= 0 AND Incentivo.CodEmpleado= '" & CodEmpleado & "'"
        SQlIncentivos = "SELECT DetalleIncentivo.NumIncentivo, DetalleIncentivo.Valor, DetalleIncentivo.NumVez, DetalleIncentivo.Pagado, DetalleIncentivo.NumNomina, Incentivo.CodEmpleado, TipoIncentivo.Incentivo, TipoIncentivo.CodTipoIncentivo FROM  Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo INNER JOIN TipoIncentivo ON Incentivo.CodTipoIncentivo = TipoIncentivo.CodTipoIncentivo WHERE (DetalleIncentivo.Pagado = 0) AND (Incentivo.CodEmpleado = '" & CodEmpleado & "') AND (DetalleIncentivo.NumNomina = " & NumNomina & ") AND (DetalleIncentivo.NumVez <> 'n')"
        'SQlIncentivos = "SELECT DetalleIncentivo.NumIncentivo, DetalleIncentivo.Valor, DetalleIncentivo.NumVez, DetalleIncentivo.Pagado, DetalleIncentivo.NumNomina, Incentivo.CodEmpleado, TipoIncentivo.Incentivo, TipoIncentivo.CodTipoIncentivo FROM  Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo INNER JOIN TipoIncentivo ON Incentivo.CodTipoIncentivo = TipoIncentivo.CodTipoIncentivo WHERE (DetalleIncentivo.Pagado = 0) AND (Incentivo.CodEmpleado = '" & CodEmpleado & "')"
        DtaIncentivos.RecordSource = SQlIncentivos
        DtaIncentivos.Refresh
        
        '/////////////// Deducciones //////////////////////////
'        SQlDeducciones = "SELECT  MAX(Deduccion.NumDeduccion) AS NumDeduccion, TipoDeduccion.Deduccion, Deduccion.CodEmpleado, AVG(DetalleDeduccion.Valor) AS Valor, COUNT(DetalleDeduccion.NumVez) AS NumVez, Deduccion.NUmNomina FROM TipoDeduccion INNER JOIN  Deduccion INNER JOIN  DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion ON TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion  " & _
'                         "Where (DetalleDeduccion.Pagado = 0) GROUP BY TipoDeduccion.Deduccion, Deduccion.CodEmpleado, Deduccion.NUmNomina Having (Deduccion.CodEmpleado = " & CodEmpleado & ") ORDER BY NumDeduccion"
'        SQlDeducciones = "SELECT id,Deduccion.CodTipoDeduccion, DetalleDeduccion.NumDeduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.Pagado, DetalleDeduccion.NumNomina, Deduccion.CodEmpleado FROM Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion WHERE DetalleDeduccion.Pagado=0 AND Deduccion.CodEmpleado= '" & CodEmpleado & "' AND (DetalleDeduccion.NumNomina = " & NumNomina & ")"
        SQlDeducciones = "SELECT id,Deduccion.CodTipoDeduccion, DetalleDeduccion.NumDeduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.Pagado, DetalleDeduccion.NumNomina, Deduccion.CodEmpleado FROM Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion WHERE DetalleDeduccion.Pagado=0 AND Deduccion.CodEmpleado= '" & CodEmpleado & "' "
        DtaDeducciones.RecordSource = SQlDeducciones
        DtaDeducciones.Refresh
        
        '///////////////Prestamos//////////////////////////
        SQlPrestamo = "SELECT MovPrestamo.NumPrestamo, MovPrestamo.NumCuota, MovPrestamo.CuotaIgual, MovPrestamo.Cancelado, MovPrestamo.NumNomina, Prestamo.CodEmpleado, Prestamo.Moneda FROM Prestamo INNER JOIN MovPrestamo ON Prestamo.NumPrestamo = MovPrestamo.NumPrestamo WHERE MovPrestamo.Cancelado=0 AND Prestamo.CodEmpleado='" & CodEmpleado & "'"
        DtaPrestamo.RecordSource = SQlPrestamo
        DtaPrestamo.Refresh
        
     
               
      
        SQLDestajos = "SELECT TipoDestajo.Destajo, TipoDestajo.Monto, DestalleDestajos.Cantidad, [DestalleDestajos].[Cantidad]*[TipoDestajo].[Monto] AS Total, DestalleDestajos.NUmNomina, DestalleDestajos.CodEmpleado,DestalleDestajos.CodTipoDestajo FROM TipoDestajo INNER JOIN DestalleDestajos ON TipoDestajo.COdTipoDestajo = DestalleDestajos.CodTipoDestajo WHERE DestalleDestajos.CodEmpleado='" & CodEmpleado & "' AND DestalleDestajos.NUmNomina=" & NumNomina & ""
        Me.DtaDestajo.RecordSource = SQLDestajos
'        InputBox "", "", Me.DtaDestajo.RecordSource
        DtaDestajo.Refresh
        
 '//////////////////////////////////////////////////////////////////////////////////////////
 '///////////CALCULO DE VIATICOS SEGUN LA TABLA////////////////////////////////////////////
 '/////////////////////////////////////////////////////////////////////////////////////////
 If Not MontoViaticos = 0 Then
    Me.AdoViaticos.RecordSource = "SELECT * From Incentivo Where (Pagado = 0) And (CodEmpleado = " & CodEmpleado & ")"
    Me.AdoViaticos.Refresh
    
    
    If Me.AdoViaticos.Recordset.EOF Then
     '////////////////////////BUSCAR EL ULTIMO NUMERO DE INCENTIVO///////////////////////////////////////////////////
      Me.AdoBusca.RecordSource = "SELECT NumIncentivo, CodEmpleado, CodTipoIncentivo, NumVeces, Pagado From Incentivo"
      Me.AdoBusca.Refresh
      If AdoBusca.Recordset.EOF Then
         NumIncentivo = 0
      Else
         Me.AdoBusca.Recordset.MoveLast
         NumIncentivo = Me.AdoBusca.Recordset("NumIncentivo") + 1
      End If
    
     '//////////////AGREGO EL ENCABEZADO DEL INCENTIVO///////////////////////////////////////////////
        Me.AdoViaticos.Recordset.AddNew
        AdoViaticos.Recordset("NumIncentivo") = NumIncentivo
        AdoViaticos.Recordset("CodEmpleado") = CodEmpleado
        AdoViaticos.Recordset("CodtipoIncentivo") = "04"
        AdoViaticos.Recordset("numveces") = 1
        AdoViaticos.Recordset("pagado") = False
        AdoViaticos.Recordset.Update
    
        Me.AdoDetalleViaticos.RecordSource = "SELECT * From [dbo].[DetalleIncentivo] "
        Me.AdoDetalleViaticos.Refresh
        
         
            AdoDetalleViaticos.Recordset.AddNew
                AdoDetalleViaticos.Recordset("ID") = 1
                AdoDetalleViaticos.Recordset("NumIncentivo") = NumIncentivo
                AdoDetalleViaticos.Recordset("valor") = MontoViaticos
                AdoDetalleViaticos.Recordset("NumVez") = 1
                AdoDetalleViaticos.Recordset("pagado") = False
            AdoDetalleViaticos.Recordset.Update
          
    
    
    
    End If
 
 
 
 End If
 
        If CodEmpleado = "39813" Then
          CodEmpleado = 39813
        End If
  
 '///////////////////////////////////////////////////////////////////////////////////////////
 '/////////////////BUSCO EL TIPO DE NOMINA PARA CALCULAR HORAS PRODUCIDAS Y SALARIO//////////
 '//////////////////////////////////////////////////////////////////////////////////////////
       
     Select Case DtaTipoNomina.Recordset("Periodo")
     
            Case "Semanal Sabado"
                    Me.DtaHorasProducidas.RecordSource = "SELECT dbo.DetalleHorasProduccion.CodEmpleado, dbo.DetalleHorasProduccion.NumNomina, dbo.DetalleHorasProduccion.NumLinea, dbo.DetalleHorasProduccion.Lunes + dbo.DetalleHorasProduccion.Martes + dbo.DetalleHorasProduccion.Miercoles + dbo.DetalleHorasProduccion.Jueves + dbo.DetalleHorasProduccion.Viernes AS TotalDias,dbo.Empleado.TarifaHoraria,(dbo.DetalleHorasProduccion.Lunes + dbo.DetalleHorasProduccion.Martes + dbo.DetalleHorasProduccion.Miercoles + dbo.DetalleHorasProduccion.Jueves + dbo.DetalleHorasProduccion.Viernes)* dbo.Empleado.TarifaHoraria AS TotalSalario FROM dbo.DetalleHorasProduccion INNER JOIN dbo.Empleado ON dbo.DetalleHorasProduccion.CodEmpleado = dbo.Empleado.CodEmpleado WHERE (dbo.DetalleHorasProduccion.CodEmpleado = '" & CodEmpleado & "')  AND (dbo.DetalleHorasProduccion.Pagado = 0)AND (DetalleHorasProduccion.NumNomina = " & NumNomina & ")"
                    Me.DtaHorasProducidas.Refresh
            '//////////////////////////////////////////////////////////////////////////////////////
            '//////////////LE AGREGO LAS HORAS PRODUCIDAS///////////////////////////////////////////
            '///////////////////////////////////////////////////////////////////////////////////////
                TarifaHoraria = 0
                TotalHoras = 0
                  If Not DtaHorasProducidas.Recordset.EOF Then
                     TarifaHoraria = Me.DtaHorasProducidas.Recordset("TarifaHoraria")
                     If Not IsNull(Me.DtaHorasProducidas.Recordset("TotalDias")) Then
                       TotalHoras = Me.DtaHorasProducidas.Recordset("TotalDias")
                     End If
                     
                     If TotalHoras > 48 Then
                       TotalHoras = 48
                     End If
                     
                    If TotalHoras < HoraSeptimo Then
                     If Not IsNull(Me.DtaHorasProducidas.Recordset("TotalSalario")) Then
                       TotalSalarioxHora = Me.DtaHorasProducidas.Recordset("TotalSalario")
                     End If
            
                    Else
                    '///////////Calculo el salario sumando 48 Horas + 7mo dia 8 Horas
                       TotalSalarioxHora = TarifaHoraria * TotalHoras
                       Septimo = TarifaHoraria * 8
            
                    End If
                  End If
                  
          Case "Catorcenal los Sabados"
               TotalHoras = 0
               TotalSalarioxHora = 0
               If Me.DtaTipoNomina.Recordset("CalcularHoraTrabajada") Then
                  

                   
                                       Me.DtaHorasProducidas.RecordSource = "SELECT dbo.DetalleHorasProduccion.CodEmpleado, dbo.DetalleHorasProduccion.NumNomina, dbo.DetalleHorasProduccion.NumLinea, dbo.DetalleHorasProduccion.Lunes + dbo.DetalleHorasProduccion.Martes + dbo.DetalleHorasProduccion.Miercoles + dbo.DetalleHorasProduccion.Jueves + dbo.DetalleHorasProduccion.Viernes AS TotalDias,dbo.Empleado.TarifaHoraria,(dbo.DetalleHorasProduccion.Lunes + dbo.DetalleHorasProduccion.Martes + dbo.DetalleHorasProduccion.Miercoles + dbo.DetalleHorasProduccion.Jueves + dbo.DetalleHorasProduccion.Viernes)* dbo.Empleado.TarifaHoraria AS TotalSalario,  DetalleHorasProduccion.Lunes, DetalleHorasProduccion.Martes, DetalleHorasProduccion.Miercoles, " & _
                                                                            "DetalleHorasProduccion.Jueves , DetalleHorasProduccion.Viernes, DetalleHorasProduccion.Sabado, DetalleHorasProduccion.Domingo FROM dbo.DetalleHorasProduccion INNER JOIN dbo.Empleado ON dbo.DetalleHorasProduccion.CodEmpleado = dbo.Empleado.CodEmpleado WHERE (dbo.DetalleHorasProduccion.CodEmpleado = '" & CodEmpleado & "')  AND (dbo.DetalleHorasProduccion.Pagado = 0) AND (DetalleHorasProduccion.NumNomina = " & NumNomina & ") AND (DetalleHorasProduccion.Lunes + DetalleHorasProduccion.Martes + DetalleHorasProduccion.Miercoles + DetalleHorasProduccion.Jueves + DetalleHorasProduccion.Viernes > 0)"
                                       Me.DtaHorasProducidas.Refresh
                               '//////////////////////////////////////////////////////////////////////////////////////
                               '//////////////LE AGREGO LAS HORAS PRODUCIDAS///////////////////////////////////////////
                               '///////////////////////////////////////////////////////////////////////////////////////
                                   TarifaHoraria = 0
                                   TotalHoras = 0
                                   TotalSalarioxHora = 0
                                   Septimo = 0
                                     Do While Not DtaHorasProducidas.Recordset.EOF
                                       
                                       If Me.DtaTipoNomina.Recordset("TipoPago") = "Salario Fijo" Then
                                         TarifaHoraria = Format(DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio / 112, "###,##0.00")
                                       Else
                                         TarifaHoraria = Me.DtaHorasProducidas.Recordset("TarifaHoraria")
                                       End If
                                     
                                        
                                        If Not IsNull(Me.DtaHorasProducidas.Recordset("TotalDias")) Then
                                        
                                           If Me.DtaHorasProducidas.Recordset("TotalDias") > 48 Then
                                             TotalHoras = 48
                                           Else
                                             TotalHoras = TotalHoras + Me.DtaHorasProducidas.Recordset("TotalDias")
                                           End If
                                           
                                           If Me.DtaHorasProducidas.Recordset("TotalDias") < HoraSeptimo Then
                                               If Not IsNull(Me.DtaHorasProducidas.Recordset("TotalSalario")) Then
                                                 TotalSalarioxHora = TarifaHoraria * TotalHoras
                                               End If
                                   
                                           Else
                                               '///////////Calculo el salario sumando 48 Horas + 7mo dia 8 Horas
                                                  TotalSalarioxHora = TarifaHoraria * TotalHoras
                                                  Septimo = Septimo + TarifaHoraria * 8
                                   
                                           End If
                                          
                                        
                                       End If
                                        
                                        
                                      
                                       DtaHorasProducidas.Recordset.MoveNext
                                 Loop
                                 
                         Else
                               TarifaHoraria = Format(DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio / 112, "###,##0.00")
                         End If
          
          Case Else
              TotalSalarioxHora = 0
              If Me.DtaTipoNomina.Recordset("CalcularHoraTrabajada") = 0 Then
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
               Else
                    TarifaHoraria = BuscaTarifa(CodEmpleado, NumNomina) * TasaCambio
                    TotalHoras = BuscaTotalHoras(CodEmpleado, NumNomina)
                    TotalSalarioxHora = TarifaHoraria * TotalHoras
'                    Septimo = BuscaTotalSeptimoSemana(CodEmpleado, NumNomina, HoraSeptimo, 1) + BuscaTotalSeptimoSemana(CodEmpleado, NumNomina, HoraSeptimo, 2)
                    Septimo = BuscaTotalSeptimoSemana(CodEmpleado, NumNomina, HoraSeptimo, 1)
                    Septimo = Septimo * TasaCambio

              End If
          End Select

      
        '/////////////////////////////////////////////////////////////////////////////////
        '/////////////////HAGO EL CALCULO DE LA ANTIGUEDAD DE LOS EMPLEADOS/////////////////
        '//////////////////////////////////////////////////////////////////////////////////

        Antiguedad = 0
        Me.AdoHistorico.RecordSource = "SELECT Codempleado, FechaContrato From Historico WHERE (Codempleado = '" & CodEmpleado & "')"
        Me.AdoHistorico.Refresh
        If Not Me.AdoHistorico.Recordset.EOF Then
         If Not IsNull(Me.AdoHistorico.Recordset("FechaContrato")) Then
         FechaContrato = Me.AdoHistorico.Recordset("FechaContrato")
          NumFecha1 = Format(DtaNomina.Recordset("FechaNomina"), "dd/mm/yyyy")
         'NumFecha1 = Format(Now, "dd/mm/yyyy")
          annos = CDbl(NumFecha1) - CDbl(FechaContrato)
          Anos = Format(annos / 365, "###,##0.00")
          Me.AdoAntiguedad.Refresh
          PAntiguedad = 0
          Do While Not Me.AdoAntiguedad.Recordset.EOF
           If Int(Anos) = Me.AdoAntiguedad.Recordset("años_acum") Then
           
             Select Case DtaTipoNomina.Recordset("Periodo")
             Case "Quincenal"
                  PorcientoAntiguedad = Me.AdoAntiguedad.Recordset("porcent")
                  Antiguedad = (DtaEmpleados.Recordset("SueldoPeriodo") + AumentoBasico + MontoSubsidio) * PorcientoAntiguedad / 100
                   PAntiguedad = Antiguedad
                   Exit Do
             Case "Semanal Viernes"
                  PorcientoAntiguedad = Me.AdoAntiguedad.Recordset("porcent")
                  Antiguedad = (DtaEmpleados.Recordset("SueldoPeriodo") + AumentoBasico + MontoSubsidio) * PorcientoAntiguedad / 100
                  Exit Do
             Case Else
                   PorcientoAntiguedad = Me.AdoAntiguedad.Recordset("porcent")
                   If TarifaHorario <> 0 Then
                      Antiguedad = (HoraBasico * TarifaHoraria) * PorcientoAntiguedad / 100
                   Else
                     Antiguedad = (DtaEmpleados.Recordset("SueldoPeriodo") + AumentoBasico + MontoSubsidio) * PorcientoAntiguedad / 100
                   End If
                   Exit Do
              End Select

            End If
            Me.AdoAntiguedad.Recordset.MoveNext
          Loop
         End If
        End If






        Me.DtaConsulta.RecordSource = "SELECT     sum (CASE WHEN (Fechainicio <= '" & FechaInicioNomina & "') AND (FechaFin < '" & FechaFinNomina & "')  THEN (DATEDIFF(dayofyear, '" & FechaInicioNomina & "', fechafin)) + 1 WHEN (Fechainicio > '" & FechaInicioNomina & "') AND (FechaFin >= '" & FechaFinNomina & "') THEN (DATEDIFF(dayofyear, FechaInicio, '" & FechaFinNomina & "'))  + 1 WHEN Fechainicio > '" & FechaInicioNomina & "' AND FechaFin < '" & FechaFinNomina & "' THEN (DATEDIFF(dayofyear, FechaInicio, FechaFin)) + 1 WHEN Fechainicio <= '" & FechaInicioNomina & "'  AND  FechaFin >= '" & FechaFinNomina & "' THEN (DATEDIFF(dayofyear, '" & FechaInicioNomina & "', '" & FechaFinNomina & "')) + 1 END) AS DiasVacaciones FROM         SolicitudVacaciones "
        Me.DtaConsulta.RecordSource = Me.DtaConsulta.RecordSource & " WHERE     (FechaInicio <= '" & FechaFinNomina & "') AND (FechaFin >= '" & FechaInicioNomina & "') AND CodigoEmpleado = '" & CodempleadoSoli & "' and TipoSolicitud = 'Vacaciones'"
        Me.DtaConsulta.Refresh
        
        Dim ConDiasVacaciones As Double
        ConDiasVacaciones = 0
        'si hay fichas pongo el total en empleados
        If Not Me.DtaConsulta.Recordset.EOF Then
        
        If Not IsNull(DtaConsulta.Recordset("DiasVacaciones")) Then
             ConDiasVacaciones = DtaConsulta.Recordset("DiasVacaciones")
      
        End If
        
        End If





   
        '//////////////////////////////////////////////////////////////
        '///////////////Inicializo Variables///////////////////////////
        '/////////////////////////////////////////////////////////////////
        MontoSubsidio = 0
        DiasSubsidio = 0
        MontoIncentivos = 0
        MontoDeduccion = 0
        MontoPrestamo = 0
        MontoComisiones = 0
        MontoDestajos = 0
        MontoHRSExtras = 0
        MontoDestajos = 0
        Valor = 0
        IncentivoProduccion = 0
        MontoOtrosIngresos = 0
        MontoIncentivoHoras = 0
        MontoTipoVacaciones = 0
        MontoIncentivoExcento = 0
        DiasAdicionales = 0
        MontoDiasAdicionales = 0
        
        '///////////////////////////////////////////////////////////
        '//////////////////agregar incentivos////////////////////////
        '/////////////////////////////////////////////////////////////
        
                    
'        If CodEmpleado = 10219 Then
'          CodEmpleado = 10219
'        End If

        If Not IsNull(DtaEmpleados.Recordset("CantPts")) Then
          DiasAdicionales = DtaEmpleados.Recordset("CantPts")
          MontoDiasAdicionales = (DtaEmpleados.Recordset("SueldoPeriodo") / 15) * DtaEmpleados.Recordset("CantPts")
          MontoHora = BuscaMontoHora(DtaTipoNomina.Recordset("Periodo"), TarifaHoraria, DiasSemana, Sueldo, DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio, DiasMes, CodTipoNomina)
          Respuesta = AgregarDiasAdicionales(DtaEmpleados.Recordset("CantPts"), MontoHora, DiasMes, CodEmpleado, CodTipoNomina, NumNomina)
        End If
        
        
        Me.DtaConsulta.RecordSource = "SELECT     sum (Diasdisfrutar) AS DiasDeducir FROM  SolicitudVacaciones "
        Me.DtaConsulta.RecordSource = Me.DtaConsulta.RecordSource & " WHERE     (FechaInicio <= '" & FechaFinNomina & "') AND (FechaFin >= '" & FechaInicioNomina & "') AND CodigoEmpleado = '" & CodempleadoSoli & "' and TipoSolicitud = 'Ausente'"
        Me.DtaConsulta.Refresh
        
        Dim ConDiasDescuento As Double
        ConDiasDescuento = 0
        'si hay fichas pongo el total en empleados
        If Not Me.DtaConsulta.Recordset.EOF Then

        
                If Not IsNull(DtaConsulta.Recordset("DiasDeducir")) Then
                      DiasBasico = DtaConsulta.Recordset("DiasDeducir")
'                     ConDiasDescuento = DtaConsulta.Recordset("DiasDeducir")
                         Me.DtaConsulta.RecordSource = "Select DiasBasico from Empleado where CodEmpleado = '" & CodEmpleado & "'"
                         Me.DtaConsulta.Refresh
                         
                         If Not Me.DtaConsulta.Recordset.EOF Then
                                 DtaConsulta.Recordset("DiasBasico") = DiasBasico
                                 DtaConsulta.Recordset.Update
                         End If
                         'DtaEmpleados.Recordset("DiasBasico") = ConDiasDescuento
                         
                 Else
                           Me.DtaConsulta.RecordSource = "Select DiasBasico from Empleado where CodEmpleado = '" & CodEmpleado & "'"
                         Me.DtaConsulta.Refresh
                         
                         If Not Me.DtaConsulta.Recordset.EOF Then
                                 DtaConsulta.Recordset("DiasBasico") = 0
                                 DtaConsulta.Recordset.Update
                         End If
                 End If
        
        End If
        
        
        
          '////////////////////// BUSCO SI EL EMPLEADO TIENE ACTUALMENTE SUBSIDIOS ////////////////////
        Me.DtaAuxiliar.RecordSource = "SELECT     SUM(DiasDisfrutar) AS DiasDeducir, FechaInicio AS InicioSolicitud, FechaFin AS FinSolicitud  FROM         SolicitudVacaciones WHERE    "
        Me.DtaAuxiliar.RecordSource = Me.DtaAuxiliar.RecordSource & "(FechaInicio <= '" & FechaFinNomina & " 23:59:59') AND (FechaFin >= '" & FechaInicioNomina & " 00:00:01') AND (TipoSolicitud = 'Subsidio') AND (CodigoEmpleado = '" & CodempleadoSoli & "')   GROUP BY TipoSolicitud, FechaInicio, FechaFin"
        Me.DtaAuxiliar.Refresh
        Dim pDiasSubsidio As Double, InicioSolicitud As Date, FinSolicitud As Date
        pDiasSubsidio = 0
        If Not (DtaAuxiliar.Recordset.EOF) Then
        
             DiasSubsidio = DtaAuxiliar.Recordset("DiasDeducir")
             InicioSolicitud = DtaAuxiliar.Recordset("InicioSolicitud")
             FinSolicitud = DtaAuxiliar.Recordset("FinSolicitud")
             
                 Me.DtaAuxiliar.RecordSource = "Select DiasBasico from Empleado where CodEmpleado = '" & CodEmpleado & "'"
                 Me.DtaAuxiliar.Refresh
                 
                
                
                 '////// Sumo al salario basico ////////////////
                 Dim DiasAPagar As Double
                 DiasAPagar = 0
                If InicioSolicitud <= dFechaInicioNomina And FinSolicitud >= dFechaFinNomina Then ' IN - FN
                     DiasAPagar = DateDiff("d", dFechaInicioNomina, dFechaFinNomina) + 1
                End If
                
                If InicioSolicitud >= dFechaInicioNomina And FinSolicitud >= dFechaFinNomina Then ' IS - FN
                     DiasAPagar = DateDiff("d", InicioSolicitud, dFechaFinNomina) + 1
                End If
                
                If InicioSolicitud >= dFechaInicioNomina And FinSolicitud <= dFechaFinNomina Then ' IS - FS
                      DiasAPagar = DateDiff("d", InicioSolicitud, FinSolicitud) + 1
                End If
                
                If InicioSolicitud <= dFechaInicioNomina And FinSolicitud <= dFechaFinNomina Then ' IN - FS
                     DiasAPagar = DateDiff("d", dFechaInicioNomina, FinSolicitud) + 1
                End If
                
                Dim MontoDiaSub As Double
                MontoDiaSub = BuscaMontoDia(DtaTipoNomina.Recordset("Periodo"), TarifaHoraria, DiasSemana, Sueldo, DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio, DiasMes, CodTipoNomina)
                pDiasSubsidio = DiasAPagar
                MontoSubsidio = (DiasAPagar * MontoDiaSub) * 0.4
                DiasBasico = ConDiasDescuento + DiasAPagar
                
                
                
                If Not Me.DtaAuxiliar.Recordset.EOF Then
                            DtaAuxiliar.Recordset("DiasBasico") = DiasBasico
                            DtaAuxiliar.Recordset.Update
                End If
                
                 
         Else
                   Me.DtaAuxiliar.RecordSource = "Select DiasBasico from Empleado where CodEmpleado = '" & CodEmpleado & "'"
                   Me.DtaAuxiliar.Refresh
                 
                    If Not Me.DtaAuxiliar.Recordset.EOF Then
                            DtaAuxiliar.Recordset("DiasBasico") = DiasBasico
                            DtaAuxiliar.Recordset.Update
                    End If
         End If
        
        
        If (DiasBasico) > 0 Then
          MontoDia = BuscaMontoDia(DtaTipoNomina.Recordset("Periodo"), TarifaHoraria, DiasSemana, Sueldo, DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio, DiasMes, CodTipoNomina)
'          DiasBasico = ConDiasDescuento
          MontoBasico = DiasBasico * MontoDia
          If Me.DtaTipoNomina.Recordset("TipoPago") = "Salario Destajo" Then
            TotalHoras = TotalHoras - (DiasBasico * 8)
          Else
            TotalHoras = TotalHoras - DiasBasico
          End If
        Else
          DiasBasico = 0
          MontoBasico = 0
        End If
        
        TipoVacaciones = False

       If CodEmpleado = "38451" Then
         CodEmpleado = "38451"
        End If

         
        '///////////////////////////////////////INCENTIVOS SIN N DE INFINITO /////////////////////////////////////////////
        DtaIncentivos.Refresh
        If Not DtaIncentivos.Recordset.EOF Then
            TipoVacaciones = False
            NumIncentivo = DtaIncentivos.Recordset("Numincentivo")
            If DtaIncentivos.Recordset("CodTipoIncentivo") = "14" Or DtaIncentivos.Recordset("CodTipoIncentivo") = "15" Or DtaIncentivos.Recordset("CodTipoIncentivo") = "16" Or DtaIncentivos.Recordset("CodTipoIncentivo") = "17" Or DtaIncentivos.Recordset("CodTipoIncentivo") = "18" Then
              MontoIncentivoExcento = DtaIncentivos.Recordset("valor")
            ElseIf DtaIncentivos.Recordset("Incentivo") = "Vacaciones" Then
             MontoTipoVacaciones = DtaIncentivos.Recordset("valor")
            Else
              MontoIncentivos = DtaIncentivos.Recordset("valor")
            End If
            DtaIncentivos.Recordset("NumNomina") = NumNomina
            DtaIncentivos.Recordset.Update
        End If
        
        '///////////////////////////////////////INCENTIVOS CON N DE INFINITO /////////////////////////////////////////////
        Me.DtaConsulta.RecordSource = "SELECT  DetalleIncentivo.NumIncentivo, DetalleIncentivo.Valor, DetalleIncentivo.NumVez, DetalleIncentivo.Pagado, DetalleIncentivo.NumNomina, Incentivo.CodEmpleado, TipoIncentivo.Incentivo , TipoIncentivo.CodTipoIncentivo FROM Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo INNER JOIN TipoIncentivo ON Incentivo.CodTipoIncentivo = TipoIncentivo.CodTipoIncentivo WHERE  (DetalleIncentivo.Pagado = 0) AND (Incentivo.CodEmpleado = '" & CodEmpleado & "') AND (DetalleIncentivo.NumVez = 'n')"
        Me.DtaConsulta.Refresh
        If Not DtaConsulta.Recordset.EOF Then
            TipoVacaciones = False
            NumIncentivo = DtaConsulta.Recordset("Numincentivo")
            If DtaConsulta.Recordset("CodTipoIncentivo") = "14" Or DtaConsulta.Recordset("CodTipoIncentivo") = "15" Or DtaConsulta.Recordset("CodTipoIncentivo") = "16" Or DtaConsulta.Recordset("CodTipoIncentivo") = "17" Or DtaConsulta.Recordset("CodTipoIncentivo") = "18" Then
              MontoIncentivoExcento = MontoIncentivoExcento + DtaConsulta.Recordset("valor")
            ElseIf DtaConsulta.Recordset("Incentivo") = "Vacaciones" Then
             MontoTipoVacaciones = MontoTipoVacaciones + DtaConsulta.Recordset("valor")
            Else
              MontoIncentivos = MontoIncentivos + DtaConsulta.Recordset("valor")
            End If
            DtaConsulta.Recordset("NumNomina") = NumNomina
            DtaConsulta.Recordset.Update
        End If
        
        
         '//////////Al Total Devengado le Sumo las Vacaciones si tiene/////////////
         DtaDetalleNomina.RecordSource = "SELECT  NomVaca.NumNomVaca, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2,DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones,DetalleNomVaca.SalarioMensual * (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento) / '30' - DetalleNomVaca.AdelantoVacaciones AS MontoAPagar, NomVaca.CodTipoNomina , NomVaca.FechaAplica, NomVaca.Transfereir, Empleado.CodEmpleado FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca  " & _
                                         "WHERE (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (NomVaca.FechaAplica BETWEEN CONVERT(DATETIME, '" & Format(FechaInicio, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) AND (NomVaca.Transfereir = 1) AND (Empleado.CodEmpleado = " & CodEmpleado & ") ORDER BY Empleado.CodEmpleado1"
'         DtaDetalleNomina.RecordSource = "SELECT NomVaca.NumNomVaca, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, DetalleNomVaca.SalarioMensual * (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento) / '30' - DetalleNomVaca.AdelantoVacaciones AS MontoAPagar, NomVaca.CodTipoNomina , NomVaca.FechaAplica, NomVaca.Transfereir, Empleado.CodEmpleado FROM NomVaca INNER JOIN Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca  " & _
'                                    "WHERE (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (NomVaca.FechaAplica BETWEEN CONVERT(DATETIME, '" & Format(FechaInicio, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "',, 102)) AND (NomVaca.Transfereir = 1) AND (Empleado.CodEmpleado = " & CodEmpleado & ") ORDER BY Empleado.CodEmpleado1"
         DtaDetalleNomina.Refresh
         If Me.DtaDetalleNomina.Recordset.EOF Then
          MontoVacaciones = 0
          DiasVacaciones = 0
         Else
          If Not IsNull(DtaDetalleNomina.Recordset("MontoAPagar")) Then
           MontoTipoVacaciones = MontoTipoVacaciones + val(DtaDetalleNomina.Recordset("MontoAPagar"))
           DiasVacaciones = val(DtaDetalleNomina.Recordset("DiasAPagar")) - val(DtaDetalleNomina.Recordset("DiasDescuento"))
          End If
         End If
        
     
        
        
'        Do While Not DtaIncentivos.Recordset.EOF
'        If NumIncentivo <> DtaIncentivos.Recordset("Numincentivo") Then
'              NumIncentivo = DtaIncentivos.Recordset("Numincentivo")
'              If DtaIncentivos.Recordset("CodTipoIncentivo") = "14" Or DtaIncentivos.Recordset("CodTipoIncentivo") = "15" Or DtaIncentivos.Recordset("CodTipoIncentivo") = "16" Or DtaIncentivos.Recordset("CodTipoIncentivo") = "17" Or DtaIncentivos.Recordset("CodTipoIncentivo") = "18" Then
'                MontoIncentivoExcento = MontoIncentivoExcento + DtaIncentivos.Recordset("valor")
'              ElseIf DtaIncentivos.Recordset("Incentivo") = "Vacaciones" Then
'                MontoTipoVacaciones = DtaIncentivos.Recordset("valor")
'              Else
'                MontoIncentivos = MontoIncentivos + DtaIncentivos.Recordset("valor")
'              End If
'              DtaIncentivos.Recordset("NumNomina") = NumNomina
'              DtaIncentivos.Recordset.Update
'        End If
'
'
'
'        DtaIncentivos.Recordset.MoveNext
'        Loop
        
'        MontoIncentivos = MontoIncentivos
        
        '////////////////////////////////////////////////////////////////////
        '/////////////////agregar deducciones/////////////////////////////////
        '////////////////////////////////////////////////////////////////////

        
        DtaDeducciones.Refresh
        If Not DtaDeducciones.Recordset.EOF Then
            Numdeduccion = DtaDeducciones.Recordset("Numdeduccion")
            MontoDeduccion = DtaDeducciones.Recordset("valor")
            DtaDeducciones.Recordset("NumNomina") = NumNomina
            DtaDeducciones.Recordset.Update
        End If
        
        Adelantos = 0
        Do While Not DtaDeducciones.Recordset.EOF
        If Numdeduccion <> DtaDeducciones.Recordset("Numdeduccion") Then
           Numdeduccion = DtaDeducciones.Recordset("Numdeduccion")
           MontoDeduccion = MontoDeduccion + DtaDeducciones.Recordset("valor")
            DtaDeducciones.Recordset("NumNomina") = NumNomina
            DtaDeducciones.Recordset.Update
        End If
        If DtaDeducciones.Recordset("codtipodeduccion") = "02" Then
            Adelantos = Adelantos + DtaDeducciones.Recordset("valor")
        End If
        DtaDeducciones.Recordset.MoveNext
        Loop
        
        '/////////////////////////////////////////////////////
        '/////////////agregar prestamo///////////////////////////////////
        '/////////////////////////////////////////////////////////////////

        
        If Not DtaPrestamo.Recordset.EOF Then
            If DtaPrestamo.Recordset("Moneda") = "CS" Then
              If Me.DtaTipoNomina.Recordset("Moneda") = "CS" Then
                 MontoPrestamo = DtaPrestamo.Recordset("CuotaIgual")
              Else
                  TasaCambio = BuscaTasaCambio(FechaNomina)
                  MontoPrestamo = DtaPrestamo.Recordset("CuotaIgual") / TasaCambio
              End If
              
            Else  '////////////SI EL PRESTAMO ES EN DOLARES //////////////////////////
                If Me.DtaTipoNomina.Recordset("Moneda") = "CS" Then
                  TasaCambio = BuscaTasaCambio(FechaNomina)
                  MontoPrestamo = DtaPrestamo.Recordset("CuotaIgual") * TasaCambio
                Else
                   MontoPrestamo = DtaPrestamo.Recordset("CuotaIgual")
                End If
            End If
                'DtaPrestamo.Recordset.Edit
                DtaPrestamo.Recordset("NumNomina") = NumNomina
                DtaPrestamo.Recordset.Update
        End If
        
 
        '////////////////////////////////////////////////////////
        '/////////////////agregar destajos///////////////////////
        '////////////////////////////////////////////////////////
      
       If Not DtaDestajo.Recordset.EOF Then
          MontoDestajos = Me.DtaDestajo.Recordset("Total")
       Else
          '///////////////////////////////////////////////////////////////////////////////
          '/////////////////Agrego la Produccion Manual///////////////////////////////////
          '///////////////////////////////////////////////////////////////////////////////
          Me.AdoDetalleProduccionManual.RecordSource = "SELECT  * From DetalleProduccionManual  " & _
                                                       "WHERE (Pagado = 0) AND (CodEmpleado = " & CodEmpleado & ") AND (FechaProduccion BETWEEN CONVERT(DATETIME, '" & FechaInicio & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102))"
          Me.AdoDetalleProduccionManual.Refresh
          Do While Not Me.AdoDetalleProduccionManual.Recordset.EOF
       
            MontoDestajos = Me.AdoDetalleProduccionManual.Recordset("MontoProduccion") + MontoDestajos
            Me.AdoDetalleProduccionManual.Recordset("NumNomina") = NumNomina
            Me.AdoDetalleProduccionManual.Recordset.Update
            Me.AdoDetalleProduccionManual.Recordset.MoveNext
          Loop
       End If
       
       MontoDestajos = MontoDestajos + MontoSalarioPorciento(CodEmpleado)

        
    '///////////////////////////////////////////////////////////////////////////////////////
    '////////////////Agrego los Destajos de la Tabla Produccion/////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////


    
        Me.DtaDestajo.RecordSource = "SELECT DetalleProduccion.CodEmpleado, DetalleProduccion.NumNomina, DetalleProduccion.CodReferencia, DetalleProduccion.CodProceso,DetalleProduccion.Ref, DetalleProduccion.Lunes, DetalleProduccion.Martes, DetalleProduccion.Miercoles, DetalleProduccion.Jueves, DetalleProduccion.Viernes, DetalleProduccion.Sabado, DetalleProduccion.Domingo, DetalleProduccion.TotalUnidades, DetalleProduccion.SalarioPieza,DetalleProduccion.Precio , DetalleProduccion.unidad, DetalleProduccion.Pagado, Nomina.Activa FROM DetalleProduccion INNER JOIN Nomina ON DetalleProduccion.NumNomina = Nomina.NumNomina Where (DetalleProduccion.CodEmpleado = '" & CodEmpleado & "') And (DetalleProduccion.NumNomina = " & NumNomina & ") And (Nomina.Activa = 1)"
        Me.DtaDestajo.Refresh
        
        Do While Not DtaDestajo.Recordset.EOF
         MontoDestajos = Me.DtaDestajo.Recordset("SalarioPieza") + MontoDestajos
         CodProceso = Me.DtaDestajo.Recordset("CodProceso")
         CodReferencia = Me.DtaDestajo.Recordset("CodReferencia")
         UnidadesProducidas = Me.DtaDestajo.Recordset("TotalUnidades")
         Me.AdoIncentivoPro.RecordSource = "SELECT CodProceso, CodReferencia, Meta, Rango1, Rango2, Rango3, Rango4, Monto1, Monto2, Monto3, Monto4 From MetaIncentivo WHERE (CodProceso = '" & CodProceso & "') AND (CodReferencia = '" & CodReferencia & "')"
         Me.AdoIncentivoPro.Refresh
         If Not Me.AdoIncentivoPro.Recordset.EOF Then
         
         '///////////////////////////////////////////////////////////
         '///////EN ESTE PROCESO BUSCO SI EXITE INCENTIVOS POR METAS////
         '//////////////////////////////////////////////////////////////
                    Rango = (UnidadesProducidas / Me.AdoIncentivoPro.Recordset("Meta")) * 100
           If Rango >= Me.AdoIncentivoPro.Recordset("Rango1") Then
               MontoOtrosIngresos = Me.AdoIncentivoPro.Recordset("Monto1") + MontoOtrosIngresos
           ElseIf Rango >= Me.AdoIncentivoPro.Recordset("Rango2") And Rango < Me.AdoIncentivoPro.Recordset("Rango1") Then
              MontoOtrosIngresos = Me.AdoIncentivoPro.Recordset("Monto2") + MontoOtrosIngresos
           ElseIf Rango >= Me.AdoIncentivoPro.Recordset("Rango3") And Rango < Me.AdoIncentivoPro.Recordset("Rango2") Then
              MontoOtrosIngresos = Me.AdoIncentivoPro.Recordset("Monto3") + MontoOtrosIngresos
           ElseIf Rango >= Me.AdoIncentivoPro.Recordset("Rango4") And Rango < Me.AdoIncentivoPro.Recordset("Rango3") Then
              MontoOtrosIngresos = Me.AdoIncentivoPro.Recordset("Monto4") + MontoOtrosIngresos
           End If
         End If

         Me.DtaDestajo.Recordset.MoveNext
        Loop
        
        Select Case DtaTipoNomina.Recordset("Periodo")
        Case "Semanal Viernes"
            'Busco la cantidad de sabados basado en el calendario en fecha planilla
            CantSabados = SemanasPeriodos(DtaNomina.Recordset("Ano"), DtaNomina.Recordset("Mes"), CodTipoNomina)
            
        Case "Semanal Sabado"
            'busco la cantidad de sabados basados en el calendario en fecha planilla
'            CantSabados = SemanasPeriodos(Year(DtaNomina.Recordset("FechaNomina")), DtaNomina.Recordset("Mes"), CodTipoNomina)
             CantSabados = SemanasPeriodos(DtaNomina.Recordset("Ano"), DtaNomina.Recordset("Mes"), CodTipoNomina)
        Case Else
             CantSabados = SabadosMes(NumFecha2) 'para saber cuantas semanas tiene el mes
        End Select

  '/////////////////////////////////////////////////////////////////////////
  '//////////////////CALCULO EL INCENTIVO X PUNTUALIDAD//////////////////////
'  '///////////////////////////////////////////////////////////////////////////
'  If CodempleadoSoli = "S118010060" Then
'    CodempleadoSoli = "S118010060"
'  End If
  
  
  TotalPuntualidad = 0
  If Me.DtaTipoNomina.Recordset("CalcularHoraTrabajada") Then
        Me.DtaHorasProducidas.Refresh
       Do While Not Me.DtaHorasProducidas.Recordset.EOF
       
           '////////////////////////////////////////////////////////////////////////////////////////////////
           '//////////////////////////////////////INCENTIVO DE PUNTUALIDAD X DIA ///////////////////////////
           '////////////////////////////////////////////////////////////////////////////////////////////////
           If Me.DtaHorasProducidas.Recordset("Lunes") > HoraPuntualidad Then: TotalPuntualidad = TotalPuntualidad + MontoPuntualidad
           If Me.DtaHorasProducidas.Recordset("Martes") > HoraPuntualidad Then: TotalPuntualidad = TotalPuntualidad + MontoPuntualidad
           If Me.DtaHorasProducidas.Recordset("Miercoles") > HoraPuntualidad Then: TotalPuntualidad = TotalPuntualidad + MontoPuntualidad
           If Me.DtaHorasProducidas.Recordset("Jueves") > HoraPuntualidad Then: TotalPuntualidad = TotalPuntualidad + MontoPuntualidad
           If Me.DtaHorasProducidas.Recordset("Viernes") > HoraPuntualidad Then: TotalPuntualidad = TotalPuntualidad + MontoPuntualidad
           If Me.DtaHorasProducidas.Recordset("Sabado") > HoraPuntualidad Then: TotalPuntualidad = TotalPuntualidad + MontoPuntualidad
           If Me.DtaHorasProducidas.Recordset("Domingo") > HoraPuntualidad Then: TotalPuntualidad = TotalPuntualidad + MontoPuntualidad
           
           
           '/////////////////////////////////////////////////////////////////////////////////////////////////
           '///////////////////////////////////INCENTIVO DE VIATICOS ////////////////////////////////////////
           '/////////////////////////////////////////////////////////////////////////////////////////////////
           If Me.DtaHorasProducidas.Recordset("Lunes") > HoraPuntualidad Then: TotalPuntualidad = TotalPuntualidad + ValorViaticoxDia
           If Me.DtaHorasProducidas.Recordset("Martes") > HoraPuntualidad Then: TotalPuntualidad = TotalPuntualidad + ValorViaticoxDia
           If Me.DtaHorasProducidas.Recordset("Miercoles") > HoraPuntualidad Then: TotalPuntualidad = TotalPuntualidad + ValorViaticoxDia
           If Me.DtaHorasProducidas.Recordset("Jueves") > HoraPuntualidad Then: TotalPuntualidad = TotalPuntualidad + ValorViaticoxDia
           If Me.DtaHorasProducidas.Recordset("Viernes") > HoraPuntualidad Then: TotalPuntualidad = TotalPuntualidad + ValorViaticoxDia
           If Me.DtaHorasProducidas.Recordset("Sabado") > HoraPuntualidad Then: TotalPuntualidad = TotalPuntualidad + ValorViaticoxDia
           If Me.DtaHorasProducidas.Recordset("Domingo") > HoraPuntualidad Then: TotalPuntualidad = TotalPuntualidad + ValorViaticoxDia


           
        
        
                    If MontoDestajos = 0 And TotalSalarioxHora = 0 Then
                         TotalPuntualidad = 0
'                    Else
'                         If TotalHoras >= HoraPuntualidad Then
'                           TotalPuntualidad = MontoPuntualidad
'                         Else
'                           TotalPuntualidad = 0
'                         End If
                     End If
        
          Me.DtaHorasProducidas.Recordset.MoveNext
        Loop
    
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////RESTAR DIAS VIATICOS /////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        If TotalPuntualidad > 0 Then
          Dim RestarDias As Double
          RestarDias = RestarDiasViaticos(CDate(FechaInicio), CDate(FechaFin), CodempleadoSoli) * ValorViaticoxDia
          If TotalPuntualidad >= RestarDias Then
           TotalPuntualidad = TotalPuntualidad - RestarDias
           Else
             TotalPuntualidad = 0
           End If
        End If
        

 
  Else
  
        If MontoDestajos = 0 And TotalSalarioxHora = 0 Then
         TotalPuntualidad = 0
        Else
         If TotalHoras >= HoraPuntualidad Then
           TotalPuntualidad = MontoPuntualidad
         Else
           TotalPuntualidad = 0
         End If
        End If
        
  End If
  
       '///////////////////////////////////////////////////////////////////////////////////////////////////
       '///////////////////////////////////////MONTO VIATICOS /////////////////////////
       '///////////////////////////////////////////////////////////////////////////////
       If Not IsNull(DtaEmpleados.Recordset("MontoViatico")) Then
         MontoViaticoEmpleado = DtaEmpleados.Recordset("MontoViatico")
       Else
         MontoViaticoEmpleado = 0
       End If
       
       '////////////////////////SUMO LOS VIATICOS ////////////////
       MontoViaticos = MontoViaticos + MontoViaticoEmpleado


       '////////////////////////////////////////////////////////////////////////////////////
       '//////////////////////////////agregar comisiones////////////////////////////////////
       '/////////////////////////////////////////////////////////////////////////////////////
        MontoComisiones = DtaEmpleados.Recordset("PorcentajeComision")
      
        MontoOtrosIngresos = DtaEmpleados.Recordset("OtrosIngresos") + MontoOtrosIngresos
        If Not IsNull(DtaEmpleados.Recordset("DescripOtrIngre")) Then
          DescripOtrIngre = DtaEmpleados.Recordset("DescripOtrIngre")
        End If
      
       'total devengado del período sin horas extras
'        If DtaEmpleados.Recordset("SueldoPeriodo") = 0 Then
       If CalcularPuntos = True Then
           Select Case DtaTipoNomina.Recordset("Periodo")
             Case "Quincenal"
                Salario = (SalarioMinimo * (1 + (SalarioPorciento / 100)) + ValorPunto * TotalPuntos) / 2
                TotalDevengado = (SalarioMinimo * (1 + (SalarioPorciento / 100)) + ValorPunto * TotalPuntos) / 2 + MontoComisiones + MontoOtrosIngresos - MontoBasico '+ TotalSalarioxHora + Septimo + MontoDestajos
           End Select
        ElseIf Me.DtaTipoNomina.Recordset("CalcularHoraTrabajada") Then
            TotalDevengado = MontoComisiones + MontoOtrosIngresos + TotalSalarioxHora - MontoBasico    '+ TotalSalarioxHora + Septimo + MontoDestajos
            Salario = TotalSalarioxHora - MontoBasico
        Else
            TotalDevengado = DtaEmpleados.Recordset("SueldoPeriodo") + MontoComisiones + MontoOtrosIngresos + TotalSalarioxHora - MontoBasico   '+ TotalSalarioxHora + Septimo + MontoDestajos
            Salario = DtaEmpleados.Recordset("SueldoPeriodo") + TotalSalarioxHora - MontoBasico
        End If

        '///////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////BUSCO SI EL EMPLEADO ESTA DOLARIZADO /////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////
        If Dolarizado = True Then
          TotalDevengado = DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio
          Salario = DtaEmpleados.Recordset("SueldoPeriodo") * TasaCambio + TotalSalarioxHora
        End If

   
'       Select Case DtaTipoNomina.Recordset("Periodo")
'
'        Case "Semanal Viernes"
'            MontoHora = Format(TarifaHoraria, "###,##0.00")
'        Case "Semanal Sabado"
'            MontoHora = Format(TarifaHoraria / (DiasSemana * 8), "###,##0.00")
'        Case "Catorcenal los Viernes"
'            MontoHora = Format(DtaEmpleados.Recordset("SueldoPeriodo") / 112, "###,##0.00")
'        Case "Catorcenal los Sabados"
'            MontoHora = Format(DtaEmpleados.Recordset("SueldoPeriodo") / 112, "###,##0.00")
'        Case "Quincenal"
'            MontoHora = Format(Salario / ((DiasMes * 8) / 2), "###,##0.000000")
''            MontoHora = Format(DtaEmpleados.Recordset("SueldoPeriodo") / ((DiasMes * 8) / 2), "###,##0.000000")
'        Case "Mensual"
'            MontoHora = Format(DtaEmpleados.Recordset("SueldoPeriodo") / (DiasMes * 8), "###,##0.00")
'        Case "Trimestral"
'            MontoHora = Format(DtaEmpleados.Recordset("SueldoPeriodo") / (DiasMes * 8 * 3), "###,##0.00")
'        Case "Semestral"
'            MontoHora = Format(DtaEmpleados.Recordset("SueldoPeriodo") / (DiasMes * 8 * 6), "###,##0.00")
'        End Select
        
        '///////////////////////////////////////////////////////////////////////////
        '/////////////////agregar horas extras///////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////
        
               
       
'       MontoHora = BuscaMontoHora(DtaTipoNomina.Recordset("Periodo"), TarifaHoraria, DiasSemana, Salario + Septimo, Salario + Septimo, DiasMes, CodTipoNomina)
       
'        SqlHrsExtras = "SELECT CodEmpleado, NumNomina, CantHoras, Pagada From HorasExtras WHERE (CodEmpleado = '" & CodEmpleado & "') AND (Pagada = 0) AND (NumNomina = " & NumNomina & ")"
'        DtaHrsExtras.RecordSource = SqlHrsExtras
'        DtaHrsExtras.Refresh
'
'        MontoHRSExtras = 0
'        HE = 0
'        If Not DtaHrsExtras.Recordset.EOF Then
'           If Not IsNull(DtaHrsExtras.Recordset("canthoras")) Then
'                MontoHRSExtras = DtaHrsExtras.Recordset("canthoras") * MontoHora * 2
'                HE = DtaHrsExtras.Recordset("canthoras")
'            End If
'        End If
'
        '///////////////////////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////CALCULO DE HORAS TURNO /////////////////////////////////////////////
        '/////////////////////////////////////////////////////////////////////////////////////////////////
      
'        If CalcularHorasTurno = True Then
'            SqlHrsExtras = "SELECT  * FROM  HorasTurno WHERE (CodEmpleado = '" & CodEmpleado & "') AND (Pagada = 0) AND (NumNomina = " & NumNomina & ")"
'            DtaHrsExtras.RecordSource = SqlHrsExtras
'            DtaHrsExtras.Refresh
'
'            MontoHorasTurno = 0
'            HT = 0
'            If Not DtaHrsExtras.Recordset.EOF Then
'               If Not IsNull(DtaHrsExtras.Recordset("canthoras")) Then
'                    HT = Int(DtaHrsExtras.Recordset("canthoras") / 10)
'                    If HT > 4 Then
'                      HT = 4
'                    End If
'                    HT = HT * 1.5
'                    MontoHorasTurno = HT * MontoHora * 2
'               End If
'            End If
'        End If
        

        
       '/////////////////////////////////////////////////////////////////////////////////////
       '///////EN ESTE PROCESO BUSCO SI EXISTE INCENTIVO POR PRODUCCION POR DEPARTAMENTO//////
       '/////////////////////////////////////////////////////////////////////////////////////
       

            IncentivoProduccion = 0
            If Not MontoDestajos = 0 Then
                             
              Me.AdoIncentivoPro.RecordSource = "SELECT NumeroIncentivo, CodDepartamento, CantidadHoras, Valor From IncentivoProduccion WHERE (CodDepartamento = '" & CodDepartamento & "') "
              Me.AdoIncentivoPro.Refresh
              If Not Me.AdoIncentivoPro.Recordset.EOF Then
                       
                 Valor = Me.AdoIncentivoPro.Recordset("Valor")
                                  
                 If TotalHoras >= Me.AdoIncentivoPro.Recordset("CantidadHoras") Then
                    IncentivoProduccion = Valor
                 End If
              End If
            End If

        
            BonoProduccion = 0

            Select Case FormatoNomina
             
               Case "Nomina Produccion"
               
                    '/////////////////////////////////////////////////////////////////////////
                    '////////PARA CASOS EN QUE LA PRODUCCION SE MENOR QUE LAS HORAS TRABAJADAS
                    '/////////////////////////////////////////////////////////////////////////
                    Select Case DtaTipoNomina.Recordset("Periodo")
                      Case "Semanal Viernes"
                       If TotalSalarioxHora > MontoDestajos Then
                          MontoDestajos = 0
                          PagoProduccion = 0
                         If TotalHoras >= HoraSeptimo Then
                            Septimo = TarifaHoraria * 8
                         Else
                         Septimo = 0
                         End If
                       Else
                         TotalSalarioxHora = 0
                         MontoHRSExtras = 0
                         HE = 0
                         PagoProduccion = 1
                         If TotalHoras >= HoraSeptimo Then
                         Septimo = MontoDestajos / 6
                         Else
                         Septimo = 0
                         End If
                      End If
                   End Select

               Case "Nomina Bono Produccion"

                    If MontoDestajos > (TotalSalarioxHora + MontoHRSExtras) Then

                         If TotalHoras >= HoraSeptimo Then
                            Septimo = TarifaHoraria * 8
                            SeptimoAnterior = MontoDestajos / 6
                         Else
                            Septimo = 0
                            SeptimoAnterior = 0
                         End If
                         
                     If Calcular7mo = True Then
                        BonoProduccion = (MontoDestajos + SeptimoAnterior) - TotalSalarioxHora - MontoHRSExtras
                        BonoProduccion = BonoProduccion - Septimo
                     Else
                        BonoProduccion = MontoDestajos - TotalSalarioxHora - MontoHRSExtras
                        BonoProduccion = BonoProduccion - Septimo
                     End If
                     
                    Else
                       
                        BonoProduccion = 0
                        IncentivoProduccion = 0
                    
                         If TotalHoras >= HoraSeptimo Then
'                            Septimo = TotalSalarioxHora / 6
                           Septimo = TarifaHoraria * 8
                         Else
                            Septimo = 0
                         End If
                     End If
                    MontoDestajos = 0
                    TotalDevengado = BonoProduccion + MontoOtrosIngresos '+ Septimo + TotalSalarioxHora
   
           End Select
           
        '---------------------------------------------------------------------------------------
        '------------------------------SUMO LA ANTIGUEDAD AL TOTAL DEVENgaDO --------------------
        '----------------------------------------------------------------------------------------
        TotalDevengado = TotalDevengado + Antiguedad


            
          
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////CALCULO EL INCENTIVO PORCENTUAL////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        
        If Not PorcientoIncentivo = 0 Then
           MontoIncentivoHoras = MontoHRSExtras * (PorcientoIncentivo / 100)
           DescripOtrIngre = PorcientoIncentivo & "% Incentivo"
           MontoOtrosIngresos = MontoOtrosIngresos + MontoIncentivoHoras
        Else
           MontoIncentivoHoras = 0
        End If
        

'         If TipoVacaciones = True Then MontoVacaciones = MontoVacaciones + MontoTipoVacaciones
          '-------------------SUMO AL SALARIO EL PORCIENTO DE SALARIO---------------------------
          
         
         
         '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      
             Select Case DtaTipoNomina.Recordset("Periodo")
                Case "Semanal Viernes"
                     SalarioMensual = (TotalDevengado + MontoVacaciones) * CantSabados
                Case "Semanal Sabado"
                    SalarioMensual = (TotalDevengado + MontoVacaciones) * CantSabados
                    
                Case "Catorcenal los Viernes"
                   If DiaFin < 28 Then
                    If Not DtaDetalleNomina.Recordset.EOF Then
                     SalarioMensual = (((TotalDevengado * 15) / 14) + MontoVacaciones) * 2
                    End If
                   Else
                    If Not DtaDetalleNomina.Recordset.EOF Then
                     SalarioMensual = TotalDevengado + TotalDevengadoAnterior + MontoVacaciones
                    End If
                   End If
                Case "Catorcenal los Sabados"
                 If DiaFin < 28 Then
                    SalarioMensual = (((TotalDevengado * 15) / 14) + MontoVacaciones) * 2
                 Else
                    SalarioMensual = TotalDevengado + TotalDevengadoAnterior + MontoVacaciones
                 End If
                Case "Quincenal"
                    If DiaFin < 28 Then
                       SalarioMensual = (TotalDevengado + MontoVacaciones) * 2
                    Else
                       SalarioMensual = TotalDevengado + TotalDevengadoAnterior + MontoVacaciones
                    End If
                Case "Mensual"
                 
                    SalarioMensual = TotalDevengado + MontoVacaciones
              
'                Case "Trimestral"
'                    SalarioMensual = (val(DtaDetalleNomina.Recordset("VacacionesPagadas")) + val(DtaDetalleNomina.Recordset("AdelantosVacaciones"))) * 2 + TotalDevengado / 3
'                Case "Semestral"
'                    SalarioMensual = (val(DtaDetalleNomina.Recordset("VacacionesPagadas")) + val(DtaDetalleNomina.Recordset("AdelantosVacaciones"))) * 2 + TotalDevengado / 6
                End Select
              'total devengado del período incluyendo horas extras


'   If CodEmpleado = 9626 Then
'     CodEmpleado = 9626
'   End If

       '/////////////////////////////// HORAS EXTRAS /////////////////////////////////////////
       
'       If DtaEmpleados.Recordset("Nombre1") = "Ronaldo" Then
'       DtaEmpleados.Recordset("Nombre1") = "Ronaldo"
'       End If
       
'       MontoHora = BuscaMontoHora(DtaTipoNomina.Recordset("Periodo"), TarifaHoraria, DiasSemana, Sueldo, TotalDevengado + TotalSalarioxHora + IncentivoProduccion + MontoHorasTurno, DiasMes, CodTipoNomina)
        MontoHora = BuscaMontoHora(DtaTipoNomina.Recordset("Periodo"), TarifaHoraria, DiasSemana, Sueldo, ((DtaEmpleados.Recordset("SueldoPeriodo") + AumentoBasico + MontoSubsidio) * TasaCambio) + Antiguedad + MontoSalarioPorciento(CodEmpleado), DiasMes, CodTipoNomina)
       
        SqlHrsExtras = "SELECT CodEmpleado, NumNomina, CantHoras, Pagada From HorasExtras WHERE (CodEmpleado = '" & CodEmpleado & "') AND (Pagada = 0) AND (NumNomina = " & NumNomina & ")"
        DtaHrsExtras.RecordSource = SqlHrsExtras
        DtaHrsExtras.Refresh

        MontoHRSExtras = 0
        HE = 0
        
        
        If EmpleadoConstruccion = True Then
            If Not DtaHrsExtras.Recordset.EOF Then
                If DtaEmpleados.Recordset("SalarioFijo") = "S" Then
                   MontoHRSExtras = DtaHrsExtras.Recordset("canthoras") * MontoHora * DiasHorasExtra
                Else
                   MontoHRSExtras = DtaHrsExtras.Recordset("canthoras") * MontoHora * DiasHorasExtra
                End If
                 HE = DtaHrsExtras.Recordset("canthoras")
            End If
                Septimo = (TarifaHoraria * 8) + MontoHRSExtras
                
               
        Else
                If Not DtaHrsExtras.Recordset.EOF Then
                    If Not IsNull(DtaHrsExtras.Recordset("canthoras")) Then
                        MontoHRSExtras = DtaHrsExtras.Recordset("canthoras") * MontoHora * DiasHorasExtra
                        HE = DtaHrsExtras.Recordset("canthoras")
                    End If
                Else
                    MontoHRSExtras = 0
                    HE = 0
                End If
                
               
        End If
        
       
    
    
        

        
    
    


'*************************************************************************************************
'*************************************************************************************************
'****************************CALCULO DE TODAS LAS DEDUCCIONES*************************************
'**************************************************************************************************
'*************************************************************************************************


        '///////////// BUSCO LOS DIAS A DESCONTAR EN EL LAS FICHAS /////////////////

    
        
        '////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////// Deduccion por dias de descuento /////////////////////////////////
        
        'calculo la deduccion por los dias de descuento
        DeduccionPorFalta = 0
    If CodEmpleado = "10290" Then
        CodEmpleado = "10290"
    End If
       
            If ConDiasDescuento > 0 Then
             DeduccionPorFalta = BuscaDeduccionPorFalta(DtaTipoNomina.Recordset("Periodo"), TotalDevengado - TotalPuntualidad, MontoVacaciones, ConDiasDescuento, DiasMes)
                    
                    '/////////////////////////////////////////////////////////////////////
                    '/////GRABO LA DEDUCCION POR FALTAS/////////////////////////////////
                    '///////////////////////////////////////////////////////////////////
                    
                    Me.DtaConsulta.RecordSource = "SELECT Deduccion.NumDeduccion, Deduccion.CodEmpleado, Deduccion.CodTipoDeduccion, DetalleDeduccion.NumNomina " & _
                                                  "FROM Deduccion INNER JOIN " & _
                                                  "DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion " & _
                                                  "WHERE (Deduccion.CodTipoDeduccion = '01') AND (DetalleDeduccion.NumNomina = " & CDbl(NumNomina) & ") AND (Deduccion.CodEmpleado = " & CDbl(CodEmpleado) & ") "
                    Me.DtaConsulta.Refresh
                    
                    If Me.DtaConsulta.Recordset.EOF Then
                    
                         'creo la nueva deducción
                        DtaConsecutivos.Refresh
                        'DtaConsecutivos.Recordset.Edit
                        DtaConsecutivos.Recordset("Deducciones") = DtaConsecutivos.Recordset("Deducciones") + 1
                        DtaConsecutivos.Recordset.Update
                        
                        Me.DtaDeduccion.Refresh
                        If Me.DtaDeduccion.Recordset.EOF Then
                            NumeroDeduccion = 1
                        Else
                            Me.DtaDeduccion.Recordset.MoveLast
                            NumeroDeduccion = Me.DtaDeduccion.Recordset("NumDeduccion") + 1
                        End If
                    
                        DtaDeduccion.Recordset.AddNew
                        DtaDeduccion.Recordset("NumDeduccion") = NumeroDeduccion
                        DtaDeduccion.Recordset("CodEmpleado") = DtaEmpleados.Recordset("CodEmpleado")
                        DtaDeduccion.Recordset("codtipodeduccion") = "01"
                        DtaDeduccion.Recordset("numveces") = "1"
                        DtaDeduccion.Recordset("pagado") = 0
                        DtaDeduccion.Recordset("NumNomina") = NumNomina
                        DtaDeduccion.Recordset.Update
                    
                        '////////////////////////////////////////////////////////////////////
                        '//GRABO EL DETALLE DE LA DEDUCCION POR FALTAS//////////////////////
                        '///////////////////////////////////////////////////////////////////
                        
                        Me.DtaConsulta.RecordSource = "SELECT id,Deduccion.CodTipoDeduccion, DetalleDeduccion.NumDeduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumVez, DetalleDeduccion.Pagado, DetalleDeduccion.NumNomina, Deduccion.CodEmpleado FROM Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion "
                        Me.DtaConsulta.Refresh
                        
                        If Me.DtaConsulta.Recordset.EOF Then
                          IdDeduccion = 1
                        Else
                          Me.DtaConsulta.Recordset.MoveLast
                          IdDeduccion = Me.DtaConsulta.Recordset("Id") + 2
                        End If
                    
                        DtaDeducciones.Recordset.AddNew
                        'DtaDeducciones.Recordset("Id") = IdDeduccion
                        DtaDeducciones.Recordset("Numdeduccion") = NumeroDeduccion
                        DtaDeducciones.Recordset("valor") = Format(DeduccionPorFalta, "##,##0.00")
                        DtaDeducciones.Recordset("NumVez") = 1
                        DtaDeducciones.Recordset("pagado") = 0
                        DtaDeducciones.Recordset("NumNomina") = NumNomina
                        DtaDeducciones.Recordset.Update
                        
                        MontoDeduccion = MontoDeduccion + DeduccionPorFalta
                    Else
                    
                        Me.DtaConsulta.RecordSource = "SELECT Deduccion.CodEmpleado, Deduccion.CodTipoDeduccion, DetalleDeduccion.NumNomina, DetalleDeduccion.Valor " & _
                                                      "FROM  DetalleDeduccion INNER JOIN Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion " & _
                                                      "WHERE  (Deduccion.CodEmpleado = " & CDbl(CodEmpleado) & ") AND (Deduccion.CodTipoDeduccion = '01') AND (DetalleDeduccion.NumNomina = " & CDbl(NumNomina) & ") "
                       
                        Me.DtaConsulta.Refresh
                        If Not Me.DtaConsulta.Recordset.EOF Then
                           '////////////////////////////////////////////////////////////////////////////////
                           '///////////EDITO LA DEDUCCION SI EXISTE////////////////////////////////////////
                           '///////////////////////////////////////////////////////////////////////////////
                            DtaConsulta.Recordset("valor") = Format(DeduccionPorFalta, "##,##0.00")
                            DtaConsulta.Recordset.Update
                        End If
                    
                    End If

            End If

  
 
' End If 'del if que pregunta si hay dias de descuento



MontoInss = 0
MontoInssPatronal = 0

        If CodEmpleado = "38605" Then
          CodEmpleado = 38605
        End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////SUMO LOS OTROS INGRESOS E INCENTIVOS ///////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    MontoInss = CalcularMontoINSS(CodTipoNomina, TotalDevengado + AumentoBasico + MontoVacaciones + MontoDestajos + Septimo + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoHorasTurno + MontoTipoVacaciones + Reembolso, SalarioMensual)
    MontoInss = MontoInssRegistros(0)
    MontoInssPatronal = MontoInssRegistros(1)
    MontoInssMensual = MontoInssRegistros(2)
    MontoInssPatronalMensual = MontoInssRegistros(3)

''//////////////Verifico si el calculo del Inss es Porcentual//////////////
'CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
'Me.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, PorcientoInss, TasaInssPatronal, TasaInss, PorcientoIr, TasaIr From TipoNomina WHERE (PorcientoInss = 1) AND (CodTipoNomina = '" & CodTipoNomina & "' )"
'Me.DtaConsulta.Refresh
'If DtaConsulta.Recordset.EOF Then
' If DtaEmpleados.Recordset("ExentoInss") = 0 Then
'
'         DtaInss.Refresh
'         Do While Not DtaInss.Recordset.EOF
'                If DtaTipoNomina.Recordset("Periodo") = "Semanal Viernes" Then
'                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
'                      MontoInss = DtaInss.Recordset("montolaboral1")
'                      MontoInssPatronal = DtaInss.Recordset("montopatronal1")
'                      Exit Do
'                   End If
'
'                ElseIf DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
'                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
'                      MontoInss = DtaInss.Recordset("montolaboral1")
'                      MontoInssPatronal = DtaInss.Recordset("montopatronal1")
'                      Exit Do
'                   End If
'                ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Viernes" Then
'
'                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
'                      If CantSabados = 4 Then
'                        If DiaFin < 28 Then
'                        MontoInss = (DtaInss.Recordset("montolaboral4") / 2)
'                        MontoInssPatronal = (DtaInss.Recordset("montopatronal4") / 2)
'                        Exit Do
'                       Else
'                        MontoInssMensual = DtaInss.Recordset("montolaboral4")
'                        MontoInssPatronalMensual = DtaInss.Recordset("montopatronal4")
'                        MontoInss = MontoInssMensual - MontoInssAnterior
'                        MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
'                       End If
'                      Else
'              '/////////Calcula para Cinco Semanas////////
'                     If DiaFin < 28 Then
'                        MontoInss = (DtaInss.Recordset("montolaboral5") / 2)
'                        MontoInssPatronal = (DtaInss.Recordset("montopatronal5") / 2)
'                        Exit Do
'                     Else
'                        MontoInssMensual = DtaInss.Recordset("montolaboral5")
'                        MontoInssPatronalMensual = DtaInss.Recordset("montopatronal5")
'                        MontoInss = MontoInssMensual - MontoInssAnterior
'                        MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
'                     End If
'
'                      End If
'                   End If
'                ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
'
'                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
'                      If CantSabados = 4 Then
'               '/////////////Calcula para cuatro semanas////
'                       If DiaFin < 28 Then
'                        MontoInss = (DtaInss.Recordset("montolaboral4") / 2)
'                        MontoInssPatronal = (DtaInss.Recordset("montopatronal4") / 2)
'                        Exit Do
'                       Else
'                        MontoInssMensual = DtaInss.Recordset("montolaboral4")
'                        MontoInssPatronalMensual = DtaInss.Recordset("montopatronal4")
'                        MontoInss = MontoInssMensual - MontoInssAnterior
'                        MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
'                       End If
'                      Else
'              '/////////Calcula para Cinco Semanas////////
'                     If DiaFin < 28 Then
'                        MontoInss = (DtaInss.Recordset("montolaboral5") / 2)
'                        MontoInssPatronal = (DtaInss.Recordset("montopatronal5") / 2)
'                        Exit Do
'                     Else
'                        MontoInssMensual = DtaInss.Recordset("montolaboral5")
'                        MontoInssPatronalMensual = DtaInss.Recordset("montopatronal5")
'                        MontoInss = MontoInssMensual - MontoInssAnterior
'                        MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
'                     End If
'                      End If
'                   End If
'                ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
'
'                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
'                      If CantSabados = 4 Then
'                       '///////Calculo para 4 Semanas///////////
'                       If DiaFin < 28 Then
'                        MontoInss = (DtaInss.Recordset("montolaboral4") / 2)
'                        MontoInssPatronal = (DtaInss.Recordset("montopatronal4") / 2)
'                        Exit Do
'                       Else
'                         MontoInssMensual = DtaInss.Recordset("montolaboral4")
'                         MontoInssPatronalMensual = DtaInss.Recordset("montopatronal4")
'                         MontoInss = MontoInssMensual - MontoInssAnterior
'                         MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
'                         Exit Do
'                       End If
'                      Else
'                      '///Calculo para 5 Semansas//////////
'                       If DiaFin < 28 Then
'                        MontoInss = (DtaInss.Recordset("montolaboral5") / 2)
'                        MontoInssPatronal = (DtaInss.Recordset("montopatronal5") / 2)
'
'                        Exit Do
'                       Else
'                        MontoInssMensual = DtaInss.Recordset("montolaboral5")
'                        MontoInssPatronalMensual = DtaInss.Recordset("montopatronal5")
'                        MontoInss = MontoInssMensual - MontoInssAnterior
'                        MontoInssPatronal = MontoInssPatronalMensual - MontoInssPatronalAnterior
'                       End If
'                      End If
'                   End If
'
'                ElseIf DtaTipoNomina.Recordset("Periodo") = "Mensual" Then
'
'                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
'                      If CantSabados = 4 Then
'                        MontoInss = DtaInss.Recordset("montolaboral4")
'                        MontoInssPatronal = DtaInss.Recordset("montopatronal4")
'                        Exit Do
'                      Else
'                        MontoInss = DtaInss.Recordset("montolaboral5")
'                        MontoInssPatronal = DtaInss.Recordset("montopatronal5")
'                        Exit Do
'
'                      End If
'                   End If
'
'                ElseIf DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
'
'                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
'                      If CantSabados = 4 Then
'                        MontoInss = DtaInss.Recordset("montolaboral4") * 3
'                        MontoInssPatronal = DtaInss.Recordset("montopatronal4") * 3
'                        Exit Do
'                      Else
'                        MontoInss = DtaInss.Recordset("montolaboral5") * 3
'                        MontoInssPatronal = DtaInss.Recordset("montopatronal5") * 3
'                        Exit Do
'
'                      End If
'                   End If
'
'
'
'                ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
'
'
'                   If DtaInss.Recordset("desde") < (SalarioMensual) And DtaInss.Recordset("Hasta") > (SalarioMensual) Then
'                      If CantSabados = 4 Then
'                        MontoInss = DtaInss.Recordset("montolaboral4") * 6
'                        MontoInssPatronal = DtaInss.Recordset("montopatronal4") * 6
'                        Exit Do
'                      Else
'                        MontoInss = DtaInss.Recordset("montolaboral5") * 6
'                        MontoInssPatronal = DtaInss.Recordset("montopatronal5") * 6
'                        Exit Do
'
'                      End If
'                   End If
'
'                End If
'
'         DtaInss.Recordset.MoveNext
'         Loop
' End If 'del if que pregunta si el empleado ers excento de INSS
'Else
'
'
'
' If DtaEmpleados.Recordset("ExentoInss") = 0 Then
'  TasaInss = Me.DtaConsulta.Recordset("TasaInss")
'  TasaInssPatronal = Me.DtaConsulta.Recordset("TasaInssPatronal")
'
'  Select Case DtaTipoNomina.Recordset("Periodo")
'    Case "Quincenal"
'        If DiaFin < 28 Then
'           If (TotalDevengado + MontoVacaciones + MontoTipoVacaciones) > 54964 Then
'              MontoInss = (54964 * (TasaInss / 100)) / 2
'              MontoInssPatronal = (54964 * (TasaInssPatronal / 100)) / 2
'              MontoInssMensual = MontoInssAnterior + MontoInss
'              MontoInssPatronalMensual = MontoInssPatronalAnterior + MontoInssPatronal
'           Else
'             MontoInss = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoHorasTurno + MontoTipoVacaciones) * (TasaInss / 100)
''             MontoInss = (TotalDevengado + MontoVacaciones) * (TasaInss / 100)
'             MontoInssPatronal = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoHorasTurno + MontoTipoVacaciones) * (TasaInssPatronal / 100)
'             MontoInssMensual = MontoInssAnterior + MontoInss
'             MontoInssPatronalMensual = MontoInssPatronalAnterior + MontoInssPatronal
'           End If
'
'        Else
'           MontoInssMensual = ((TotalDevengado + MontoVacaciones + MontoTipoVacaciones) * (TasaInss / 100)) + MontoInssAnterior
'           If MontoInssMensual > (54964 * (TasaInss / 100)) Then
'              MontoInss = (54964 * (TasaInss / 100)) - MontoInssAnterior
'              MontoInssMensual = MontoInss + MontoInssAnterior
'              MontoInssPatronal = (54964 * (TasaInssPatronal / 100)) - MontoInssPatronalAnterior
'              MontoInssPatronalMensual = MontoInssPatronalAnterior + MontoInssPatronal
'           Else
'             MontoInss = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoHorasTurno + MontoTipoVacaciones) * (TasaInss / 100)
''             MontoInss = (TotalDevengado + MontoVacaciones) * (TasaInss / 100)
'             MontoInssPatronal = (TotalDevengado + MontoVacaciones + MontoTipoVacaciones) * (TasaInssPatronal / 100)
'             MontoInssMensual = MontoInssAnterior + MontoInss
'             MontoInssPatronalMensual = MontoInssPatronalAnterior + MontoInssPatronal
'           End If
'        End If
'    Case "Mensual"
'
'           If (TotalDevengado + MontoVacaciones + MontoTipoVacaciones) > 54964 Then
'              MontoInss = 2344.88
'              MontoInssPatronal = 5627.7
'              MontoInssMensual = MontoInssAnterior + MontoInss
'              MontoInssPatronalMensual = MontoInssPatronalAnterior + MontoInssPatronal
'           Else
'             MontoInss = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoHorasTurno + MontoTipoVacaciones) * (TasaInss / 100)
''             MontoInss = (TotalDevengado + MontoVacaciones) * (TasaInss / 100)
'             MontoInssPatronal = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoHorasTurno + MontoTipoVacaciones) * (TasaInssPatronal / 100)
'             MontoInssMensual = MontoInssAnterior + MontoInss
'             MontoInssPatronalMensual = MontoInssPatronalAnterior + MontoInssPatronal
'           End If
'
'
'
'
'    Case Else
'             '+ MontoOtrosIngresos
'             MontoInss = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) * (TasaInss / 100)
'             MontoInssPatronal = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) * (TasaInssPatronal / 100)
'             MontoInssMensual = MontoInssAnterior + MontoInss
'             MontoInssPatronalMensual = MontoInssPatronalAnterior + MontoInssPatronal
'
'  End Select
'
' Else
'    MontoInss = 0
'    MontoInssPatronal = 0
'    MontoInssMensual = 0
'    MontoInssPatronalMensual = 0
' End If
'
'End If 'del if del porciento


  

        
 '////////////////////////////CALCULO EL INATEC ///////////////////////////////////////////////////////////
        'le agrego al total devengado las horas extras
'        TotalDevengado = TotalDevengado + MontoHRSExtras + MontoIncentivoHoras + IncentivoProduccion + MontoIncentivos
         
        INATEC = (TotalDevengado + AumentoBasico + MontoSubsidio + MontoVacaciones + MontoDestajos + Septimo + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoHorasTurno + MontoTipoVacaciones + Reembolso) * 0.02
        'agregar INSS Laboral y Patronal
        
'////////////////////////CALCULO DEL IR///////////////////////////
MontoIr = 0
MontoIRPatronal = 0

        If CodEmpleado = "59840" Then
          CodEmpleado = 59840
        End If
        
        
MontoExento = 0
MontoExento = MontoIncentivoExcento

MontoIr = CalcularMontoIr(CodTipoNomina, TotalDevengado + AumentoBasico + MontoSubsidio + MontoVacaciones + MontoDestajos + Septimo + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoHorasTurno + MontoTipoVacaciones + Reembolso, CodEmpleado, TipoCalculoIr, IrUltimaSemana, MontoInss, NumNomina)

''''////////////////////////////////////////////////
''''///PRIMERO BUSCO EL NUMERO DEL PERIODO PARA CALCULAR IR
''''////////////////////////////////////////////////////////
''''///////////////////////Verifico si Tiene Ir Porcentual//////////////////////////////
'''CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
'''Me.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, PorcientoInss, TasaInss, PorcientoIr, TasaIr From TipoNomina WHERE (PorcientoIr = 1) AND (CodTipoNomina = '" & CodTipoNomina & "' )"
'''Me.DtaConsulta.Refresh
'''If DtaConsulta.Recordset.EOF Then
'''
''' Select Case DtaTipoNomina.Recordset("Periodo")
'''      Case "Quincenal"
'''        Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (CodTipoNomina = " & CodTipoNomina & ") AND (Inicio = '" & Format(CDate(Me.TxtFechaIni.Text), "DD/MM/YYYY") & "') ORDER BY Periodo"
'''        Me.AdoPeriodoFiscal.Refresh
'''        If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
'''           PeriodoFiscal = Me.AdoPeriodoFiscal.Recordset("Periodo") ' formula = n
'''           NumeroPeriodo = 24 - (PeriodoFiscal - 1) 'formula = 24-(n-1)
'''           AñoFiscal = Me.AdoPeriodoFiscal.Recordset("Año")
'''        End If
'''
'''       Case "Mensual"
'''        Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (CodTipoNomina = " & CodTipoNomina & ") AND (Inicio = '" & Format(CDate(Me.TxtFechaIni.Text), "DD/MM/YYYY") & "') ORDER BY Periodo"
'''        Me.AdoPeriodoFiscal.Refresh
'''        If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
'''           PeriodoFiscal = Me.AdoPeriodoFiscal.Recordset("Periodo") ' formula = n
'''           NumeroPeriodo = 12 - (PeriodoFiscal - 1) 'formula = 12-(n-1)
'''           AñoFiscal = Me.AdoPeriodoFiscal.Recordset("Año")
'''        End If
'''
'''
'''  End Select
'''End If
'''
''''//////////////////////////////////////////////////
''''///BUSCO LA FECHA INICIAL DEL AÑO FISCAL
''''////////////////////////////////////////////////////////
'''        Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (Año = " & AñoFiscal & ") AND (CodTipoNomina = " & CodTipoNomina & ") AND (Periodo = 1)ORDER BY Periodo"
'''        Me.AdoPeriodoFiscal.Refresh
'''        If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
'''          FechaInicialIr = Me.AdoPeriodoFiscal.Recordset("Inicio")
'''        End If
'''
''''//////////////////////////////////////////////////
''''///BUSCO LA FECHA DE LA ULTIMA NOMINA DEL AÑO FISCAL CALCULADA
''''////////////////////////////////////////////////////////
'''        PeriodoFiscal = PeriodoFiscal - 1
'''        Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (Año = " & AñoFiscal & ") AND (CodTipoNomina = " & CodTipoNomina & ") AND (Periodo = " & PeriodoFiscal & ") ORDER BY Periodo"
'''        Me.AdoPeriodoFiscal.Refresh
'''        PeriodoFiscal = PeriodoFiscal + 1
'''        If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
'''          FechaFinalIr = Me.AdoPeriodoFiscal.Recordset("Final")
'''        End If
'''
''''/////////////////////////////////////////////////////////////////
''''///BUSCO LAS NOMINAS ACUMULADAS/////////////////////////////////
''''///////////////////////////////////////////////////////////////////
'''
'''sql = "SELECT     DetalleNomina.CodEmpleado AS CodEmpleado, SUM(DetalleNomina.MontoINSS) AS MontoINSS, SUM(DetalleNomina.MontoIR) AS MontoIR, " & _
'''     "SUM(DetalleNomina.VacacionesPagadas) AS Vacaciones, SUM(DetalleNomina.INSSPatronal) AS INSSPatronal, SUM(DetalleNomina.IRPatronal) AS IRPatronal, " & _
'''     "SUM(DetalleNomina.INATEC) AS INATEC, " & _
'''     "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos " & _
'''     "+ DetalleNomina.Incentivos + DetalleNomina.VacacionesPagadas + DetalleNomina.AdelantosVacaciones) AS TotalDevengado, COUNT(DetalleNomina.NumNomina) AS NQuincenas, MIN(Nomina.FechaNominaINI) AS FechaIngreso FROM DetalleNomina INNER JOIN " & _
'''     "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina " & _
'''     "WHERE     (Nomina.FechaNomina <= CONVERT(DATETIME, '" & Format(FechaFinalIr, "yyyy/mm/dd") & "', 102)) AND (Nomina.FechaNominaINI >= CONVERT(DATETIME, '" & Format(FechaInicialIr, "yyyy/mm/dd") & "', 102)) " & _
'''     "GROUP BY DetalleNomina.CodEmpleado " & _
'''     "Having (DetalleNomina.CodEmpleado = " & CodEmpleado & ") "
'''
'''
'''    Me.DtaConsulta.RecordSource = sql
'''    Me.DtaConsulta.Refresh
'''    TotalDevengadoAcumulado = 0
'''    MontoIrAcumulado = 0
'''    VacacionesPagadas = 0
'''    If Not Me.DtaConsulta.Recordset.EOF Then
'''       MontoIrAcumulado = Me.DtaConsulta.Recordset("MontoIR")
'''       TotalDevengadoAcumulado = Me.DtaConsulta.Recordset("TotalDevengado") - Me.DtaConsulta.Recordset("MontoINSS")
'''       NQuincenas = Me.DtaConsulta.Recordset("NQuincenas") + 1
'''       FechaIngreso = Me.DtaConsulta.Recordset("FechaIngreso")
'''       VacacionesAcumuladas = Me.DtaConsulta.Recordset("Vacaciones")
'''    Else
'''       NQuincenas = 1
'''       FechaIngreso = Me.TxtFechaIni.Text
'''
'''    End If
'''
'''
''''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''''/////////////////////////////////////BUSCO SI EXISTE NOMINA ACUMULADA //////////////////////////////////////////////
''''////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
''' sql = "SELECT id, NumNomina, CodEmpleado, SalarioBasico, Destajo, HE, DD, HorasExtras, Comisiones, OtrosIngresos, DescripOtrIngre, Incentivos, Deducciones, Prestamo, MontoINSS, MontoIR, Vacaciones, INSSPatronal, IRPatronal, INATEC, Mes13, DiasDescuento, Adelantos, TotalSubsidio, VacacionesPagadas, DiasVacaciones, AdelantosVacaciones, HTrabajada, SeptimoDia, IncetivoProduccion, TarifaHoraria, produjo, BonoProduccion, Viaticos, Ajuste, TIngresos, TGastos, SalarioBasico + Destajo + HorasExtras + Comisiones + OtrosIngresos + Incentivos + SeptimoDia + IncetivoProduccion + BonoProduccion AS TotalDevengado,NQuincenaAcumulada From DetalleNominaAcumulada Where (NumNomina = 0) And (CodEmpleado = " & CodEmpleado & ")"
''' Me.DtaConsulta.RecordSource = sql
''' Me.DtaConsulta.Refresh
'''     If Not Me.DtaConsulta.Recordset.EOF Then
'''       MontoIrAcumulado = Me.DtaConsulta.Recordset("MontoIR") + MontoIrAcumulado
'''       TotalDevengadoAcumulado = TotalDevengadoAcumulado + Me.DtaConsulta.Recordset("TotalDevengado") - Me.DtaConsulta.Recordset("MontoINSS")
'''       VacacionesAcumuladas = Me.DtaConsulta.Recordset("Vacaciones")
'''       If Not IsNull(Me.DtaConsulta.Recordset("NQuincenaAcumulada")) Then
'''         NQuincenas = Me.DtaConsulta.Recordset("NQuincenaAcumulada") + NQuincenas
'''       End If
'''     End If
'''
'''
'''
'''
''''//////////////////////////////////////////////////
''''///BUSCO EL PERIODO DE INGRESO DEL EMPLEADO
''''////////////////////////////////////////////////////////
'''        Me.AdoPeriodoFiscal.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (CodTipoNomina = " & CodTipoNomina & ") AND (Inicio = '" & Format(CDate(FechaIngreso), "DD/MM/YYYY") & "') ORDER BY Periodo"
'''        Me.AdoPeriodoFiscal.Refresh
'''        If Not Me.AdoPeriodoFiscal.Recordset.EOF Then
'''           PeriodoIngreso = Me.AdoPeriodoFiscal.Recordset("Periodo")
'''        End If
'''
'''
'''
''''///////////////////////Verifico si Tiene Ir Porcentual//////////////////////////////
'''CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
'''Me.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, PorcientoInss, TasaInss, PorcientoIr, TasaIr From TipoNomina WHERE (PorcientoIr = 1) AND (CodTipoNomina = '" & CodTipoNomina & "' )"
'''Me.DtaConsulta.Refresh
'''If DtaConsulta.Recordset.EOF Then
''' 'Hago el Calcul del nuevo Techo para el Ir
''' Select Case DtaTipoNomina.Recordset("Periodo")
'''                Case "Semanal Viernes"
'''
'''                    If BuscaUltimaSemana(CDbl(CantSabados), CDbl(NumNomina), Format(Mes, "0#"), CDbl(AnoIni)) = True Then
'''                     MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'''                     MontoBrutoMensual = MontoBruto + TotalSueldoAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes)) - TotalInssAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes))
'''                    ElseIf IrUltimaSemana = False Then
'''                        MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'''                        MontoBrutoMensual = MontoBruto * CantSabados
'''                    Else
'''                        MontoBrutoMensual = 0
'''                    End If
'''
'''                Case "Semanal Sabado"
'''
'''                    If BuscaUltimaSemana(CDbl(CantSabados), CDbl(NumNomina), Format(Mes, "0#"), CDbl(AnoIni)) = True Then
'''                     MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'''                     MontoBrutoMensual = MontoBruto + TotalSueldoAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes)) - TotalInssAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes))
'''                    ElseIf IrUltimaSemana = False Then
'''                        MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'''                        MontoBrutoMensual = MontoBruto * CantSabados
'''                    Else
'''                        MontoBrutoMensual = 0
'''                    End If
'''
'''                Case "Catorcenal los Viernes"
'''                    If DiaFin < 28 Then
'''                     MontoBruto = (TotalDevengado + MontoOtrosIngresos + MontoTipoVacaciones) - MontoInss
'''                     MontoBrutoMensual = ((MontoBruto * 15) / 14) * 2
'''                    Else
'''                     MontoBrutoMensual = SalarioMensual - MontoInssMensual
'''                    End If
'''                Case "Catorcenal los Sabados"
'''                      MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'''                      MontoBrutoMensual = (MontoBruto * 26) / 12
'''                Case "Quincenal"
'''                  If TipoCalculoIr = "Calcular IR x 12" Then
'''                    If DiaFin < 28 Then
'''                      If IrUltimaSemana = False Then
'''                        MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'''                        MontoBrutoMensual = MontoBruto * 2
'''                        MontoBrutoAnual = MontoBrutoMensual * 12
'''                      Else
'''                        MontoBruto = 0
'''                        MontoBrutoMensual = 0
'''                        MontoBrutoAnual = 0
'''                      End If
'''                    ElseIf IrUltimaSemana = False Then
'''                        MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'''                        MontoBrutoMensual = MontoBruto * 2
'''                        MontoBrutoAnual = MontoBrutoMensual * 12
'''                        '                        If TotalDevengadoAnterior = 0 Then
''''                           MontoBrutoMensual = (SalarioMensual - MontoInssMensual) * 2
''''                           MontoBrutoAnual = MontoBrutoMensual * 12
''''                        Else
''''                           MontoBrutoMensual = SalarioMensual - MontoInssMensual
''''                           MontoBrutoAnual = MontoBrutoMensual * 12
''''                        End If
'''                    ElseIf IrUltimaSemana = True Then
'''                     MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'''                     MontoBrutoMensual = MontoBruto + TotalSueldoAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes)) - TotalInssAnterior(CDbl(NumNomina), CodEmpleado, CDbl(AnoIni), CDbl(Mes))
'''                     MontoBrutoAnual = MontoBrutoMensual * 12
'''                    End If
'''                  Else
'''                   MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss '+ MontoOtrosIngresos
'''                   RentaGravable = ((TotalDevengadoAcumulado + MontoBruto) / NQuincenas) * 24
'''
'''                   MontoBrutoAnual = RentaGravable '+ MontoVacaciones + VacacionesAcumuladas
'''                   MontoBrutoMensual = MontoBruto * 2
'''                  End If
'''
'''                Case "Mensual"
'''
'''                   MontoBruto = (TotalDevengado + MontoVacaciones + MontoDestajos + Septimo + TotalSalarioxHora + IncentivoProduccion + MontoHRSExtras + MontoViaticos + MontoIncentivos + MontoComisiones + MontoHorasTurno + MontoTipoVacaciones) - MontoInss
'''                   RentaGravable = ((TotalDevengadoAcumulado + MontoBruto) * (12 - (PeriodoIngreso - 1))) / NQuincenas
''''                   MontoBrutoAnual = RentaGravable + MontoVacaciones + VacacionesAcumuladas
'''                   MontoBrutoMensual = MontoBruto
'''                   MontoBrutoAnual = MontoBrutoMensual * 12
''''                    MontoBruto = SalarioMensual - MontoInssMensual
''''                    MontoBrutoMensual = MontoBruto
'''                Case "Trimestral"
'''
'''                    MontoBruto = SalarioMensual - MontoInssMensual
'''                    MontoBrutoMensual = MontoBruto / 3
'''                Case "Semestral"
'''
'''                    MontoBruto = SalarioMensual - MontoInssMensual
'''                    MontoBrutoMensual = MontoBruto / 6
'''End Select
'''
'''
'''  '//////////////////////////////////////////////////////////////////////////
'''  '///////////////////BUSCO EL TIPO DE MONEDA DE LA NOMINA///////////////////
'''  '//////////////////////////////////////////////////////////////////////////
'''   Me.AdoBusca.RecordSource = "SELECT CodTipoNomina, Nomina, Periodo, UltFecha, TipoPago, Moneda, MantValor, Activa, PorcientoInss, TasaInss, PorcientoIr, TasaIr,TasaInssPatronal From TipoNomina WHERE (CodTipoNomina = '" & CodTipoNomina & "')"
'''   Me.AdoBusca.Refresh
'''   If Not Me.AdoBusca.Recordset.EOF Then
'''      Moneda = Me.AdoBusca.Recordset("Moneda")
'''   Else
'''      Moneda = "C$"
'''   End If
'''
'''If DtaEmpleados.Recordset("ExentoIr") = False Then
'''        'agregar IR laboral y patronal
'''
'''        MontoIr = 0
'''        MontoIRPatronal = 0
'''        MontoDolares = 0
'''        If Moneda = "US" Then
'''         MontoDolares = MontoBrutoMensual
'''         MontoBrutoMensual = MontoBrutoMensual * TasaCambio
'''        End If
'''
'''
'''        DtaIr.Refresh
'''        DtaIr.Recordset.MoveNext
'''        MinIR = DtaIr.Recordset("desde")
'''        MinIR = MinIR - 1
'''        MinIR = (MinIR / 12)
'''        Do While Not DtaIr.Recordset.EOF
'''
'''           'ubicar la linea
'''         If DtaTipoNomina.Recordset("Periodo") = "Semanal Viernes" Then
'''            If (MontoBrutoMensual) >= MinIR Then
'''            If DtaIr.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIr.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'''               MontoIr = ((MontoBrutoMensual * 12) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
'''               MontoIr = Format(MontoIr / 12, "###,##0.00")  'MontoIr = Format(MontoIr / CantSabados / 12, "###,##0.00")
'''               MontoIRPatronal = MontoIr
'''               Exit Do
'''            End If
'''            End If
'''
'''         ElseIf DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
'''            If (MontoBrutoMensual) >= MinIR Then
'''            If DtaIr.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIr.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'''               MontoIr = ((MontoBrutoMensual * 12) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
'''               MontoIr = Format(MontoIr / 12, "###,##0.00")
'''               MontoIRPatronal = MontoIr
'''               Exit Do
'''
'''            End If
'''            End If
'''
'''        ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Viernes" Then
'''            If (MontoBrutoMensual) >= MinIR Then
'''            If DtaIr.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIr.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'''               MontoIr = ((MontoBrutoMensual * 12) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
'''  '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
'''               If DiaFin < 28 Then
'''                MontoIr = Format(MontoIr / 2 / 12, "###,##0.00")
'''                MontoIRPatronal = MontoIr
'''                Exit Do
'''               Else
'''                MontoIrMensual = Format(MontoIr / 1 / 12, "###,##0.00")
'''                MontoIr = MontoIrMensual - MontoIrAnterior
'''                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
'''               End If
'''            End If
'''            Else
'''               MontoIrMensual = 0
'''               MontoIr = MontoIrMensual - MontoIrAnterior
'''               MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
'''            End If
'''         ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
'''            If (MontoBrutoMensual) >= MinIR Then
'''            If DtaIr.Recordset("desde") <= (MontoBruto * 26) And DtaIr.Recordset("Hasta") >= (MontoBruto * 26) Then
'''               MontoIr = ((MontoBruto * 26) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
'''  '///////Verfico si el la Ultima Quincena para hacer ajustes////////////
'''               If DiaFin < 28 Then
'''                MontoIr = Format(MontoIr / 26, "###,##0.00")
'''                MontoIRPatronal = MontoIr
'''                Exit Do
'''               Else
'''                MontoIr = Format(MontoIr / 26, "###,##0.00")
'''                MontoIRPatronal = MontoIr
'''               End If
'''            End If
'''            Else
'''               MontoIrMensual = 0
'''                MontoIr = MontoIrMensual - MontoIrAnterior
'''                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
'''            End If
'''
'''
'''         ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
'''             If DtaIr.Recordset("desde") <= (MontoBrutoAnual) And DtaIr.Recordset("Hasta") >= (MontoBrutoAnual) Then
'''               MontoIr = ((MontoBrutoAnual) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
''''     //////////  Verfico si el la Ultima Quincena para hacer ajustes////////////
'''
'''                If TipoCalculoIr = "Calcular IR x 12" Then
'''                    If DiaFin < 28 Then
'''                        MontoIr = Format(MontoIr / 2 / 12, "###,##0.00")
'''                        MontoIRPatronal = MontoIr
'''                        Exit Do
'''                    ElseIf IrUltimaSemana = False Then
'''                        MontoIr = Format(MontoIr / 2 / 12, "###,##0.00")
'''                        MontoIRPatronal = MontoIr
'''                    ElseIf IrUltimaSemana = True Then
'''                        MontoIr = Format(MontoIr / 12, "###,##0.00")
'''                        MontoIRPatronal = MontoIr
'''                    End If
'''                Else
'''                If Not NumeroPeriodo = 0 Then
'''                  'NumeroPeriodo = 24-(NQuincenas-1)
'''                 MontoIr = (MontoIr - MontoIrAcumulado) / NumeroPeriodo
'''                Else
'''                 MontoIr = 0
'''                End If
'''                End If
'''
'''                MontoIRPatronal = MontoIr - MontoIrPatronalAnterior
'''                Exit Do
''''               End If
'''             End If
''''            Else
''''               MontoIrMensual = 0
'''
''''                MontoIR = MontoIrMensual - MontoIrAnterior
''''                MontoIRPatronal = MontoIrMensual - MontoIrPatronalAnterior
''''            End If
'''
'''
'''
'''         ElseIf DtaTipoNomina.Recordset("Periodo") = "Mensual" Then
''''           If (MontoBrutoAnual) >= MinIR Then
'''            If DtaIr.Recordset("desde") <= (MontoBrutoAnual) And DtaIr.Recordset("Hasta") >= (MontoBrutoAnual) Then
'''
'''               MontoIr = ((MontoBrutoAnual) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
'''
'''                MontoIr = (MontoIr - MontoIrAcumulado) / 12
'''                MontoIRPatronal = MontoIr - MontoIrPatronalAnterior
'''                Exit Do
'''
'''               Exit Do
'''            End If
''''         End If
'''         ElseIf DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
'''           If (MontoBrutoMensual) >= MinIR Then
'''            If DtaIr.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIr.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'''               MontoIr = ((MontoBrutoMensual * 12) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
'''               MontoIr = Format(MontoIr / 4, "###,##0.00")
'''               MontoIRPatronal = MontoIr
'''               Exit Do
'''            End If
'''           End If
'''         ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
'''             If (MontoBrutoMensual) >= MinIR Then
'''            If DtaIr.Recordset("desde") <= (MontoBrutoMensual * 12) And DtaIr.Recordset("Hasta") >= (MontoBrutoMensual * 12) Then
'''               MontoIr = ((MontoBrutoMensual * 12) - DtaIr.Recordset("SobreExceso")) * (DtaIr.Recordset("PorcientoImpuesto") / 100) + DtaIr.Recordset("ImpuestoBase")
'''               MontoIr = Format(MontoIr / 2, "###,##0.00")
'''               MontoIRPatronal = MontoIr
'''               Exit Do
'''            End If
'''            End If
'''         End If
'''  DtaIr.Recordset.MoveNext
'''  Loop
'''
'''    If Moneda = "US" Then
'''       MontoBrutoMensual = MontoDolares
'''       If TasaCambio <> 0 Then
'''        MontoIr = MontoIr / TasaCambio
'''        MontoIRPatronal = MontoIRPatronal / TasaCambio
'''       End If
'''    End If
'''
'''  End If 'del if que pregunta si esta excento de IR
'''        'TotalDevengado = TotalDevengado + MontoDestajo + MontoHRSExtras + MontoComisiones + MontoIncentivos
'''Else
'''
'''
'''
'''End If
        'calculo de las vacaciones y el 13mes
        'If CodEmpleado = "0162" Then
         'MsgBox "2"
        'End If
        
    
        
        If MontoIr < 0 Then
          MontoIr = 0
          MontoIRPatronal = 0
        End If
        
        SalarioDevengado = (TotalDevengado - MontoHRSExtras)
        SalarioDevengado = Format(SalarioDevengado, "###,##0.00")
        
'        MontoVacaciones = 0
'        MontoMes13 = 0
'
'        If DtaTipoNomina.Recordset("Periodo") = "Semanal Viernes" Or DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
'            If CantSabados = 5 Then
'            MontoVacaciones = Salario * 0.5 / 7
'            MontoMes13 = Salario * 0.5 / 7
'            Else
'            MontoVacaciones = Salario * 0.625 / 7
'            MontoMes13 = Salario * 0.625 / 7
'            End If
'
'        ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Viernes" Or DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then
'            '-----FORMULA: 2071.23 * 26 CATORCENAS  *  1/12 * 1/12 * 1/30 = 4320
'            MontoVacaciones = ((Salario * 26) / 4320) * 14
'            MontoMes13 = ((Salario * 26) / 4320) * 14
'        ElseIf DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
'            MontoVacaciones = Salario * 1.25 / 15
'            MontoMes13 = Salario * 1.25 / 15
'        ElseIf DtaTipoNomina.Recordset("Periodo") = "Mensual" Then
'            MontoVacaciones = Salario * 2.5 / 30
'            MontoMes13 = Salario * 2.5 / 30
'        ElseIf DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
'            MontoVacaciones = Salario * 7.5 / 90
'            MontoMes13 = Salario * 7.5 / 90
'        ElseIf DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
'            MontoVacaciones = Salario * 15 / 180
'            MontoMes13 = Salario * 15 / 180
'        End If
        
        

          '///////////////////////////////////////////////////////////////////////////////
          '//////BUSCO LAS SUSPECIONES DE LOS EMPLEADOS PARA MANTENER/////////////////////
          '/////LA NOVEDAD 09 ////////////////////////////////////////////////////////////
            Me.AdoSuspendido.RecordSource = "SELECT CodEmpleado, CodEmpleado1, Fechaini, FechaFin, Motivo, Activo, Ultimo " & _
            "From Subsidios Where (Activo = 1) And (CodEmpleado = " & CodEmpleado & ")"
            Me.AdoSuspendido.Refresh
            If Me.AdoSuspendido.Recordset.EOF Then
              agregar = True
            Else
              agregar = False
            End If
            


'//////////////Busco si la Nomina Existe para Editarla/////////////////
      AjusteINSS = 0
  DtaDetalleNomina.RecordSource = "SELECT DetalleNomina.id, DetalleNomina.BonoProduccion ,DetalleNomina.IncetivoProduccion,DetalleNomina.NumNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HE, DetalleNomina.DD, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.INATEC, DetalleNomina.Mes13, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.TotalSubsidio, DetalleNomina.VacacionesPagadas, DetalleNomina.DiasVacaciones,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.TarifaHoraria,DetalleNomina.produjo,DetalleNomina.AjusteINSS,HTurno, HorasTurno,Antiguedad, AñoAntiguedad,  DetalleNomina.DiasAdicionales, DetalleNomina.ValorDiasAdicionales" & _
                                  " ,DetalleNomina.Reembolso From DetalleNomina Where (((DetalleNomina.NumNomina) = " & NumNomina & ") And ((DetalleNomina.CodEmpleado) = '" & CodEmpleado & "'))"
  DtaDetalleNomina.Refresh
      MontoInssBasico = ((TarifaHorariaBasico * 8 * DiasMes) * (TasaInss / 100) / CantSabados)
      If MontoInss > MontoInssBasico Then
        AjusteINSS = 0
      Else
        AjusteINSS = MontoInssBasico - MontoInss
      End If
      
      

  If Not DtaDetalleNomina.Recordset.EOF Then
   

        'Edito el Registro Existente
       If agregar = True Then
        DtaDetalleNomina.Recordset("NumNomina") = NumNomina
        DtaDetalleNomina.Recordset("TarifaHoraria") = MontoHora 'TarifaHoraria
        DtaDetalleNomina.Recordset("BonoProduccion") = BonoProduccion
        DtaDetalleNomina.Recordset("SeptimoDia") = Septimo
        DtaDetalleNomina.Recordset("DiasVacaciones") = DiasVacaciones   ' ConDiasVacaciones
        DtaDetalleNomina.Recordset("HTrabajada") = TotalHoras
        DtaDetalleNomina.Recordset("IncetivoProduccion") = IncentivoProduccion
        DtaDetalleNomina.Recordset("CodEmpleado") = CodEmpleado
        If ((DtaEmpleados.Recordset("SueldoPeriodo") + TotalSalarioxHora) * Factor) = 0 Then
          DtaDetalleNomina.Recordset("produjo") = "S"
        Else
          DtaDetalleNomina.Recordset("produjo") = "N"
        End If
        DtaDetalleNomina.Recordset("Reembolso") = Reembolso
        DtaDetalleNomina.Recordset("SalarioBasico") = (Salario + AumentoBasico + MontoSubsidio) * Factor
        DtaDetalleNomina.Recordset("destajo") = MontoDestajos * Factor
        DtaDetalleNomina.Recordset("HE") = HE
        DtaDetalleNomina.Recordset("HorasExtras") = Redondear(MontoHRSExtras) * Factor
        DtaDetalleNomina.Recordset("Comisiones") = MontoComisiones + TotalPuntualidad * Factor
        DtaDetalleNomina.Recordset("incentivos") = (MontoIncentivos + MontoIncentivoExcento) * Factor
        DtaDetalleNomina.Recordset("OtrosIngresos") = MontoOtrosIngresos * Factor
        DtaDetalleNomina.Recordset("DescripOtrIngre") = DescripOtrIngre
        DtaDetalleNomina.Recordset("Deducciones") = MontoDeduccion * Factor
        DtaDetalleNomina.Recordset("Prestamo") = MontoPrestamo * Factor
        DtaDetalleNomina.Recordset("MontoInss") = Redondear(MontoInss) * Factor
        DtaDetalleNomina.Recordset("MontoIR") = MontoIr * Factor
        DtaDetalleNomina.Recordset("Vacaciones") = MontoVacaciones * Factor
        DtaDetalleNomina.Recordset("Mes13") = MontoMes13 * Factor
        DtaDetalleNomina.Recordset("INSSPatronal") = MontoInssPatronal * Factor
        DtaDetalleNomina.Recordset("IRPatronal") = MontoIRPatronal * Factor
        DtaDetalleNomina.Recordset("INATEC") = INATEC * Factor
        DtaDetalleNomina.Recordset("VacacionesPagadas") = MontoTipoVacaciones * Factor
        'DtaDetalleNomina.Recordset.AdelantoVacaciones = MontoAdelantoVaca * Factor
        'DtaDetalleNomina.Recordset.Adelanto13voMes = MontoAdelanto13 * Factor
        DtaDetalleNomina.Recordset("DD") = ConDiasDescuento * Factor
        DtaDetalleNomina.Recordset("Adelantos") = Adelantos
        DtaDetalleNomina.Recordset("DiasDescuento") = DeduccionPorFalta
        DtaDetalleNomina.Recordset("AjusteINSS") = (AjusteINSS) * Factor
        DtaDetalleNomina.Recordset("HTurno") = HT
        DtaDetalleNomina.Recordset("HorasTurno") = MontoHorasTurno * Factor
        DtaDetalleNomina.Recordset("AñoAntiguedad") = Int(Anos)
        DtaDetalleNomina.Recordset("Antiguedad") = Antiguedad * Factor
        DtaDetalleNomina.Recordset("ValorDiasAdicionales") = MontoDiasAdicionales
        DtaDetalleNomina.Recordset("DiasAdicionales") = DiasAdicionales
        DtaDetalleNomina.Recordset.Update
      Else '////////si esta suspendido grabo todo en cero/////
        DtaDetalleNomina.Recordset("NumNomina") = NumNomina
        DtaDetalleNomina.Recordset("TarifaHoraria") = 0
        DtaDetalleNomina.Recordset("DiasVacaciones") = 0
         DtaDetalleNomina.Recordset("TotalSubsidio") = 0
        DtaDetalleNomina.Recordset("SeptimoDia") = 0
        DtaDetalleNomina.Recordset("HTrabajada") = 0
        DtaDetalleNomina.Recordset("IncetivoProduccion") = 0
        DtaDetalleNomina.Recordset("CodEmpleado") = CodEmpleado
        DtaDetalleNomina.Recordset("SalarioBasico") = 0
        DtaDetalleNomina.Recordset("destajo") = 0
        DtaDetalleNomina.Recordset("HE") = 0
        DtaDetalleNomina.Recordset("HorasExtras") = 0
        DtaDetalleNomina.Recordset("Comisiones") = 0
        DtaDetalleNomina.Recordset("incentivos") = 0
        DtaDetalleNomina.Recordset("OtrosIngresos") = 0
        DtaDetalleNomina.Recordset("DescripOtrIngre") = DescripOtrIngre
        DtaDetalleNomina.Recordset("Deducciones") = 0
        DtaDetalleNomina.Recordset("Prestamo") = 0
        DtaDetalleNomina.Recordset("MontoInss") = 0
        DtaDetalleNomina.Recordset("MontoIR") = 0
        DtaDetalleNomina.Recordset("Vacaciones") = 0
        DtaDetalleNomina.Recordset("Mes13") = 0
        DtaDetalleNomina.Recordset("INSSPatronal") = 0
        DtaDetalleNomina.Recordset("IRPatronal") = 0
        DtaDetalleNomina.Recordset("INATEC") = 0
        DtaDetalleNomina.Recordset("DD") = 0
        DtaDetalleNomina.Recordset("Adelantos") = 0
        DtaDetalleNomina.Recordset("DiasDescuento") = 0
        DtaDetalleNomina.Recordset("AjusteINSS") = 0
        DtaDetalleNomina.Recordset("HTurno") = 0
        DtaDetalleNomina.Recordset("HorasTurno") = 0
        DtaDetalleNomina.Recordset("AñoAntiguedad") = 0
        DtaDetalleNomina.Recordset("Antiguedad") = 0
        DtaDetalleNomina.Recordset("VacacionesPagadas") = 0
        DtaDetalleNomina.Recordset("ValorDiasAdicionales") = 0
        DtaDetalleNomina.Recordset("DiasAdicionales") = 0
        DtaDetalleNomina.Recordset("Reembolso") = 0
        DtaDetalleNomina.Recordset.Update
       End If
     Else
        'Agrego un nuevo Registro
       If agregar = True Then
 

        DtaDetalleNomina.Recordset.AddNew
'        DtaDetalleNomina.Recordset("id") = ID
        DtaDetalleNomina.Recordset("NumNomina") = NumNomina
        DtaDetalleNomina.Recordset("SeptimoDia") = Redondear(Septimo)
        DtaDetalleNomina.Recordset("BonoProduccion") = Redondear(BonoProduccion)
        DtaDetalleNomina.Recordset("TarifaHoraria") = Redondear(MontoHora) 'TarifaHoraria
        DtaDetalleNomina.Recordset("IncetivoProduccion") = Redondear(IncentivoProduccion)
        DtaDetalleNomina.Recordset("HTrabajada") = TotalHoras
        DtaDetalleNomina.Recordset("CodEmpleado") = CodEmpleado
        'DtadetalleNomina.Recordset("SalarioBasico") = DtaEmpleados.Recordset("SueldoPeriodo")
        'DtadetalleNomina.Recordset("SalarioBasico") = TotalDevengado
        If ((DtaEmpleados.Recordset("SueldoPeriodo") + TotalSalarioxHora) * Factor) = 0 Then
         DtaDetalleNomina.Recordset("produjo") = "S"
        Else
         DtaDetalleNomina.Recordset("produjo") = "N"
        End If
        DtaDetalleNomina.Recordset("DiasVacaciones") = DiasVacaciones 'ConDiasVacaciones
        DtaDetalleNomina.Recordset("SalarioBasico") = Redondear((Salario + AumentoBasico + MontoSubsidio)) * Factor
        DtaDetalleNomina.Recordset("destajo") = Redondear(MontoDestajos) * Factor
        DtaDetalleNomina.Recordset("HE") = HE
        DtaDetalleNomina.Recordset("HorasExtras") = MontoHRSExtras * Factor
        DtaDetalleNomina.Recordset("Comisiones") = MontoComisiones + TotalPuntualidad * Factor
        DtaDetalleNomina.Recordset("incentivos") = (MontoIncentivos + MontoIncentivoExcento) * Factor
        DtaDetalleNomina.Recordset("OtrosIngresos") = MontoOtrosIngresos * Factor
        DtaDetalleNomina.Recordset("DescripOtrIngre") = DescripOtrIngre
        DtaDetalleNomina.Recordset("Deducciones") = MontoDeduccion * Factor
        DtaDetalleNomina.Recordset("Prestamo") = MontoPrestamo * Factor
        DtaDetalleNomina.Recordset("MontoInss") = Redondear(MontoInss) * Factor
        DtaDetalleNomina.Recordset("MontoIR") = MontoIr * Factor
        DtaDetalleNomina.Recordset("Vacaciones") = MontoVacaciones * Factor
         DtaDetalleNomina.Recordset("TotalSubsidio") = MontoSubsidio
        DtaDetalleNomina.Recordset("Mes13") = MontoMes13 * Factor
        DtaDetalleNomina.Recordset("INSSPatronal") = MontoInssPatronal * Factor
        DtaDetalleNomina.Recordset("IRPatronal") = MontoIRPatronal * Factor
        DtaDetalleNomina.Recordset("INATEC") = INATEC * Factor
        'DtaDetalleNomina.Recordset.AdelantoVacaciones = MontoAdelantoVaca * Factor
        'DtaDetalleNomina.Recordset.Adelanto13voMes = MontoAdelanto13 * Factor
        DtaDetalleNomina.Recordset("DD") = ConDiasDescuento * Factor
        DtaDetalleNomina.Recordset("Adelantos") = Adelantos
        DtaDetalleNomina.Recordset("DiasDescuento") = DeduccionPorFalta
        DtaDetalleNomina.Recordset("AjusteINSS") = (AjusteINSS) * Factor
        DtaDetalleNomina.Recordset("HTurno") = HT
        DtaDetalleNomina.Recordset("HorasTurno") = MontoHorasTurno * Factor
        DtaDetalleNomina.Recordset("AñoAntiguedad") = Int(Anos)
        DtaDetalleNomina.Recordset("Antiguedad") = Antiguedad * Factor
        DtaDetalleNomina.Recordset("VacacionesPagadas") = MontoTipoVacaciones * Factor
        DtaDetalleNomina.Recordset("ValorDiasAdicionales") = MontoDiasAdicionales
        DtaDetalleNomina.Recordset("DiasAdicionales") = DiasAdicionales
        DtaDetalleNomina.Recordset("Reembolso") = Reembolso
        DtaDetalleNomina.Recordset.Update
        
       Else '//SI ESTA SUSPENDIDO LLENO TODOS LOS VALORES NUEVOS EN CERO/////
       
        DtaDetalleNomina.Recordset.AddNew
'        DtaDetalleNomina.Recordset("id") = ID
        DtaDetalleNomina.Recordset("NumNomina") = NumNomina
        DtaDetalleNomina.Recordset("SeptimoDia") = 0
        DtaDetalleNomina.Recordset("DiasVacaciones") = 0
        DtaDetalleNomina.Recordset("TarifaHoraria") = 0
        DtaDetalleNomina.Recordset("IncetivoProduccion") = 0
        DtaDetalleNomina.Recordset("HTrabajada") = 0
        DtaDetalleNomina.Recordset("CodEmpleado") = CodEmpleado
        DtaDetalleNomina.Recordset("SalarioBasico") = 0
        DtaDetalleNomina.Recordset("destajo") = 0
        DtaDetalleNomina.Recordset("HE") = 0
        DtaDetalleNomina.Recordset("HorasExtras") = 0
        DtaDetalleNomina.Recordset("Comisiones") = 0
        DtaDetalleNomina.Recordset("incentivos") = 0
        DtaDetalleNomina.Recordset("OtrosIngresos") = 0
        DtaDetalleNomina.Recordset("DescripOtrIngre") = DescripOtrIngre
        DtaDetalleNomina.Recordset("Deducciones") = 0
        DtaDetalleNomina.Recordset("Prestamo") = 0
        DtaDetalleNomina.Recordset("MontoInss") = 0
        DtaDetalleNomina.Recordset("MontoIR") = 0
        DtaDetalleNomina.Recordset("Vacaciones") = 0
         DtaDetalleNomina.Recordset("TotalSubsidio") = 0
        DtaDetalleNomina.Recordset("Mes13") = 0
        DtaDetalleNomina.Recordset("INSSPatronal") = 0
        DtaDetalleNomina.Recordset("IRPatronal") = 0
        DtaDetalleNomina.Recordset("INATEC") = 0
        DtaDetalleNomina.Recordset("DD") = 0
        DtaDetalleNomina.Recordset("Adelantos") = 0
        DtaDetalleNomina.Recordset("DiasDescuento") = 0
        DtaDetalleNomina.Recordset("HTurno") = 0
        DtaDetalleNomina.Recordset("HorasTurno") = 0
        DtaDetalleNomina.Recordset("AñoAntiguedad") = 0
        DtaDetalleNomina.Recordset("Antiguedad") = 0
        DtaDetalleNomina.Recordset("VacacionesPagadas") = 0
        DtaDetalleNomina.Recordset("ValorDiasAdicionales") = 0
        DtaDetalleNomina.Recordset("DiasAdicionales") = 0
        DtaDetalleNomina.Recordset.Update
       End If

     End If
     
'/////////////Edito la Nomina Principal///////////////////////////

       If agregar = True Then
        DtaNomina.Recordset("TotalSalarioBasico") = DtaNomina.Recordset("TotalSalarioBasico") + (DtaEmpleados.Recordset("SueldoPeriodo") * Factor)
        DtaNomina.Recordset("TotalDestajo") = DtaNomina.Recordset("TotalDestajo") + (MontoDestajos * Factor)
        DtaNomina.Recordset("TotalHorasExtras") = DtaNomina.Recordset("TotalHorasExtras") + (MontoHRSExtras * Factor)
        DtaNomina.Recordset("TotalComisiones") = DtaNomina.Recordset("TotalComisiones") + (MontoComisiones * Factor)
        DtaNomina.Recordset("TotalIncentivos") = DtaNomina.Recordset("TotalIncentivos") + (MontoIncentivos * Factor)
        DtaNomina.Recordset("TotalOtrosIngresos") = DtaNomina.Recordset("TotalOtrosIngresos") + (MontoOtrosIngresos * Factor)
        DtaNomina.Recordset("TotalDeducciones") = DtaNomina.Recordset("TotalDeducciones") + (MontoDeduccion * Factor)
        DtaNomina.Recordset("TotalPrestamo") = DtaNomina.Recordset("TotalPrestamo") + (MontoPrestamo * Factor)
        DtaNomina.Recordset("TotalMontoInss") = DtaNomina.Recordset("TotalMontoInss") + (MontoInss * Factor)
        DtaNomina.Recordset("TotalMontoIR") = DtaNomina.Recordset("TotalMontoIR") + (MontoIr * Factor)
        DtaNomina.Recordset("TotalVacaciones") = DtaNomina.Recordset("TotalVacaciones") + (MontoVacaciones * Factor)
        DtaNomina.Recordset("TotalINSSPatronal") = DtaNomina.Recordset("TotalINSSPatronal") + (MontoInssPatronal * Factor)
        DtaNomina.Recordset("TotalIRPatronal") = DtaNomina.Recordset("TotalIRPatronal") + (MontoIRPatronal * Factor)
        DtaNomina.Recordset("Totalmes13") = DtaNomina.Recordset("Totalmes13") + (MontoMes13 * Factor)
        DtaNomina.Recordset("TotalInatec") = DtaNomina.Recordset("TotalInatec") + INATEC
        DtaNomina.Recordset.Update
       End If
        
        DtaEmpleados.Recordset.MoveNext
        i = i + 1
Loop
End With

'marcar como no activa
Me.MousePointer = 1

Exit Sub
TipoErrs:
ControlErrores
Unload Me
End Sub

Private Sub CmdCerrarNomina_Click()
On Error GoTo TipoErr:
Dim i As Integer, j As Integer
Dim SQlPrestamo As String, FechaInicial As Date
Dim Letra As String
Dim SqlEmpleados As String
Dim SQlIncentivos As String
Dim SQlDeducciones As String
Dim Anno As Integer
Dim Mes As String
Dim SqlPagosMensuales As String
Dim Periodo As String
Dim HayMeses As String
Dim CantEmpleados As Double, cantidad As Double
Dim TotalEmpleado As Double
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim NumVez As String, Consecutivo As Double, CodTipoSubsidio As String
Dim Valor As Double, Descripcion As String



'Anno = Year(Now)
'Mes = Month(Now)

Letra = "n"
k% = MsgBox("Desea cerrar la nómina?", vbYesNo)
If k% <> 6 Then Exit Sub


MousePointer = 11

CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
Periodo = DtaTipoNomina.Recordset("Periodo")
SQLNomina = "SELECT Nomina.* From Nomina WHERE Nomina.Activa=1 AND Nomina.CodTipoNomina= '" & CodTipoNomina & "'"
DtaNomina.RecordSource = SQLNomina
DtaNomina.Refresh

NumNomina = DtaNomina.Recordset("NumNomina")
Anno = Year(DtaNomina.Recordset("FechaNomina"))
Mes = Month(DtaNomina.Recordset("FechaNomina"))

   res = Bitacora(Now, NombreUsuario, "Calcular Nomina", "Se Cerro la Nomina: " & NumNomina)

If DtaNomina.Recordset("Procesada") = False Then
   MsgBox "Esta Nomina no ha sido Procesada, no puede ser cerrada"
   Exit Sub
End If

'revizo si hay nóminas de subsidio abiertas y no la cierro
Me.DtaNomSubsidios.RecordSource = "SELECT NumNomina, TotalNomSubsidio, FechaPago, Activa, Procesada, Cerrada From NomSubsidio Where (NumNomina = " & NumNomina & ")"
DtaNomSubsidios.Refresh
Do While Not DtaNomSubsidios.Recordset.EOF
If DtaNomSubsidios.Recordset("NumNomina") = NumNomina And DtaNomSubsidios.Recordset("Activa") = 1 Then
   MsgBox "La Nómina de Subsidio Correspondiente está activa Tiene que cerrarla para poder cerrar esta nómina"
   MousePointer = 1
   Exit Sub
End If
DtaNomSubsidios.Recordset.MoveNext
Loop

'debo crear un incentivo y una deduccion si hay de n veces
'por cada empleado con el # de nómina correspondiente

'////////////// Inicio el Calculo de la Nomina ////////////////////
'////////////// Actualizo el Control Progress Bar /////////////////
'SQLNominaEmpleado = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, Empleado.ExentoIr, Empleado.OtrosIngresos,  Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, Empleado.Ausente From Empleado WHERE Empleado.CodTipoNomina= '" & CodTipoNomina & "' AND Empleado.Activo=1 AND Empleado.Ausente=0"
SQLNominaEmpleado = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, Empleado.ExentoIr, Empleado.OtrosIngresos,  Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, Empleado.Ausente From Empleado WHERE Empleado.CodTipoNomina= '" & CodTipoNomina & "' AND Empleado.Activo=1"
DtaEmpleados.RecordSource = SQLNominaEmpleado
DtaEmpleados.Refresh
DtaEmpleados.Recordset.MoveLast
CantEmpleados = DtaEmpleados.Recordset.RecordCount
DtaEmpleados.Refresh

'SqlEmpleados = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, Empleado.ExentoIr, Empleado.OtrosIngresos,  Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, Empleado.Ausente From Empleado WHERE Empleado.CodTipoNomina= '" & CodTipoNomina & "' AND Empleado.Activo=1 AND Empleado.Ausente=0"
SqlEmpleados = "SELECT Empleado.CodEmpleado1,Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.Sindicalista, Empleado.CodDepartamento, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.PorcentajeComision, Empleado.ExentoInss, Empleado.ExentoIr, Empleado.OtrosIngresos,  Empleado.DescripOtrIngre, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Activo, Empleado.Ausente, Empleado.AumentoBasico From Empleado WHERE Empleado.CodTipoNomina= '" & CodTipoNomina & "' AND Empleado.Activo=1"
DtaEmpleados.RecordSource = SqlEmpleados
DtaEmpleados.Refresh

'CantEmpleados = DtaEmpleados.Recordset.RecordCount
With PBCalcNomina
.Min = 0
.Max = CantEmpleados
.Value = 0
j = 1
TotalEmpleado = 0
Me.Label1.Caption = "Procesando la Nomina"
Do While Not DtaEmpleados.Recordset.EOF
 
   
    Me.Caption = "Procesando:  " & j & " de " & CantEmpleados & " Empleados "
    Me.LblTotal.Caption = "Procesando:  " & j & " de " & CantEmpleados & " Empleados "
    
    
    If Me.Label1.Caption = "Procesando la Nomina" Then
         Me.Label1.Caption = "Cerrando Incentivos,Deducciones"
        ElseIf Me.Label1.Caption = "Cerrando Incentivos,Deducciones." Then
            Me.Label1.Caption = "Cerrando Incentivos,Deducciones.."
         ElseIf Me.Label1.Caption = "Cerrando Incentivos,Deducciones.." Then
            Me.Label1.Caption = "Cerrando Incentivos,Deducciones..."
           ElseIf Me.Label1.Caption = "Cerrando Incentivos,Deducciones..." Then
             Me.Label1.Caption = "Cerrando Incentivos,Deducciones...."
          ElseIf Me.Label1.Caption = "Cerrando Incentivos,Deducciones...." Then
           Me.Label1.Caption = "Cerrando Incentivos,Deducciones"
          ElseIf Me.Label1.Caption = "Cerrando Incentivos,Deducciones" Then
           Me.Label1.Caption = "Cerrando Incentivos,Deducciones."
        End If
    
    DoEvents
    
'////////////////Sistema de Nominas///////////////////////////////////
'////////////////Elimino los registros Basicos //////////////////////
'Me.DtaEmpleados.Recordset.Edit
 DtaEmpleados.Recordset("DiasDescuento") = 0
' DtaEmpleados.Recordset("TarifaHoraria") = 0
 DtaEmpleados.Recordset("PorcentajeComision") = 0
 DtaEmpleados.Recordset("OtrosIngresos") = 0
 DtaEmpleados.Recordset("DescripOtrIngre") = "Sin Descrip"
 DtaEmpleados.Recordset("AumentoBasico") = 0
Me.DtaEmpleados.Recordset.Update





     CodEmpleado = DtaEmpleados.Recordset("CodEmpleado")
     
     
        Me.DtaHorasProducidas.RecordSource = "SELECT dbo.DetalleHorasProduccion.CodEmpleado, dbo.DetalleHorasProduccion.NumNomina, dbo.DetalleHorasProduccion.NumLinea, dbo.DetalleHorasProduccion.Lunes + dbo.DetalleHorasProduccion.Martes + dbo.DetalleHorasProduccion.Miercoles + dbo.DetalleHorasProduccion.Jueves + dbo.DetalleHorasProduccion.Viernes AS TotalDias,dbo.Empleado.TarifaHoraria,(dbo.DetalleHorasProduccion.Lunes + dbo.DetalleHorasProduccion.Martes + dbo.DetalleHorasProduccion.Miercoles + dbo.DetalleHorasProduccion.Jueves + dbo.DetalleHorasProduccion.Viernes)* dbo.Empleado.TarifaHoraria AS TotalSalario, dbo.DetalleHorasProduccion.Pagado FROM dbo.DetalleHorasProduccion INNER JOIN dbo.Empleado ON dbo.DetalleHorasProduccion.CodEmpleado = dbo.Empleado.CodEmpleado WHERE (dbo.DetalleHorasProduccion.CodEmpleado = '" & CodEmpleado & "')  AND (dbo.DetalleHorasProduccion.Pagado = 0)"
        Me.DtaHorasProducidas.Refresh
        Do While Not Me.DtaHorasProducidas.Recordset.EOF
           Me.DtaHorasProducidas.Recordset("Pagado") = 1
           Me.DtaHorasProducidas.Recordset.Update
         Me.DtaHorasProducidas.Recordset.MoveNext
        Loop
        
        
        Me.DtaDestajo.RecordSource = "SELECT CodProceso, CodReferencia, Ref, Precio, Unidad, Lunes, Martes, Miercoles, Jueves, Viernes, Sabado, Domingo, TotalUnidades, SalarioPieza,CodEmpleado , NumNomina, Pagado From DetalleProduccion WHERE     (CodEmpleado = '" & CodEmpleado & "') AND (NumNomina = " & NumNomina & ") AND (Pagado = 0)"
        Me.DtaDestajo.Refresh
        Do While Not DtaDestajo.Recordset.EOF
          Me.DtaDestajo.Recordset("Pagado") = 1
          Me.DtaDestajo.Recordset.Update
          Me.DtaDestajo.Recordset.MoveNext
        Loop
        
        Me.AdoDetalleProduccionManual.RecordSource = "SELECT *  From DetalleProduccionManual "
        Me.AdoDetalleProduccionManual.Refresh
        Do While Not Me.AdoDetalleProduccionManual.Recordset.EOF
         Me.AdoDetalleProduccionManual.Recordset("Pagado") = 1
         Me.AdoDetalleProduccionManual.Recordset.Update
         
         Me.AdoDetalleProduccionManual.Recordset.MoveNext
        Loop

     SQlIncentivos = "SELECT Incentivo.NumIncentivo, Incentivo.CodEmpleado, Incentivo.CodTipoIncentivo, DetalleIncentivo.Valor, DetalleIncentivo.NumVez, DetalleIncentivo.Pagado, DetalleIncentivo.NumNomina FROM Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo WHERE Incentivo.CodEmpleado= '" & CodEmpleado & "' AND DetalleIncentivo.NumVez <> '" & Letra & "' AND (DetalleIncentivo.Pagado = 0) "
     DtaIncentivos.RecordSource = SQlIncentivos
     DtaIncentivos.Refresh
     Do While Not DtaIncentivos.Recordset.EOF
            '        Me.DtaDetalleIncentivo.Refresh
            '        If Me.DtaDetalleIncentivo.Recordset.EOF Then
            '          ID = 1
            '        Else
            '         Me.DtaDetalleIncentivo.Recordset.MoveLast
            '         ID = Me.DtaDetalleIncentivo.Recordset("id") + 1
            '        End If
        
            '        DtaDetalleIncentivo.Recordset.AddNew
            '        DtaDetalleIncentivo.Recordset("Numincentivo") = DtaIncentivos.Recordset("Numincentivo")
            '        DtaDetalleIncentivo.Recordset("valor") = DtaIncentivos.Recordset("valor")
            '        DtaDetalleIncentivo.Recordset("NumVez") = 1
            '        DtaDetalleIncentivo.Recordset("NumNomina") = NumNomina
            '        DtaDetalleIncentivo.Recordset("pagado") = 1
            '        DtaDetalleIncentivo.Recordset("id") = ID
            '        DtaDetalleIncentivo.Recordset.Update

    DtaIncentivos.Recordset.MoveNext
    Loop
     
     
     
  '////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  '///////////////////////////////////BUSCO LOS DATOS DE LA NOMINA DE SUBSIDIOS////////////////////////////////////
  '//////////////////////////////////////////////////////////////////////////////////////////////////////////////
      
      
       Me.DtaConsulta.RecordSource = "SELECT * FROM  DetalleSubsidio INNER JOIN Subsidio ON DetalleSubsidio.NumSubsidio = Subsidio.NumSubsidio " & _
                                     "WHERE (DetalleSubsidio.NumNominaSubsidio = " & NumNomina & ") And (Subsidio.CodEmpleado = " & CodEmpleado & ")"
       Me.DtaConsulta.Refresh
        Do While Not DtaConsulta.Recordset.EOF
           If IsNumeric(Me.DtaConsulta.Recordset("NumVez")) Then
             NumVez = val(Me.DtaConsulta.Recordset("NumVez")) - 1
           ElseIf Me.DtaConsulta.Recordset("NumVez") = "n" Then
             NumVez = "n"
           Else
             NumVez = 0
           End If
           
           Consecutivo = ConsecutivoSubsidio("Subsidio")
           CodTipoSubsidio = Me.DtaConsulta.Recordset("CodTipoSubsidio")
           Valor = Me.DtaConsulta.Recordset("Valor")
           If Not IsNull(Me.DtaConsulta.Recordset("Descripcion")) Then
             Descripcion = Me.DtaConsulta.Recordset("Descripcion")
           Else
             Descripcion = ""
           End If
           
           
           If NumVez <> "0" Then
            '///////////////////////////////AGREO LA NOMINA DE SUBSIDIO //////////////////////////////////////////////////
            rs.Open "INSERT INTO Subsidio ([NumSubsidio],[CodEmpleado],[CodTipoSubsidio],[NumVeces],[Pagado]) Values (" & Consecutivo & "," & CodEmpleado & ",'" & CodTipoSubsidio & "','" & NumVez & "',0)", Conexion
            rs.Open "INSERT INTO DetalleSubsidio ([Id],[NumSubsidio],[Valor],[NumVez],[Pagado],[Descripcion]) Values (" & Consecutivo & "," & Consecutivo & "," & Valor & " ,'" & NumVez & "',0, '" & Descripcion & "')", Conexion
           End If
        
          

           Me.DtaConsulta.Recordset.MoveNext
        Loop
     
     
     
DtaEmpleados.Recordset.MoveNext
.Value = TotalEmpleado
If CantEmpleados <> TotalEmpleado Then
    TotalEmpleado = TotalEmpleado + 1
 End If
 j = j + 1
Loop

End With


'//////////BUSCO EL PERIODO FISCAL PARA CERRARLO/////////////////
'////////////////////////////////////////////////////////////////////////
     DtaNominas.Refresh
     Do While Not DtaNominas.Recordset.EOF
     If Me.DtaNominas.Recordset("CodTipoNomina") = Me.DtaTipoNomina.Recordset("CodTipoNomina") And DtaNominas.Recordset("Activa") = True Then
      FechaInicial = Format(DtaNominas.Recordset("FechaNominaINI"), "dd/mm/yyyy")
     End If
      DtaNominas.Recordset.MoveNext
     Loop
        
       Me.DtaConsulta.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual From PeriodoFiscal WHERE (CodTipoNomina = " & CodTipoNomina & ") AND (Inicio = '" & Format(CDate(FechaInicial), "DD/MM/YYYY") & "')ORDER BY Periodo"
       Me.DtaConsulta.Refresh
        If Not DtaConsulta.Recordset.EOF Then
            Me.DtaConsulta.Recordset("Actual") = 0
            
'            Mes = Me.DtaConsulta.Recordset("mes")
            Me.DtaConsulta.Recordset.Update

         End If

'  rs.Open "UPDATE DetalleDeduccion SET DetalleDeduccion.Pagado = 0 WHERE DetalleDeduccion.NumNomina= " & NumNomina & " AND DetalleDeduccion.numvez<> '" & Letra & "'", Conexion
'  rs.Open "UPDATE Deduccion SET Deduccion.Pagado = 1 WHERE Deduccion.NumNomina= " & NumNomina & ""

  rs.Open "UPDATE DetalleDeduccion SET DetalleDeduccion.Pagado = 1 WHERE DetalleDeduccion.NumNomina= " & NumNomina & " AND DetalleDeduccion.numvez<> '" & Letra & "'", Conexion
  rs.Open "UPDATE DetalleIncentivo SET DetalleIncentivo.Pagado = 1 WHERE DetalleIncentivo.NumNomina= " & NumNomina & " AND DetalleIncentivo.NUmVez<> '" & Letra & "'"
  rs.Open "UPDATE MovPrestamo SET MovPrestamo.Cancelado = 1 WHERE MovPrestamo.NumNomina= " & NumNomina & ""
  rs.Open "UPDATE MovPrestamo SET MovPrestamo.Cancelado = 1 WHERE MovPrestamo.NumNomina= " & NumNomina & ""
  rs.Open "UPDATE Comisiones SET Comisiones.Pagado = 1 WHERE Comisiones.NumNomina= " & NumNomina & ""
  rs.Open "UPDATE Destajo SET Destajo.Pagado = 1 WHERE Destajo.NumNomina= " & NumNomina & ""
  rs.Open "UPDATE HorasExtras SET HorasExtras.Pagada = 1 WHERE HorasExtras.NumNomina= " & NumNomina & ""
  rs.Open "UPDATE Nomina SET Nomina.Cerrada = 1 WHERE Nomina.NumNomina= " & NumNomina & ""
  rs.Open "UPDATE Nomina SET Nomina.Activa = 0 WHERE Nomina.NumNomina= " & NumNomina & ""
  rs.Open "UPDATE TipoNomina SET TipoNomina.Activa = 0 WHERE TipoNomina.CodTIpoNomina= '" & CodTipoNomina & "'"
  rs.Open "Update NomSubsidio Set [Activa] = 0 ,[Procesada] = 1 ,[Cerrada] = 1 Where (NumNomina = " & NumNomina & ")"
  rs.Open "Update DetalleSubsidio Set [Pagado] = 1 Where [NumNominaSubsidio] = " & NumNomina & " "
  
  
  
     

 
 
 ' rs.Close
  
  'rs = Nothing
  'cn = Nothing
   'dbs.Close
'actualizo el saldo del prestamo


SQlPrestamo = "SELECT Prestamo.* From Prestamo"
DtaPrestamo.RecordSource = SQlPrestamo
DtaPrestamo.Refresh


Me.DtaMovPrestamo.Refresh
If Not Me.DtaPrestamo.Recordset.EOF Then
DtaMovPrestamo.Recordset.MoveLast
End If
cantidad = DtaMovPrestamo.Recordset.RecordCount
Me.Label1.Caption = "Cerrando Prestamos"
With PBCalcNomina

.Min = 0
.Value = 0
.Max = cantidad


j = 0
TotalEmpleado = 0
If Not Me.DtaMovPrestamo.Recordset.EOF Then
DtaMovPrestamo.Recordset.MoveFirst
End If
Do While Not DtaMovPrestamo.Recordset.EOF

   
    Me.Caption = "Procesando:  " & j & " de " & cantidad & " Empleados "
    Me.LblTotal.Caption = "Procesando:  " & j & " de " & cantidad & " Empleados "
    
    
    If Me.Label1.Caption = "Cerrando Prestamos" Then
         Me.Label1.Caption = "Cerrando Prestamos."
        ElseIf Me.Label1.Caption = "Cerrando Prestamos." Then
            Me.Label1.Caption = "Cerrando Prestamos.."
         ElseIf Me.Label1.Caption = "Cerrando Prestamos.." Then
            Me.Label1.Caption = "Cerrando Prestamos..."
           ElseIf Me.Label1.Caption = "Cerrando Prestamos..." Then
             Me.Label1.Caption = "Cerrando Prestamos...."
          ElseIf Me.Label1.Caption = "Cerrando Prestamos...." Then
           Me.Label1.Caption = "Cerrando Prestamos"
          
        End If
    
    DoEvents
    If DtaMovPrestamo.Recordset("NumNomina") = NumNomina Then
     DtaPrestamo.Refresh
     Do While Not DtaPrestamo.Recordset.EOF
     ' MsgBox ((Str(DtaPrestamo.Recordset("NumPrestamo")) + " Movprestamo ") + Str(DtaMovPrestamo.Recordset.NumPrestamo))
        If DtaPrestamo.Recordset("NumPrestamo") = DtaMovPrestamo.Recordset("NumPrestamo") Then
            'DtaPrestamo.Recordset.Edit
            DtaPrestamo.Recordset("Saldo") = DtaPrestamo.Recordset("Saldo") - DtaMovPrestamo.Recordset("CuotaIgual")
            DtaPrestamo.Recordset.Update
        End If
    DtaPrestamo.Recordset.MoveNext
    Loop
  End If
  DtaMovPrestamo.Recordset.MoveNext
.Value = TotalEmpleado

If CantEmpleados <> TotalEmpleado Then
    TotalEmpleado = TotalEmpleado + 1
End If
 j = j + 1
  
 Loop
End With

'coloco en los pagos mensuales el pago de cada quien
Me.Label1.Caption = "Trasladano Registros Historicos"
With PBCalcNomina

.Min = 0
.Value = 0
.Max = CantEmpleados

j = 1
TotalEmpleado = 0


DtaEmpleados.Refresh 'ya hice el SQL Arriba
Do While Not DtaEmpleados.Recordset.EOF
   
   
    Me.Caption = "Procesando:  " & j & " de " & CantEmpleados & " Empleados "
    Me.LblTotal.Caption = "Procesando:  " & j & " de " & CantEmpleados & " Empleados "
    
    
    If Me.Label1.Caption = "Trasladano Registros Historicos" Then
         Me.Label1.Caption = "Trasladano Registros Historicos."
        ElseIf Me.Label1.Caption = "Trasladano Registros Historicos." Then
            Me.Label1.Caption = "Trasladano Registros Historicos.."
         ElseIf Me.Label1.Caption = "Trasladano Registros Historicos.." Then
            Me.Label1.Caption = "Trasladano Registros Historicos..."
           ElseIf Me.Label1.Caption = "Trasladano Registros Historicos..." Then
             Me.Label1.Caption = "Trasladano Registros Historicos...."
          ElseIf Me.Label1.Caption = "Trasladano Registros Historicos...." Then
           Me.Label1.Caption = "Trasladano Registros Historicos"
          
        End If
    
    DoEvents


    CodEmpleado = DtaEmpleados.Recordset("CodEmpleado")
    
    
  '---------------------------------------------------------------------------------------------------------------------
  '----------------------------BUSCO EL PRIMER REGISTRO DE LAS DEDUCCIONES PARA AGREGARLE UN NUMERO NOMINA -------------
  '----------------------------------------------------------------------------------------------------------------------
     Dim NumeroDeduduccion As Double
     SQlDeducciones = "SELECT MAX(Deduccion.NumDeduccion) AS NumDeduccion, TipoDeduccion.Deduccion, Deduccion.CodEmpleado, AVG(DetalleDeduccion.Valor) AS Valor, COUNT(DetalleDeduccion.NumVez) As NumVez FROM TipoDeduccion INNER JOIN Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion ON TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion Where (DetalleDeduccion.Pagado = 0) GROUP BY TipoDeduccion.Deduccion, Deduccion.CodEmpleado Having (Deduccion.CodEmpleado = " & CodEmpleado & ") ORDER BY NumDeduccion"
     DtaDeducciones.RecordSource = SQlDeducciones
     DtaDeducciones.Refresh
     Do While Not DtaDeducciones.Recordset.EOF
     
      NumeroDeduccion = DtaDeducciones.Recordset("NumDeduccion")
     
     Me.DtaConsulta.RecordSource = "SELECT DetalleDeduccion.NumNomina, Deduccion.NumDeduccion, TipoDeduccion.Deduccion, Deduccion.CodEmpleado, DetalleDeduccion.Valor, DetalleDeduccion.NumVez FROM TipoDeduccion INNER JOIN Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion ON TipoDeduccion.CodTipoDeduccion = Deduccion.CodTipoDeduccion Where (DetalleDeduccion.Pagado = 0) And (Deduccion.CodEmpleado = " & CodEmpleado & ") And (Deduccion.Numdeduccion = " & NumeroDeduccion & ") ORDER BY Deduccion.NumDeduccion, DetalleDeduccion.NumVez"
     Me.DtaConsulta.Refresh
      If Not DtaConsulta.Recordset.EOF Then
         Me.DtaConsulta.Recordset("NumNomina") = -1
         Me.DtaConsulta.Recordset.Update
      End If

      DtaDeducciones.Recordset.MoveNext
     Loop
     
     
  '---------------------------------------------------------------------------------------------------------------------
  '----------------------------BUSCO EL PRIMER REGISTRO DE LAS DEDUCCIONES PARA AGREGARLE UN NUMERO NOMINA -------------
  '----------------------------------------------------------------------------------------------------------------------
     Dim NumeroIncentivo As Double
     SQlIncentivos = "SELECT MAX(Incentivo.NumIncentivo) AS NumIncentivo, TipoIncentivo.Incentivo, Incentivo.CodEmpleado, AVG(DetalleIncentivo.Valor) AS Valor, COUNT(DetalleIncentivo.NumVez) AS NumVez, DetalleIncentivo.Pagado FROM TipoIncentivo INNER JOIN Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo ON TipoIncentivo.CodTipoIncentivo = Incentivo.CodTipoIncentivo GROUP BY TipoIncentivo.Incentivo, Incentivo.CodEmpleado, DetalleIncentivo.Pagado Having (Incentivo.CodEmpleado = " & CodEmpleado & ") And (DetalleIncentivo.Pagado = 0) "
     DtaIncentivos.RecordSource = SQlIncentivos
     DtaIncentivos.Refresh
     Do While Not DtaIncentivos.Recordset.EOF
     
      NumeroIncentivo = DtaIncentivos.Recordset("NumIncentivo")
     
     Me.DtaConsulta.RecordSource = "SELECT Incentivo.NumIncentivo, TipoIncentivo.Incentivo, Incentivo.CodEmpleado, DetalleIncentivo.Valor, DetalleIncentivo.Pagado, DetalleIncentivo.NumNomina, DetalleIncentivo.NumVez FROM TipoIncentivo INNER JOIN Incentivo INNER JOIN DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo ON TipoIncentivo.CodTipoIncentivo = Incentivo.CodTipoIncentivo Where (Incentivo.CodEmpleado = " & CodEmpleado & ") And (DetalleIncentivo.Pagado = 0) And (DetalleIncentivo.NumNomina = 0) And (Incentivo.NumIncentivo = " & NumeroIncentivo & ") ORDER BY Incentivo.NumIncentivo, DetalleIncentivo.NumVez"
     Me.DtaConsulta.Refresh
      If Not DtaConsulta.Recordset.EOF Then
         Me.DtaConsulta.Recordset("NumNomina") = -1
         Me.DtaConsulta.Recordset.Update
      End If

      DtaIncentivos.Recordset.MoveNext
     Loop


DtaEmpleados.Recordset.MoveNext

If CantEmpleados <> TotalEmpleado Then
    TotalEmpleado = TotalEmpleado + 1
 End If
 j = j + 1
 .Value = TotalEmpleado
Loop
End With
MsgBox "La Nomina Ha sido Cerrada"
DtaTipoNomina.Refresh
MousePointer = 1
Exit Sub

TipoErr:
ControlErrores


End Sub

Private Sub CmdDenominacion_Click()
On Error GoTo TipoErr
Quien = "Nomina"
CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
SQLNomina = "SELECT Nomina.* From Nomina WHERE Nomina.Activa=1 AND Nomina.CodTipoNomina= '" & CodTipoNomina & "'"
DtaNomina.RecordSource = SQLNomina
DtaNomina.Refresh

NumNomina = DtaNomina.Recordset("NumNomina")
FrmMonedas.Show 1
Exit Sub
TipoErr:
    ControlErrores
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
       NumNomina = FrmCalcularNomina.DtaNomina.Recordset("NumNomina")
       
   res = Bitacora(Now, NombreUsuario, "Calcular Nomina", "Se Exporto la Nomina BANPRO: " & NumNomina)

SQlReportes = "SELECT     Empleado.CuentaBanco, Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
SQlReportes = SQlReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
SQlReportes = SQlReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
SQlReportes = SQlReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNominaINI,Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
SQlReportes = SQlReportes & "                      Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
SQlReportes = SQlReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.Antiguedad AS TotalDevengado," & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion+ DetalleNomina.Antiguedad)" & vbLf
SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion,Empleado.CodEmpleado1" & vbLf
SQlReportes = SQlReportes & " FROM         Nomina INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       Grupo INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       Cargo INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       TipoNomina INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
SQlReportes = SQlReportes & "                       TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina" & vbLf
SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ") AND((dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
SQlReportes = SQlReportes & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia)" & vbLf
SQlReportes = SQlReportes & "                      - (dbo.DetalleNomina.Prestamo + dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.MontoIR + dbo.DetalleNomina.Deducciones) <> 0)" & vbLf
SQlReportes = SQlReportes & " ORDER BY Empleado.Nombre1" & vbLf


       Me.DtaConsulta.RecordSource = SQlReportes
       Me.DtaConsulta.Refresh

'       Me.DtaExporta.Refresh
'        'Me.'Me.DtaExporta.Recordset.Edit
'       Me.DtaExporta.Recordset("CodigoBAC") = val(Me.TxtCod.Text)
'       Me.DtaExporta.Recordset.Update

       Mes = Month(Me.DtaConsulta.Recordset("FechaNomina"))
       Año = Year(Me.DtaConsulta.Recordset("FechaNomina"))
       CanDias = Day(Me.DtaConsulta.Recordset("FechaNomina"))
       Dias = Day(Me.DtaConsulta.Recordset("FechaNomina"))
'       Cod = Me.TxtCod.Text

      ConvertirMes (Mes)
      Select Case DtaTipoNomina.Recordset("Periodo")
        Case "Quincenal"
                If CanDias > 15 Then
                   QuinLetra = "Segunda Quincena de " & Convertir
                Else
                   QuinLetra = "Primera Quincena de" & Convertir
                End If
        Case "Catorcenal los Sabados"
                
                    QuinLetra = PeriodoNominaLetras(Me.DtaConsulta.Recordset("CodTipoNomina"), Month(Me.DtaConsulta.Recordset("FechaNomina")), Year(Me.DtaConsulta.Recordset("FechaNomina")), Me.DtaConsulta.Recordset("FechaNominaINI"), Me.DtaConsulta.Recordset("FechaNomina")) & " Catorcena de " & Convertir

       End Select
  
   Case "NominaVacaciones"
      NumNomVaca = Frm13Vaca.TxtNumNomVaca.Text
      '///////////////////////////Cargo la Consulta de Vacaciones////////////////////////////////
      SQlReportes = "SELECT NomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, ([DetalleNomVaca].[SalarioMensual]*([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento])/" & DiasMes & ")-[DetalleNomVaca].[AdelantoVacaciones] AS MontoAPagar, [DetalleNomVaca].[SalarioMensual]*([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento])/" & DiasMes & " AS TotalDevengado, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, ([DetalleNomVaca].[DiasAPagar]-[DetalleNomVaca].[DiasDescuento]) AS TotalDescuento " & vbLf
      SQlReportes = SQlReportes & "FROM NomVaca INNER JOIN (Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado) ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca Where (((NomVaca.NumNomVaca) = " & NumNomVaca & " )) ORDER BY DetalleNomVaca.CodEmpleado"
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
    'Dim Heading(8) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
    'Heading(0) = "Nombre"
    'Heading(1) = "Apellidos"
    'Heading(2) = "Direccion"
    'Heading(3) = "Poblacion"
    'Heading(4) = "Provincia"
    'Heading(5) = "Pais"
    'Heading(6) = "Telefono"
    'Heading(7) = "DNI"
            
   
    Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
    'Call Formato_Excel(8, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
V = 1
H = 0
i = 1

 
  Do While Not Me.DtaConsulta.Recordset.EOF 'esto nos sirve pa leer los datos desde
       
       CodEmpleado = DtaConsulta.Recordset("CodEmpleado")
       
       '-----------------------------------------------------------------------------------------------------------
       '-------------------------BUSCO EL MONTO DE SUBSIDIO PARA SUMARLO ------------------------------------------
       '------------------------------------------------------------------------------------------------------------
       SqlString = "SELECT TOP (200) DetalleNomSubsidio.NumNominaSubsidio, DetalleNomSubsidio.CodEmpleado, Empleado.CodEmpleado1, DetalleNomSubsidio.Subsidio FROM DetalleNomSubsidio INNER JOIN Empleado ON DetalleNomSubsidio.CodEmpleado = Empleado.CodEmpleado Where (DetalleNomSubsidio.NumNominaSubsidio = " & NumNomina & ") And (DetalleNomSubsidio.CodEmpleado = " & CodEmpleado & ")"
       MDIPrimero.DtaConsulta.RecordSource = SqlString
       MDIPrimero.DtaConsulta.Refresh
       If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
         MontoSubsidio = MDIPrimero.DtaConsulta.Recordset("Subsidio")
       Else
         MontoSubsidio = 0
       End If
       
       
       If Not IsNull(DtaConsulta.Recordset("CuentaBanco")) Then
         CodigoCuenta = DtaConsulta.Recordset("CuentaBanco")
       Else
         CodigoCuenta = ""
       End If
 'la tabla de access para despues colocarlos en las celdas correspondientes
       
       Nombre = Me.DtaConsulta.Recordset("Nombres")
       Neto = Format(Me.DtaConsulta.Recordset("NetoPagar") + MontoSubsidio, "####0.00")
       Neto1 = Format(Me.DtaConsulta.Recordset("NetoPagar") + MontoSubsidio, "##,##0.00")
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
            .MoveNext

   
        End With
  Loop
  
  '/////////////////////////////SELECCION SOLO LOS EMPLEADOS QUE TIENEN SUBSIDIO Y NO TIENEN SALARIO
  Me.DtaConsulta.RecordSource = "SELECT TOP (200) DetalleNomSubsidio.id, DetalleNomSubsidio.NumNominaSubsidio, DetalleNomSubsidio.CodEmpleado, DetalleNomSubsidio.Subsidio, Empleado.CodEmpleado1, (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS Neto, Empleado.CuentaBanco, Empleado.Dolarizado, Empleado.FechaAntiguedad, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres " & _
                                "FROM DetalleNomSubsidio INNER JOIN Empleado ON DetalleNomSubsidio.CodEmpleado = Empleado.CodEmpleado INNER JOIN Nomina ON DetalleNomSubsidio.NumNominaSubsidio = Nomina.NumNomina INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado AND Nomina.NumNomina = DetalleNomina.NumNomina  " & _
                                "WHERE (DetalleNomSubsidio.NumNominaSubsidio = " & NumNomina & ") AND ((DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) = 0) AND (DetalleNomSubsidio.Subsidio <> 0) "
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

Private Sub CmdExportaCSV_Click()
On Error GoTo TipoErrs
Dim SQLExporta As String, Longitud As Integer, Respuesta As Integer
Dim Cadena As String, Mes As String, Dia As String, Ano As String
Dim TextoMonto As String, TipoMovimiento As String, j As Integer, SQlReportes As String
Dim Codigo As String
salir = False
Me.Barra.Visible = True
Me.CommonDialog1.ShowSave
Directorio = ""
Directorio = Me.CommonDialog1.FileName + ".csv"
'Fecha1 = Year(Me.DTFecha1.Value) & "-" & Month(Me.DTFecha1.Value) & "-" & Day(Me.DTFecha1.Value)
'Fecha2 = Year(Me.DTFecha2.Value) & "-" & Month(Me.DTFecha2.Value) & "-" & Day(Me.DTFecha2.Value)
NumNomina = DtaNomina.Recordset("NumNomina")

SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
SQlReportes = SQlReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
SQlReportes = SQlReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
SQlReportes = SQlReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
SQlReportes = SQlReportes & "                      Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
SQlReportes = SQlReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion AS TotalDevengado," & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion)" & vbLf
SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion,Empleado.CodEmpleado1" & vbLf
SQlReportes = SQlReportes & " FROM         Nomina INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       Grupo INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       Cargo INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       TipoNomina INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
SQlReportes = SQlReportes & "                       TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina" & vbLf
SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ")" & vbLf
SQlReportes = SQlReportes & " ORDER BY Nomina.NumNomina, Empleado.CodEmpleado1" & vbLf
   

Me.AdoBusca.RecordSource = SQlReportes
AdoBusca.Refresh
Me.AdoBusca.Recordset.MoveLast
Maximo = AdoBusca.Recordset.RecordCount
If (Dir(Directorio) <> "") Then
  Respuesta = MsgBox("Reescribir el Archivo?", vbYesNo, "Enlace Pacioli")
  If Respuesta = 6 Then
     Kill (Directorio)
               Open Directorio For Output As #1
                     
                AdoBusca.Recordset.MoveFirst
                With Barra
                   .Min = 0
                   .Value = 0
                   .Max = Maximo
                   j = 0
                   
                Cadena = "CodEmpleado" & "," & "Nombres" & "," & "SalarioBasico" & "," & "Destajo" & "," & "HorasExtras" & "," & "Comisiones" & "," & "Incentivos" & "," & "VacacionesPagadas" & "," & "SeptimoDia" & "," & "IncetivoProduccion" & "," & "OtrosIngresos" & "," & "TotalDevengado" & "," & "Prestamo" & "," & "MontoINSS" & "," & "MontoIr" & "," & "Deducciones" & "," & "INSSPatronal"
                Print #1, Cadena
                
                 Do While Not AdoBusca.Recordset.EOF
                 '////////Inicialiso las variables/////////////////
 
                      Cadena = AdoBusca.Recordset("CodEmpleado1") & "," & AdoBusca.Recordset("Nombres") & "," & AdoBusca.Recordset("SalarioBasico") & "," & AdoBusca.Recordset("Destajo") & "," & AdoBusca.Recordset("HorasExtras") & "," & AdoBusca.Recordset("Comisiones") & "," & AdoBusca.Recordset("Incentivos") & "," & AdoBusca.Recordset("VacacionesPagadas") & "," & AdoBusca.Recordset("SeptimoDia") & "," & AdoBusca.Recordset("IncetivoProduccion") & "," & AdoBusca.Recordset("OtrosIngresos") & "," & AdoBusca.Recordset("TotalDevengado") & "," & AdoBusca.Recordset("Prestamo") & "," & AdoBusca.Recordset("MontoINSS") & "," & AdoBusca.Recordset("MontoIr") & "," & AdoBusca.Recordset("Deducciones") & "," & AdoBusca.Recordset("INSSPatronal")

                    Print #1, Cadena
                                    
                    
                    
                  AdoBusca.Recordset.MoveNext
                  j = j + 1
                  Me.Caption = "Procesando:  " & j & " de " & Maximo & " Registros "
                  DoEvents
                  .Value = j
                  Cadena = ""
                  Loop
                  End With
                  
                 Close #1

                MsgBox "La Exportacion, fue Creada con Exito", vbExclamation, "Sistema de Enlace"
                salir = True
  End If
Else '//////En caso que no exista el Archivo///////////
                
                Open Directorio For Output As #1
                'SQLExporta = "SELECT Empleado.CodEmpleado, Empleado.CodDepartamento, Historico.CodCuenta, Historico.CuentaCredito, DetalleNomina.NumNomina, Nomina.Fecha, [DetalleNomina]![SalarioBasico]+[DetalleNomina]![Destajo]+[DetalleNomina]![HorasExtras]+[DetalleNomina]![Comisiones]+[DetalleNomina]![Incentivos]-[DetalleNomina]![Deducciones]-[DetalleNomina]![Prestamo]-[DetalleNomina]![MontoINSS]-[DetalleNomina]![MontoIR]+[DetalleNomina]![TotalSubsidio] AS GranTotal FROM Nomina INNER JOIN ((Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado) ON Nomina.NumNomina = DetalleNomina.NumNomina Where DetalleNomina.NumNomina = " & NumNomina & " ORDER BY Empleado.CodEmpleado"
                
                AdoBusca.Recordset.MoveFirst
                With Barra
                   .Min = 0
                   .Value = 0
                   .Max = Maximo
                   j = 0
                 Do While Not AdoBusca.Recordset.EOF

                      Cadena = AdoBusca.Recordset("CodEmpleado1") & "," & AdoBusca.Recordset("Nombres") & "," & AdoBusca.Recordset("SalarioBasico") & "," & AdoBusca.Recordset("Destajo") & "," & AdoBusca.Recordset("HorasExtras") & "," & AdoBusca.Recordset("Comisiones") & "," & AdoBusca.Recordset("Incentivos") & "," & AdoBusca.Recordset("VacacionesPagadas") & "," & AdoBusca.Recordset("SeptimoDia") & "," & AdoBusca.Recordset("IncetivoProduccion") & "," & AdoBusca.Recordset("OtrosIngresos") & "," & AdoBusca.Recordset("TotalDevengado") & "," & AdoBusca.Recordset("Prestamo") & "," & AdoBusca.Recordset("MontoINSS") & "," & AdoBusca.Recordset("MontoIr") & "," & AdoBusca.Recordset("Deducciones") & "," & AdoBusca.Recordset("INSSPatronal")



                    Print #1, Cadena
                                    
                    
                    
                  AdoBusca.Recordset.MoveNext
                  j = j + 1
                  .Value = j
                  Me.Caption = "Procesando:  " & j & " de " & Maximo & " Registros "
                  DoEvents
                  Cadena = ""
                  Loop
                  End With
                  
                 Close #1

                MsgBox "La Exportacion, fue Creada con Exito", vbExclamation, "Sistema de Nominas"
                Me.Barra.Visible = False
  End If
Exit Sub
TipoErrs:
  MsgBox Err.Description
End Sub

Private Sub CmdExportar_Click()
Quien = "CalcularNomina"
FrmExportaBac.Show 1
End Sub

Private Sub CmdMonedasDpto_Click()
On Error GoTo TipoErr
Dim CodDepartamento As String, DescripcionDpto As String, SQlReportes As String
Quien = "MonedasDepartamento"
CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
SQLNomina = "SELECT Nomina.* From Nomina WHERE Nomina.Activa=1 AND Nomina.CodTipoNomina= '" & CodTipoNomina & "'"
DtaNomina.RecordSource = SQLNomina
DtaNomina.Refresh

NumNomina = DtaNomina.Recordset("NumNomina")

 res = Bitacora(Now, NombreUsuario, "Calcular Nomina", "Se Exporto la Nomina BAC: " & NumNomina)

     Me.AdoDepartamento.Refresh
     Do While Not Me.AdoDepartamento.Recordset.EOF
     
          CodigoDepartamento = Me.AdoDepartamento.Recordset("CodDepartamento")
          DescripcionDpto = Me.AdoDepartamento.Recordset("Departamento")
     
          SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal,  Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo,  Cargo.Cargo, DetalleNomina.BonoProduccion, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones,  " & _
                      "DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal,  DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo,  Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE, DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion  AS TotalDevengado, DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir,  " & _
                      "(DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +  DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar, DetalleNomina.TarifaHoraria, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1, departamento.departamento , departamento.CodDepartamento,Nomina.FechaNominaINI  " & _
                      "FROM  Nomina INNER JOIN  Grupo INNER JOIN  Cargo INNER JOIN  TipoNomina INNER JOIN  Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN  DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON  TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN  Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento " & _
                      "WHERE     (Nomina.NumNomina = " & NumNomina & ") AND ((DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) <> 0) AND (Departamento.CodDepartamento = '" & CodigoDepartamento & "') " & _
                      "ORDER BY Empleado.Nombre1, Empleado.CodEmpleado "
                      
                      Me.AdoBusca.RecordSource = SQlReportes
                      Me.AdoBusca.Refresh
                      If Not Me.AdoBusca.Recordset.EOF Then
                      
                        MsgBox "Departamento Procesado " & DescripcionDpto
                        FrmMonedas.Show 1
                      End If

         Me.AdoDepartamento.Recordset.MoveNext
      Loop


Exit Sub
TipoErr:
    ControlErrores
End Sub

Private Sub CmdPrnNomina_Click()
Dim FormatoColilla As String


Dim rpt As Object
Dim fPreview As New FrmPreview

' Set rpt = New Arep
'On Error GoTo TipoErr
Dim FechaIni As Variant, FechaFin As Variant
Dim Espacio As String


CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
'SQLNomina = "SELECT Nomina.* From Nomina WHERE Nomina.Activa=True AND Nomina.CodTipoNomina= '" & CodTipoNomina & "'"
'DtaNomina.RecordSource = SQLNomina
'DtaNomina.Refresh
Espacio = " "
NumNomina = DtaNomina.Recordset("NumNomina")

res = Bitacora(Now, NombreUsuario, "Calcular Nomina", "Se imprimio Colillas de la nomina: " & NumNomina)

DtaNominas.Refresh
Do While Not DtaNominas.Recordset.EOF
   If Me.DtaNominas.Recordset("CodTipoNomina") = Me.DtaTipoNomina.Recordset("CodTipoNomina") And DtaNominas.Recordset("Activa") = True Then
      FechaIni = Format(DtaNominas.Recordset("FechaNominaINI"), "dd/mm/yyyy")
      FechaFin = Format(DtaNominas.Recordset("FechaNomina"), "dd/mm/yyyy")
   End If
   DtaNominas.Recordset.MoveNext
Loop

'///////////////////////////INTRUCCION SQL SERVER
'SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
'SQlReportes = SQlReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
'SQlReportes = SQlReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
'SQlReportes = SQlReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
'SQlReportes = SQlReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
'SQlReportes = SQlReportes & "                      Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
'SQlReportes = SQlReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
'SQlReportes = SQlReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.BonoProduccion, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
'SQlReportes = SQlReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
'SQlReportes = SQlReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
'SQlReportes = SQlReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
'SQlReportes = SQlReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
'SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion AS TotalDevengado," & vbLf
'SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
'SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
'SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion)" & vbLf
'SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
'SQlReportes = SQlReportes & "                      DetalleNomina.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion,Empleado.CodEmpleado1,Empleado.NumeroInss, Empleado.NumCedula " & vbLf
'SQlReportes = SQlReportes & " FROM         Nomina INNER JOIN" & vbLf
'SQlReportes = SQlReportes & "                       Grupo INNER JOIN" & vbLf
'SQlReportes = SQlReportes & "                       Cargo INNER JOIN" & vbLf
'SQlReportes = SQlReportes & "                       TipoNomina INNER JOIN" & vbLf
'SQlReportes = SQlReportes & "                       Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
'SQlReportes = SQlReportes & "                       DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
'SQlReportes = SQlReportes & "                       TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina" & vbLf
'SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ") AND((dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
'SQlReportes = SQlReportes & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia)" & vbLf
'SQlReportes = SQlReportes & "                      - (dbo.DetalleNomina.Prestamo + dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.MontoIR + dbo.DetalleNomina.Deducciones) <> 0)" & vbLf
'SQlReportes = SQlReportes & " ORDER BY Nomina.NumNomina, Empleado.CodEmpleado1" & vbLf

SQlReportes = "SELECT Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS,Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13,Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2,Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo, Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo,DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos,DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas, DetalleNomina.BonoProduccion, DetalleNomina.Prestamo, " & _
              "DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre,DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo,Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres,DetalleNomina.HE,DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasTurno + DetalleNomina.Antiguedad + DetalleNomina.Reembolso AS TotalDevengado,DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir, " & _
              "(DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasTurno + DetalleNomina.Antiguedad + DetalleNomina.Reembolso) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar, DetalleNomina.TarifaHoraria, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1, Empleado.NumeroInss, Empleado.NumCedula, departamento.departamento,DetalleNomina.HorasTurno,DetalleNomina.HTurno,DetalleNomina.Antiguedad,DetalleNomina.AñoAntiguedad, DetalleNomina.Reembolso " & _
              "FROM  Nomina INNER JOIN Grupo INNER JOIN Cargo INNER JOIN TipoNomina INNER JOIN Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON " & _
              "TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento  " & _
              "WHERE (Nomina.NumNomina = " & NumNomina & ") AND (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasTurno + DetalleNomina.Antiguedad <> 0) ORDER BY Nomina.NumNomina, Empleado.Apellido1, Empleado.Apellido2, Empleado.Nombre1, Empleado.Nombre2"  ', Empleado.CodEmpleado1

MDIPrimero.DtaEmpresa.Refresh
If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
  FormatoColilla = MDIPrimero.DtaEmpresa.Recordset("FormatoColilla")
End If

Select Case FormatoColilla
  Case "Colilla Trides"
    ArepColillasPago4.AdoColillas.Source = SQlReportes
    ArepColillasPago4.LblPeriodo.Caption = "   Colilla de Pago # " & NumNomina & ", Corespondiente del " & FrmCalcularNomina.LblFecha1.Caption & " al " & FrmCalcularNomina.LblFecha2.Caption
    ArepColillasPago4.lbltitulo.Caption = Titulo
    ArepColillasPago4.AdoColillas.ConnectionString = ConexionReporte
    ArepColillasPago4.Show 1
 
  
  Case "Colilla Horas Turno"
    ArepColillasPagoTurno.AdoColillas.Source = SQlReportes
    ArepColillasPago2.LblPeriodo.Caption = "   Colilla de Pago # " & NumNomina & ", Corespondiente del " & FrmCalcularNomina.LblFecha1.Caption & " al " & FrmCalcularNomina.LblFecha2.Caption
    ArepColillasPagoTurno.lbltitulo.Caption = Titulo
    ArepColillasPagoTurno.AdoColillas.ConnectionString = ConexionReporte
    ArepColillasPagoTurno.Show 1

  Case "Colilla Comercial3"
  
    ArepColillasPago3.AdoColillas.Source = SQlReportes
    ArepColillasPago3.LblPeriodo.Caption = FechaIni & " al " & FechaFin
    ArepColillasPago3.lbltitulo.Caption = Titulo
    ArepColillasPago3.AdoColillas.ConnectionString = ConexionReporte
    ArepColillasPago3.Show 1
  
  Case "Colilla Bono Produccion"
  
'     Set rpt = New ArepColillasBono
'     rpt.AdoColillas.ConnectionString = ConexionReporte
'     rpt.AdoColillas.Source = SQlReportes
'     fPreview.RunReport rpt
'     fPreview.Show 1

    ArepColillasBono.AdoColillas.Source = SQlReportes
    ArepColillasBono.LblPeriodo.Caption = FechaIni & " al " & FechaFin
    ArepColillasBono.lbltitulo.Caption = Titulo
    ArepColillasBono.AdoColillas.ConnectionString = ConexionReporte
    ArepColillasBono.Show 1

  Case "Colilla Comercial"
    ArepColillasPago.AdoColillas.Source = SQlReportes
    ArepColillasPago.LblPeriodo.Caption = FechaIni & " al " & FechaFin
    PeriodoReporte = "Desde " & FechaIni & " Hasta " & FechaFin
    ArepColillasPago.lbltitulo.Caption = Titulo
    ArepColillasPago.AdoColillas.ConnectionString = ConexionReporte
    ArepColillasPago.Show 1
    
  Case "Colilla Comercial2"
    ArepColillasPago2.AdoColillas.Source = SQlReportes
    ArepColillasPago2.LblPeriodo.Caption = "   Colilla de Pago # " & NumNomina & ", Corespondiente del " & FrmCalcularNomina.LblFecha1.Caption & " al " & FrmCalcularNomina.LblFecha2.Caption
    ArepColillasPago2.lbltitulo.Caption = Titulo
    ArepColillasPago2.AdoColillas.ConnectionString = ConexionReporte
    ArepColillasPago2.Show 1
     
  
  Case "Colilla Produccion"

    ArepColillas.AdoColillas.Source = SQlReportes
    ArepColillas.LblPeriodo.Caption = FechaIni & " al " & FechaFin
    ArepColillas.lbltitulo.Caption = Titulo
    ArepColillas.AdoColillas.ConnectionString = ConexionReporte
    ArepColillas.Show 1
    
   Case "Colilla Produccion Tamaño Legal"

    ArepColillaProduccionLegal.AdoColillas.Source = SQlReportes
    ArepColillaProduccionLegal.LblPeriodo.Caption = FechaIni & " al " & FechaFin
    ArepColillaProduccionLegal.lbltitulo.Caption = Titulo
    ArepColillaProduccionLegal.AdoColillas.ConnectionString = ConexionReporte
    ArepColillaProduccionLegal.Show 1

End Select


Exit Sub
TipoErr:
    ControlErrores

End Sub

Private Sub CmdprNomina_Click()
Dim rpt As Object
Dim fPreview As New FrmPreview, Dias As Double
'On Error GoTo TipoErr
Dim Moneda As String, FormatoNomina As String
Set rpt = New ArepNominaProduccionLegal

CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")

Me.AdoBusca.RecordSource = "SELECT CodTipoNomina, Nomina, Periodo, UltFecha, TipoPago, Moneda, MantValor, Activa, PorcientoInss, TasaInss, PorcientoIr, TasaIr,TasaInssPatronal From TipoNomina WHERE  (CodTipoNomina = '" & CodTipoNomina & "')"
Me.AdoBusca.Refresh
If Not Me.AdoBusca.Recordset.EOF Then
  Moneda = Me.AdoBusca.Recordset("Moneda")
  
  Select Case Me.AdoBusca.Recordset("Periodo")
    Case "Semanal Viernes": Dias = 28
    Case "Semanal Sabado": Dias = 28
    Case "Catorcenal los Viernes": Dias = 14
    Case "Catorcenal los Sabados": Dias = 14
    Case "Quincenal": Dias = 15
    Case "Mensual": Dias = 30
    Case "Trimestral": Dias = 90
    Case "Semestral": Dias = 180
    
  End Select
  
  
End If

NumNomina = DtaNomina.Recordset("NumNomina")
NumeroNominas = DtaNomina.Recordset("NumNomina")

res = Bitacora(Now, NombreUsuario, "Calcular Nomina", "Se imprimio Nomina de la nomina: " & NumNomina)
'/

'SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
'SQlReportes = SQlReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
'SQlReportes = SQlReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
'SQlReportes = SQlReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
'SQlReportes = SQlReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
'SQlReportes = SQlReportes & "                      Cargo.Cargo, DetalleNomina.BonoProduccion, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
'SQlReportes = SQlReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
'SQlReportes = SQlReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
'SQlReportes = SQlReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
'SQlReportes = SQlReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
'SQlReportes = SQlReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
'SQlReportes = SQlReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
'SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad + DetalleNomina.HorasTurno AS TotalDevengado," & vbLf
'SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
'SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
'SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad + DetalleNomina.HorasTurno)" & vbLf
'SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
'SQlReportes = SQlReportes & "                      DetalleNomina.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion,Empleado.CodEmpleado1, Empleado.NumeroInss,DetalleNomina.AjusteINSS, Empleado.NumCedula,DetalleNomina.Antiguedad,DetalleNomina.AñoAntiguedad, Empleado.SueldoPeriodo * 2 AS SalarioMensualHM, Empleado.SueldoPeriodo * 2 / 30 AS SalarioDiaHM, 15 AS D, DetalleNomina.DiasVacaciones AS DV,   Empleado.DiasBasico AS DD, 15 - DetalleNomina.DiasVacaciones - Empleado.DiasBasico AS DL,  15 - DetalleNomina.DiasVacaciones - Empleado.DiasBasico + DetalleNomina.DiasVacaciones + DetalleNomina.DiasAdicionales AS T, (Nombre1 + ' ' + Nombre2 + ' ' + Apellido1 + ' ' + Apellido2) as NombreCompleto , Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30) AS VDD, Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30) AS DiasLab" & vbLf
'SQlReportes = SQlReportes & ",  Empleado.SueldoPeriodo * 2 / 30 / 8 * 2 AS ValorHE,    Empleado.SueldoPeriodo * 2 / 30 / 8 * 2 * DetalleNomina.HE AS TotalHE,  (Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30)) + Empleado.SueldoPeriodo * 2 / 30 / 8 * 2 * DetalleNomina.HE AS TotalPagar, DetalleNomina.Deducciones as OtrasDeduciones , (Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30)) * 0.02 AS Inatec, DetalleNomina.DiasAdicionales as DA, DetalleNomina.ValorDiasAdicionales as VDA" & vbLf
'SQlReportes = SQlReportes & " FROM         Nomina INNER JOIN" & vbLf
'SQlReportes = SQlReportes & "                       Grupo INNER JOIN" & vbLf
'SQlReportes = SQlReportes & "                       Cargo INNER JOIN" & vbLf
'SQlReportes = SQlReportes & "                       TipoNomina INNER JOIN" & vbLf
'SQlReportes = SQlReportes & "                       Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
'SQlReportes = SQlReportes & "                       DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
'SQlReportes = SQlReportes & "                       TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina" & vbLf
'SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ") AND((dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
'SQlReportes = SQlReportes & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia)" & vbLf
'SQlReportes = SQlReportes & "                      - (dbo.DetalleNomina.Prestamo + dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.MontoIR + dbo.DetalleNomina.Deducciones) <> 0)" & vbLf
'SQlReportes = SQlReportes & " ORDER BY Empleado.CodGrupo, Empleado.Nombre1" & vbLf  'Empleado.CodEmpleado1


'SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones,   Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones,       Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado,    Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo, Cargo.Cargo, DetalleNomina.BonoProduccion, DetalleNomina.SalarioBasico,    DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones,    DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, "
'SQlReportes = SQlReportes & " DetalleNomina.MontoIR,  DetalleNomina.Vacaciones,    DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo,     Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE,       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas     + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad + DetalleNomina.HorasTurno AS TotalDevengado,     DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir,    (DetalleNomina.SalarioBasico "
'SQlReportes = SQlReportes & " + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas   + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad + DetalleNomina.HorasTurno)      - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar, DetalleNomina.TarifaHoraria, DetalleNomina.HTrabajada,    DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1, Empleado.NumeroInss, DetalleNomina.AjusteINSS, Empleado.NumCedula, DetalleNomina.Antiguedad,     DetalleNomina.AñoAntiguedad, Empleado.SueldoPeriodo * 2 AS SalarioMensualHM, Empleado.SueldoPeriodo * 2 / 30 AS SalarioDiaHM, 15 AS D, DetalleNomina.DiasVacaciones AS DV,     Empleado.DiasBasico AS DD, 15 - DetalleNomina.DiasVacaciones - Empleado.DiasBasico AS DL, "
'SQlReportes = SQlReportes & "    15 - DetalleNomina.DiasVacaciones - Empleado.DiasBasico + DetalleNomina.DiasVacaciones + DetalleNomina.DiasAdicionales AS T,     Empleado.Nombre1 + ' ' + Empleado.Nombre2 +' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS NombreCompleto, Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30) AS VDD,      Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30) AS DiasLab, Empleado.SueldoPeriodo * 2 / 30 / 8 * 2 AS ValorHE, Empleado.SueldoPeriodo * 2 / 30 / 8 * 2 * DetalleNomina.HE AS TotalHE, (Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30))    + Empleado.SueldoPeriodo * 2 / 30 / 8 * 2 * DetalleNomina.HE AS TotalPagar, DetalleNomina.Deducciones AS OtrasDeduciones,     (Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30)) * 0.02 AS Inatec,  DetalleNomina.DiasAdicionales AS DA, DetalleNomina.ValorDiasAdicionales AS VDA, Empleado.SueldoPeriodo * 2 AS SalarioMensualProduccion, "
'SQlReportes = SQlReportes & "        Departamento.Departamento  FROM         Nomina INNER JOIN    Grupo INNER JOIN      Cargo INNER JOIN    TipoNomina INNER JOIN   Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN    DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND       Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN   Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento"
'SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ") AND((dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
'SQlReportes = SQlReportes & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia)" & vbLf
'SQlReportes = SQlReportes & "                      - (dbo.DetalleNomina.Prestamo + dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.MontoIR + dbo.DetalleNomina.Deducciones) <> 0)" & vbLf
'SQlReportes = SQlReportes & " ORDER BY  Empleado.CodGrupo, Empleado.Nombre1" & vbLf  'Empleado.CodEmpleado1

 If Me.AdoBusca.Recordset("Periodo") = "Catorcenal los Sabados" Then

             SQlReportes = "SELECT     Nomina.NumNomina,DetalleNomina.Reembolso, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones,"
             SQlReportes = SQlReportes & "         Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones,"
             SQlReportes = SQlReportes & "         Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado,"
             SQlReportes = SQlReportes & "         Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo, Cargo.Cargo, DetalleNomina.BonoProduccion, DetalleNomina.SalarioBasico,"
             SQlReportes = SQlReportes & "         DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones,"
             SQlReportes = SQlReportes & "        DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones,"
             SQlReportes = SQlReportes & "         DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo,"
             SQlReportes = SQlReportes & "         Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE,"
             SQlReportes = SQlReportes & "         DetalleNomina.SalarioBasico, DetalleNomina.SalarioBasico, DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas"
             SQlReportes = SQlReportes & "          + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad + DetalleNomina.HorasTurno + DetalleNomina.SalarioBasico + DetalleNomina.Reembolso AS TotalDevengado,"
             SQlReportes = SQlReportes & "         DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir,"
             SQlReportes = SQlReportes & "         (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas"
             SQlReportes = SQlReportes & "          + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad + DetalleNomina.HorasTurno)"
             SQlReportes = SQlReportes & "         - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar, DetalleNomina.TarifaHoraria, DetalleNomina.HTrabajada,"
             SQlReportes = SQlReportes & "         DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1, Empleado.NumeroInss, DetalleNomina.AjusteINSS, Empleado.NumCedula, DetalleNomina.Antiguedad,"
             SQlReportes = SQlReportes & "         DetalleNomina.AñoAntiguedad, CASE WHEN Empleado.SueldoPeriodo = 0 THEN Empleado.TarifaHoraria * 8 * 28 ELSE Empleado.SueldoPeriodo * 2 END AS SalarioMensualHM, CASE WHEN Empleado.SueldoPeriodo = 0 THEN Empleado.TarifaHoraria * 8  ELSE Empleado.SueldoPeriodo * 2 / 28 END AS SalarioDiaHM, " & Dias & " AS D, DetalleNomina.DiasVacaciones AS DV,"
             SQlReportes = SQlReportes & "         Empleado.DiasBasico AS DD, " & Dias & " - DetalleNomina.DiasVacaciones - Empleado.DiasBasico AS DL,"
             SQlReportes = SQlReportes & "         " & Dias & " - DetalleNomina.DiasVacaciones - Empleado.DiasBasico + DetalleNomina.DiasVacaciones + DetalleNomina.DiasAdicionales AS T,"
             SQlReportes = SQlReportes & "         Empleado.Apellido1 + ' ' + Empleado.Apellido2 + ' ' + Empleado.Nombre1 + ' ' + Empleado.Nombre2 AS NombreCompleto, Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30) AS VDD,"
             SQlReportes = SQlReportes & "         Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30) AS DiasLab, Empleado.SueldoPeriodo * 2 / 30 / 8 * 2 AS ValorHE,"
             SQlReportes = SQlReportes & "         CASE WHEN Empleado.SueldoPeriodo = 0 THEN Empleado.TarifaHoraria * 2 * DetalleNomina.HE ELSE Empleado.SueldoPeriodo * 2 / 28 / 8 * 2 * DetalleNomina.HE END AS TotalHE, (Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30))"
             SQlReportes = SQlReportes & "         + Empleado.SueldoPeriodo * 2 / 28 / 8 * 2 * DetalleNomina.HE AS TotalPagar, DetalleNomina.Deducciones AS OtrasDeduciones,"
             SQlReportes = SQlReportes & "         (Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30)) * 0.02 AS Inatec, DetalleNomina.DiasAdicionales AS DA, DetalleNomina.ValorDiasAdicionales AS VDA,"
             SQlReportes = SQlReportes & "         CASE WHEN Empleado.SueldoPeriodo = 0 THEN Empleado.TarifaHoraria * 8 * 28  ELSE Empleado.SueldoPeriodo  * 2 END AS SalarioMensualProduccion, Departamento.Departamento, Historico.FechaContrato, (SELECT     SUM(DetalleDeduccion.Valor) AS Valor"
             SQlReportes = SQlReportes & "              FROM          DetalleDeduccion INNER JOIN"
             SQlReportes = SQlReportes & "                                    Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion"
             SQlReportes = SQlReportes & "             WHERE      (Deduccion.CodEmpleado = Empleado.CodEmpleado) AND (Deduccion.NUmNomina = Nomina.NumNomina) AND (NOT (Deduccion.CodTipoDeduccion = '02'))"
             SQlReportes = SQlReportes & "               GROUP BY Deduccion.CodEmpleado, Deduccion.NUmNomina) AS OtrasDeduccionesHM"
             SQlReportes = SQlReportes & "         FROM         Nomina INNER JOIN"
             SQlReportes = SQlReportes & "         Grupo INNER JOIN"
             SQlReportes = SQlReportes & "         Cargo INNER JOIN"
             SQlReportes = SQlReportes & "         TipoNomina INNER JOIN"
             SQlReportes = SQlReportes & "         Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN"
             SQlReportes = SQlReportes & "         DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND"
             SQlReportes = SQlReportes & "         Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN"
             SQlReportes = SQlReportes & "         Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN"
             SQlReportes = SQlReportes & "         Historico ON Empleado.CodEmpleado = Historico.Codempleado"
             SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ") AND((dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
             SQlReportes = SQlReportes & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia)" & vbLf
             SQlReportes = SQlReportes & "                      - (dbo.DetalleNomina.Prestamo + dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.MontoIR + dbo.DetalleNomina.Deducciones) <> 0)" & vbLf
             SQlReportes = SQlReportes & " ORDER BY  Departamento.Departamento, Empleado.CodGrupo, Empleado.Apellido1, Empleado.Apellido2" & vbLf  'Empleado.CodEmpleado1
  Else
             SQlReportes = "SELECT     Nomina.NumNomina,DetalleNomina.Reembolso, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones,"
             SQlReportes = SQlReportes & "         Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones,"
             SQlReportes = SQlReportes & "         Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado,"
             SQlReportes = SQlReportes & "         Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo, Cargo.Cargo, DetalleNomina.BonoProduccion, DetalleNomina.SalarioBasico,"
             SQlReportes = SQlReportes & "         DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones,"
             SQlReportes = SQlReportes & "        DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones,"
             SQlReportes = SQlReportes & "         DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo,"
             SQlReportes = SQlReportes & "         Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE,"
             SQlReportes = SQlReportes & "         DetalleNomina.SalarioBasico, DetalleNomina.SalarioBasico, DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas"
             SQlReportes = SQlReportes & "          + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad + DetalleNomina.HorasTurno + DetalleNomina.SalarioBasico + DetalleNomina.Reembolso AS TotalDevengado,"
             SQlReportes = SQlReportes & "         DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir,"
             SQlReportes = SQlReportes & "         (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas"
             SQlReportes = SQlReportes & "          + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad + DetalleNomina.HorasTurno)"
             SQlReportes = SQlReportes & "         - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar, DetalleNomina.TarifaHoraria, DetalleNomina.HTrabajada,"
             SQlReportes = SQlReportes & "         DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1, Empleado.NumeroInss, DetalleNomina.AjusteINSS, Empleado.NumCedula, DetalleNomina.Antiguedad,"
             SQlReportes = SQlReportes & "         DetalleNomina.AñoAntiguedad, Empleado.SueldoPeriodo * 2 AS SalarioMensualHM, Empleado.SueldoPeriodo * 2 / 30 AS SalarioDiaHM, " & Dias & " AS D, DetalleNomina.DiasVacaciones AS DV,"
             SQlReportes = SQlReportes & "         Empleado.DiasBasico AS DD, " & Dias & " - DetalleNomina.DiasVacaciones - Empleado.DiasBasico AS DL,"
             SQlReportes = SQlReportes & "         " & Dias & " - DetalleNomina.DiasVacaciones - Empleado.DiasBasico + DetalleNomina.DiasVacaciones + DetalleNomina.DiasAdicionales AS T,"
             SQlReportes = SQlReportes & "         Empleado.Apellido1 + ' ' + Empleado.Apellido2 + ' ' + Empleado.Nombre1 + ' ' + Empleado.Nombre2 AS NombreCompleto, Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30) AS VDD,"
             SQlReportes = SQlReportes & "         Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30) AS DiasLab, Empleado.SueldoPeriodo * 2 / 30 / 8 * 2 AS ValorHE,"
             SQlReportes = SQlReportes & "         Empleado.SueldoPeriodo * 2 / 30 / 8 * 2 * DetalleNomina.HE AS TotalHE, (Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30))"
             SQlReportes = SQlReportes & "         + Empleado.SueldoPeriodo * 2 / 30 / 8 * 2 * DetalleNomina.HE AS TotalPagar, DetalleNomina.Deducciones AS OtrasDeduciones,"
             SQlReportes = SQlReportes & "         (Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30)) * 0.02 AS Inatec, DetalleNomina.DiasAdicionales AS DA, DetalleNomina.ValorDiasAdicionales AS VDA,"
             SQlReportes = SQlReportes & "         Empleado.SueldoPeriodo * 2 AS SalarioMensualProduccion, Departamento.Departamento, Historico.FechaContrato, (SELECT     SUM(DetalleDeduccion.Valor) AS Valor"
             SQlReportes = SQlReportes & "              FROM          DetalleDeduccion INNER JOIN"
             SQlReportes = SQlReportes & "                                    Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion"
             SQlReportes = SQlReportes & "             WHERE      (Deduccion.CodEmpleado = Empleado.CodEmpleado) AND (Deduccion.NUmNomina = Nomina.NumNomina) AND (NOT (Deduccion.CodTipoDeduccion = '02'))"
             SQlReportes = SQlReportes & "               GROUP BY Deduccion.CodEmpleado, Deduccion.NUmNomina) AS OtrasDeduccionesHM"
             SQlReportes = SQlReportes & "         FROM         Nomina INNER JOIN"
             SQlReportes = SQlReportes & "         Grupo INNER JOIN"
             SQlReportes = SQlReportes & "         Cargo INNER JOIN"
             SQlReportes = SQlReportes & "         TipoNomina INNER JOIN"
             SQlReportes = SQlReportes & "         Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN"
             SQlReportes = SQlReportes & "         DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND"
             SQlReportes = SQlReportes & "         Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN"
             SQlReportes = SQlReportes & "         Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN"
             SQlReportes = SQlReportes & "         Historico ON Empleado.CodEmpleado = Historico.Codempleado"
             SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ") AND((dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
             SQlReportes = SQlReportes & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia)" & vbLf
             SQlReportes = SQlReportes & "                      - (dbo.DetalleNomina.Prestamo + dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.MontoIR + dbo.DetalleNomina.Deducciones) <> 0)" & vbLf
             SQlReportes = SQlReportes & " ORDER BY  Departamento.Departamento, Empleado.CodGrupo, Empleado.Apellido1, Empleado.Apellido2" & vbLf
  End If
'TotalDevengado

If Moneda = "US" Then
 ArepNominaDolares.AdoNomina.Source = SQlReportes
 ArepNominaDolares.lbltitulo.Caption = Titulo
 ArepNominaDolares.LblSubtitulo.Caption = SubTitulo
 If Dir(RutaLogo) <> "" Then
 ArepNominaDolares.ImgLogo.Picture = LoadPicture(RutaLogo)
 End If
 ArepNominaDolares.AdoNomina.ConnectionString = ConexionReporte
 ArepNominaDolares.LblFecha.Caption = Format(Now, "dddddd")
 ArepNominaDolares.LblDesde = Me.LblFecha1.Caption
 ArepNominaDolares.LblHasta = Me.LblFecha2.Caption
 ArepNominaDolares.Show 1

'           fPreview.arv.ReportSource = ArepNominaDolares
'           fPreview.Show 1

Else

 MDIPrimero.DtaEmpresa.Refresh
 If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
   FormatoNomina = MDIPrimero.DtaEmpresa.Recordset("FormatoNomina")
 End If
 
 Select Case FormatoNomina
 
 Case "Nomina Destajo"
   Set rpt = New ArepNominaDestajo
   
    rpt.lbltitulo.Caption = Titulo
    If Dir(RutaLogo, vbDirectory) <> "" Then
        rpt.ImgLogo.Picture = LoadPicture(RutaLogo)
    End If
    
    
    rpt.NumeroNomina = NumNomina
    rpt.LblDesde.Caption = "Desde: " & Me.LblFecha1.Caption & "   Hasta: " & Me.LblFecha2.Caption
    rpt.AdoNomina.Source = SQlReportes
      
        rpt.AdoNomina.ConnectionString = ConexionReporte
'        ArepNominaProduccionLegal.Show 1
           fPreview.arv.ReportSource = rpt
           fPreview.Show 1
 
 
 Case "Nomina Hanter Metal"
   Set rpt = New ArepNominasHM
    rpt.lbltitulo.Caption = Titulo
    If Dir(RutaLogo) <> "" Then
        rpt.ImgLogo.Picture = LoadPicture(RutaLogo)
    End If
    
    
    rpt.NumeroNomina = NumNomina
    rpt.LblDesde.Caption = "Desde: " & Me.LblFecha1.Caption & "   Hasta: " & Me.LblFecha2.Caption
    rpt.AdoNomina.Source = SQlReportes
      
        rpt.AdoNomina.ConnectionString = ConexionReporte
'        ArepNominaProduccionLegal.Show 1
           fPreview.arv.ReportSource = rpt
           fPreview.Show 1
 
   Case "Nomina Produccion Tamaño Legal"
        rpt.AdoNomina.Source = SQlReportes
        rpt.LblDesde.Caption = Me.LblFecha1.Caption
    
        rpt.LblHasta.Caption = Me.LblFecha2.Caption
        'rpt.lblFecha = Format(Now, "dddddd")
        rpt.lbltitulo.Caption = Titulo
        rpt.LblSubtitulo.Caption = SubTitulo
        If Dir(RutaLogo) <> "" Then
          rpt.ImgLogo.Picture = LoadPicture(RutaLogo)
        End If
        rpt.AdoNomina.ConnectionString = ConexionReporte
'        ArepNominaProduccionLegal.Show 1
           fPreview.arv.ReportSource = rpt
           fPreview.Show 1
   
 
   Case "Colilla Produccion Tamaño Legal"
   
      
        ArepColillaProduccionLegal.AdoColillas.Source = SQlReportes
        ArepColillaProduccionLegal.LblPeriodo.Caption = FechaIni & " al " & FechaFin
        ArepColillaProduccionLegal.lbltitulo.Caption = Titulo
        ArepColillaProduccionLegal.AdoColillas.ConnectionString = ConexionReporte
'        ArepColillaProduccionLegal.Show 1
           fPreview.arv.ReportSource = ArepColillaProduccionLegal
           fPreview.Show 1
       Me.AdoDepartamento.Recordset.MoveNext

 
   Case "Nomina Comercial2"
    ArepNominaComercial2.AdoNomina.Source = SQlReportes
    ArepNominaComercial2.lbltitulo.Caption = Titulo
    ArepNominaComercial2.LblSubtitulo.Caption = SubTitulo
    ArepNominaComercial2.ImgLogo.Picture = LoadPicture(RutaLogo)
    ArepNominaComercial2.AdoNomina.ConnectionString = ConexionReporte
    ArepNominaComercial2.LblFecha.Caption = Format(Now, "dddddd")
    ArepNominaComercial2.LblDesde = Me.LblFecha1.Caption
    ArepNominaComercial2.LblHasta = Me.LblFecha2.Caption
'    Dim rpt As Object
'    Dim fPreview As New FrmPreview
    
         Set rpt = New ArepNominaComercial2
         rpt.AdoNomina.ConnectionString = ConexionReporte
         rpt.AdoNomina.Source = SQlReportes
         fPreview.RunReport rpt


     fPreview.Show 1
  
  Case "Nomina Comercial"
    ArepNominaComercial.AdoNomina.Source = SQlReportes
    ArepNominaComercial.lbltitulo.Caption = Titulo
    ArepNominaComercial.LblSubtitulo.Caption = SubTitulo
    ArepNominaComercial.LblDesde.Caption = Me.LblFecha1.Caption
    ArepNominaComercial.LblHasta.Caption = Me.LblFecha2.Caption

     ArepNominaComercial.ImgLogo.Picture = LoadPicture(RutaLogo)

    ArepNominaComercial.AdoNomina.ConnectionString = ConexionReporte
    ArepNominaComercial.LblFecha.Caption = Format(Now, "dddddd")
    'ArepNominaComercial.LblDesde = Me.LblFecha1.Caption
    'ArepNominaComercial.LblHasta = Me.LblFecha2.Caption

         Set rpt = New ArepNominaComercial
         rpt.AdoNomina.ConnectionString = ConexionReporte
         rpt.AdoNomina.Source = SQlReportes
         fPreview.RunReport rpt
         fPreview.Show 1
'     ArepNominaComercial.Show 1

   Case "Nomina Produccion"
    ArepNomina.AdoNomina.Source = SQlReportes
    ArepNomina.lbltitulo.Caption = Titulo
    ArepNomina.LblSubtitulo.Caption = SubTitulo
    If Dir(RutaLogo) <> "" Then
     ArepNomina.ImgLogo.Picture = LoadPicture(RutaLogo)
    End If
    ArepNomina.AdoNomina.ConnectionString = ConexionReporte
    ArepNomina.LblFecha.Caption = Format(Now, "dddddd")
    ArepNomina.LblDesde = Me.LblFecha1.Caption
    ArepNomina.LblHasta = Me.LblFecha2.Caption
    ArepNomina.NumeroNomina = NumNomina
'    ArepNomina.Show 1
         Set rpt = New ArepNomina
         rpt.AdoNomina.ConnectionString = ConexionReporte
         rpt.AdoNomina.Source = SQlReportes
         fPreview.RunReport rpt


     fPreview.Show 1
    
   Case "Nomina Bono Produccion"
    ArepNominaBono.AdoNomina.Source = SQlReportes
    ArepNominaBono.lbltitulo.Caption = Titulo
    ArepNominaBono.LblSubtitulo.Caption = SubTitulo
    ArepNominaBono.ImgLogo.Picture = LoadPicture(RutaLogo)
    ArepNominaBono.AdoNomina.ConnectionString = ConexionReporte
    ArepNominaBono.LblFecha.Caption = Format(Now, "dddddd")
    ArepNominaBono.LblDesde = Me.LblFecha1.Caption
    ArepNominaBono.LblHasta = Me.LblFecha2.Caption
'    ArepNominaBono.Show 1
           fPreview.arv.ReportSource = ArepNominaBono
           fPreview.Show 1
   
  End Select
End If
Exit Sub
TipoErr:
    ControlErrores
End Sub

Private Sub CmdPrNomSocios_Click()
'On Error GoTo TipoErr
'
'Dim SQLDestajos As String
'Dim SQlDeducciones As String
'Dim SqlEmpleados As String
'
'Dim Lunes As Double
'Dim Martes As Double
'Dim Miercoles As Double
'Dim Jueves As Double
'Dim Viernes As Double
'Dim Sabado As Double
'Dim Domingo As Double
'
'Dim FCR As Double
'Dim PS As Double
'Dim Afiliacion As Double
'Dim Cofel As Double
'Dim Almacen As Double
'Dim DiadeLeche As Double
'Dim Cuota As Double
'
'
'CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
'SQLNomina = "SELECT Nomina.* From Nomina WHERE Nomina.Activa=True AND Nomina.CodTipoNomina= '" & CodTipoNomina & "'"
'DtaNomina.RecordSource = SQLNomina
'DtaNomina.Refresh
'
'NumNomina = DtaNomina.Recordset("NumNomina")
'
'DtaNewNomina.Refresh
'Do While Not DtaNewNomina.Recordset.EOF
'    If DtaNewNomina.Recordset.NumNomina = NumNomina Then
'        DtaNewNomina.Recordset.Delete
'    End If
'DtaNewNomina.Recordset.MoveNext
'Loop
'
'SqlEmpleados = "SELECT Empleado.*, Empleado.CodTipoNomina From Empleado WHERE Empleado.CodTipoNomina= '" & CodTipoNomina & "'"
'DtaEmpleados.RecordSource = SqlEmpleados
'DtaEmpleados.Refresh
'
'Do While Not DtaEmpleados.Recordset.EOF
'   CodEmpleado = DtaEmpleados.Recordset("CodEmpleado")
'        SQLDestajos = "SELECT DestalleDestajos.CodTipoDestajo, TipoDestajo.Destajo, DestalleDestajos.Cantidad, TipoDestajo.Monto, DestalleDestajos.CodEmpleado, DestalleDestajos.NUmNomina FROM TipoDestajo INNER JOIN DestalleDestajos ON TipoDestajo.COdTipoDestajo = DestalleDestajos.CodTipoDestajo WHERE DestalleDestajos.CodEmpleado='" & CodEmpleado & "' AND DestalleDestajos.NUmNomina= " & NumNomina & ""
'        DtaDestajo.RecordSource = SQLDestajos
'        DtaDestajo.Refresh
'
'        Lunes = 0
'        Martes = 0
'        Miercoles = 0
'        Jueves = 0
'        Viernes = 0
'        Sabado = 0
'        Domingo = 0
'
'        Do While Not DtaDestajo.Recordset.EOF
'           If InStr(1, DtaDestajo.Recordset.destajo, "Lunes") <> 0 Then
'              Lunes = Lunes + DtaDestajo.Recordset.Monto * DtaDestajo.Recordset.cantidad
'           ElseIf InStr(1, DtaDestajo.Recordset.destajo, "Martes") <> 0 Then
'              Martes = Martes + DtaDestajo.Recordset.Monto * DtaDestajo.Recordset.cantidad
'           ElseIf InStr(1, DtaDestajo.Recordset.destajo, "Miercoles") <> 0 Then
'              Miercoles = Miercoles + DtaDestajo.Recordset.Monto * DtaDestajo.Recordset.cantidad
'           ElseIf InStr(1, DtaDestajo.Recordset.destajo, "Jueves") <> 0 Then
'              Jueves = Jueves + DtaDestajo.Recordset.Monto * DtaDestajo.Recordset.cantidad
'           ElseIf InStr(1, DtaDestajo.Recordset.destajo, "Viernes") <> 0 Then
'              Viernes = Viernes + DtaDestajo.Recordset.Monto * DtaDestajo.Recordset.cantidad
'           ElseIf InStr(1, DtaDestajo.Recordset.destajo, "Sabado") <> 0 Then
'              Sabado = Sabado + DtaDestajo.Recordset.Monto * DtaDestajo.Recordset.cantidad
'           ElseIf InStr(1, DtaDestajo.Recordset.destajo, "Domingo") <> 0 Then
'              Domingo = Domingo + DtaDestajo.Recordset.Monto * DtaDestajo.Recordset.cantidad
'           End If
'
'        DtaDestajo.Recordset.MoveNext
'        Loop
'
'        FCR = 0
'        PS = 0
'        Afiliacion = 0
'        Cofel = 0
'        Almacen = 0
'        DiadeLeche = 0
'        Cuota = 0
'
'        SQlDeducciones = "SELECT Deduccion.CodEmpleado, DetalleDeduccion.Valor, DetalleDeduccion.NumNomina, Deduccion.CodTipoDeduccion FROM Deduccion INNER JOIN DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion WHERE Deduccion.CodEmpleado='" & CodEmpleado & "' AND DetalleDeduccion.NUmNomina= " & NumNomina & ""
'        DtaDeducciones.RecordSource = SQlDeducciones
'        DtaDeducciones.Refresh
'        Do While Not DtaDeducciones.Recordset.EOF
'            Select Case DtaDeducciones.Recordset("codtipodeduccion")
'                Case "03"
'                        FCR = DtaDeducciones.Recordset("valor")
'                Case "04"
'                        PS = DtaDeducciones.Recordset("valor")
'                Case "05"
'                        Afiliacion = DtaDeducciones.Recordset("valor")
'                Case "06"
'                        Cofel = DtaDeducciones.Recordset("valor")
'                Case "07"
'                        Almacen = DtaDeducciones.Recordset("valor")
'                Case "08"
'                        DiadeLeche = DtaDeducciones.Recordset("valor")
'                Case "09"
'                        Cuota = DtaDeducciones.Recordset("valor")
'            End Select
'        DtaDeducciones.Recordset.MoveNext
'        Loop
'    DtaNewNomina.Recordset.AddNew
'        DtaNewNomina.Recordset.NumNomina = NumNomina
'        DtaNewNomina.Recordset.CodEmpleado = CodEmpleado
'        DtaNewNomina.Recordset.Lunes = Lunes
'        DtaNewNomina.Recordset.Martes = Martes
'        DtaNewNomina.Recordset.Miercoles = Miercoles
'        DtaNewNomina.Recordset.Jueves = Jueves
'        DtaNewNomina.Recordset.Viernes = Viernes
'        DtaNewNomina.Recordset.Sabado = Sabado
'        DtaNewNomina.Recordset.Domingo = Domingo
'        DtaNewNomina.Recordset.FCR = FCR
'        DtaNewNomina.Recordset.PS = PS
'        DtaNewNomina.Recordset.Afiliacion = Afiliacion
'        DtaNewNomina.Recordset.Cofel = Cofel
'        DtaNewNomina.Recordset.Almacen = Almacen
'        DtaNewNomina.Recordset.DiadeLeche = DiadeLeche
'        DtaNewNomina.Recordset.Cuota = Cuota
'    DtaNewNomina.Recordset.Update
'
'
'DtaEmpleados.Recordset.MoveNext
'Loop
'ARNomina.Show
'
'Exit Sub
'TipoErr:
'ControlErrores
End Sub

Private Sub CmdSalir_Click()
Unload Me

End Sub

Private Sub Command1_Click()

End Sub

Private Sub DbgrNominas_Click()
On Error GoTo TipoErr
DtaNominas.Refresh
Do While Not DtaNominas.Recordset.EOF
   If Me.DtaNominas.Recordset("CodTipoNomina") = Me.DtaTipoNomina.Recordset("CodTipoNomina") And DtaNominas.Recordset("Activa") = True Then
      LblFecha1.Caption = Format(DtaNominas.Recordset("FechaNominaINI"), "Long Date")
      LblFecha2.Caption = Format(DtaNominas.Recordset("FechaNomina"), "Long Date")
      Me.TxtFechaIni.Text = Format(DtaNominas.Recordset("FechaNominaINI"), "dd/mm/yyyy")
   End If
   DtaNominas.Recordset.MoveNext
Loop

Exit Sub
TipoErr:
ControlErrores
End Sub

Private Sub DtaTipoNomina_Reposition()
On Error GoTo TipoErr
If Not DtaTipoNomina.Recordset.EOF Then
        DtaNominas.Refresh
        Do While Not DtaNominas.Recordset.EOF
           If DtaNominas.Recordset("CodTipoNomina") = DtaTipoNomina.Recordset("CodTipoNomina") And DtaNominas.Recordset("Activa") = True Then
              LblFecha1.Caption = Format(DtaNominas.Recordset("FechaNominaINI"), "Long Date")
              LblFecha2.Caption = Format(DtaNominas.Recordset("FechaNomina"), "Long Date")
              Me.TxtFechaIni.Text = Format(DtaNominas.Recordset("FechaNominaINI"), "dd/mm/yyyy")
           End If
           DtaNominas.Recordset.MoveNext
        Loop
End If

Exit Sub
TipoErr:
ControlErrores

End Sub


Private Sub Form_Load()

On Error GoTo TipoErr
Dim SQLTipoNomina As String
 Me.DbgrNominas.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DbgrNominas.OddRowStyle.BackColor = &H80000005
 Me.DbgrNominas.AlternatingRowStyle = True
 
 Me.Picture1.BackColor = RGB(222, 227, 247)
 Me.Frame1.BackColor = RGB(222, 227, 247)

With Me.AdoConfiguracion
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT HorasPuntualidad, HorasSeptimo, HorasBasico, MontoPuntualidad, MontoViaticos From ConfiguracionIncentivo"
   .Refresh
End With

With Me.DtaExporta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With


With Me.AdoDetalleIncentivo
    .ConnectionString = Conexion
End With
 
 
With Me.adoIncentivo
   .ConnectionString = Conexion
End With
 
 
With Me.AdoHistorico
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
 End With
 
 With Me.AdoBusca
   .ConnectionString = Conexion
 End With
 
 With Me.DtaAuxiliar
   .ConnectionString = Conexion
 End With
 
  With Me.AdoViaticos
   .ConnectionString = Conexion
 End With
 
   With Me.AdoDetalleProduccionManual
   .ConnectionString = Conexion
 End With
 
 With Me.AdoDepartamento
   .ConnectionString = Conexion
   .RecordSource = "SELECT  * From departamento"
   .Refresh
 End With
 
   With Me.AdoDetalleViaticos
   .ConnectionString = Conexion
 End With
 
 
  With Me.AdoPeriodoFiscal
   .ConnectionString = Conexion
 End With
 
  With Me.AdoSuspendido
   .ConnectionString = Conexion
 End With
 
 With Me.AdoIncentivoPro
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
 End With

With Me.AdoAntiguedad
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT años_acum, porcent From Antiguedad"
   .Refresh
 End With

With Me.DtaNominas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaHorasProducidas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaDetalleNominaAnterior
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaNominaMes
  '.DatabaseName = Ruta
  .ConnectionString = Conexion
End With

With Me.DtaComisiones
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaConsecutivos
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

With Me.DtaDestajo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaDetalleDeduccion
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaDetalleIncentivo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaDetalleNomina
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaEmpleados
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With


With Me.DtaHrsExtras
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With


With Me.DtaIncentivos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With


With Me.DtaInss
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaIr
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With



With Me.DtaMovPrestamo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaNomina
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaNomSubsidios
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
End With

With Me.DtaControles
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaPagosMensuales
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaNewNomina
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

SQLTipoNomina = "SELECT  CodTipoNomina, Nomina, Periodo, TipoPago, MantValor, CalcularHoraTrabajada, Activa,IrUltimaSemana, Moneda FROM TipoNomina WHERE TipoNomina.Activa=1"
DtaTipoNomina.RecordSource = SQLTipoNomina
DtaTipoNomina.Refresh




If DtaTipoNomina.Recordset.EOF Then
   MsgBox "No hay Nóminas Activas"
   Unload Me
   Exit Sub
End If

Exit Sub
TipoErr:
ControlErrores
End Sub

Private Sub HorasExtra_Click()
ArepNomina.lbltitulo.Caption = Titulo
ArepNomina.LblSubtitulo.Caption = SubTitulo
ArepNomina.ImgLogo.Picture = LoadPicture(RutaLogo)
End Sub

Private Sub FrmReportes_Click()
On Error GoTo TipoErrs
CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
NumNomina = DtaNomina.Recordset("NumNomina")
NumeroNominas = DtaNomina.Recordset("NumNomina")
FrmNominaActiva.Show 1
Exit Sub
TipoErrs:
ControlErrores
End Sub

'comprasmymmantica.com.ni

Private Sub Image1_Click()

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

Quien = "CalcularNomina"

Espacio = " "
Select Case Quien
 Case "CalcularNomina"
       '//////////////////////Cargo la Consulta de la Nomina///////////////////////
       NumNomina = FrmCalcularNomina.DtaNomina.Recordset("NumNomina")

SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
SQlReportes = SQlReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
SQlReportes = SQlReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
SQlReportes = SQlReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
SQlReportes = SQlReportes & "                      Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
SQlReportes = SQlReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion AS TotalDevengado," & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion)" & vbLf
SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion,Empleado.CodEmpleado1" & vbLf
SQlReportes = SQlReportes & " FROM         Nomina INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       Grupo INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       Cargo INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       TipoNomina INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
SQlReportes = SQlReportes & "                       TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina" & vbLf
SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ") AND((dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
SQlReportes = SQlReportes & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia)" & vbLf
SQlReportes = SQlReportes & "                      - (dbo.DetalleNomina.Prestamo + dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.MontoIR + dbo.DetalleNomina.Deducciones) <> 0)" & vbLf
SQlReportes = SQlReportes & " ORDER BY Empleado.CodEmpleado1" & vbLf


       Me.DtaConsulta.RecordSource = SQlReportes
       Me.DtaConsulta.Refresh

       Mes = Month(Me.DtaConsulta.Recordset("FechaNomina"))
       Año = Year(Me.DtaConsulta.Recordset("FechaNomina"))
       CanDias = Day(Me.DtaConsulta.Recordset("FechaNomina"))
       Dias = Day(Me.DtaConsulta.Recordset("FechaNomina"))
       Cod = 1

       ConvertirMes (Mes)
      If CanDias > 15 Then
         QuinLetra = "Segunda Quincena de " & Convertir
      Else
         QuinLetra = "Primera Quincena de" & Convertir
      End If
  


End Select
            
   res = Bitacora(Now, NombreUsuario, "Calcular Nomina", "Se Exporto la Nomina Bancentro: " & NumNomina)
   
    Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
    'Call Formato_Excel(8, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
V = 2
H = 1
i = 1
'           objExcel.ActiveSheet.Columns("B").NumberFormat = "000#"
           
           objExcel.ActiveSheet.Cells(1, 1) = "B"
            objExcel.ActiveSheet.Cells(1, 6) = Año
            objExcel.ActiveSheet.Cells(1, 7) = Mes
            objExcel.ActiveSheet.Cells(1, 8) = Dias

 
     Do While Not Me.DtaConsulta.Recordset.EOF 'esto nos sirve pa leer los datos desde
  CodEmpleado = DtaConsulta.Recordset("CodEmpleado")
 'la tabla de access para despues colocarlos en las celdas correspondientes
       Nombre = Me.DtaConsulta.Recordset("Nombres")
       Neto = Format(Me.DtaConsulta.Recordset("NetoPagar"), "####0.00")
       Neto1 = Format(Me.DtaConsulta.Recordset("NetoPagar"), "##,##0.00")
       Longitud = Len(Neto)
       NetoT = Mid(Neto, Longitud - 1, 3)
       NetoT = (Mid(Neto, 1, Longitud - 3)) & NetoT
       With DtaConsulta.Recordset
       
           If Not (V = 1) Then
             objExcel.ActiveSheet.Cells(V, H) = "T"
           End If
            'objExcel.Cells(1, 1).Format = Text
            objExcel.ActiveSheet.Cells(V, H + 1) = i
            objExcel.ActiveSheet.Cells(V, H + 2) = Año
            objExcel.ActiveSheet.Cells(V, H + 3) = Mes
            objExcel.ActiveSheet.Cells(V, H + 4) = Dias
            objExcel.ActiveSheet.Cells(V, H + 5) = NetoT
            objExcel.ActiveSheet.Cells(V, H + 6) = QuinLetra
            Nombres = Mid(Nombre, 1, 25)
             objExcel.ActiveSheet.Cells(V, H + 12) = Nombres
            V = V + 1
            i = i + 1
            TotalNomina = TotalNomina + Neto1
            .MoveNext
   
        End With
     Loop
     
     
       Neto = Format(TotalNomina, "####0.00")
       Longitud = Len(Neto)
       NetoT = Mid(Neto, Longitud - 1, 3)
       NetoT = (Mid(Neto, 1, Longitud - 3)) & NetoT
   
            objExcel.ActiveSheet.Cells(1, 10) = i - 1
            objExcel.ActiveSheet.Cells(1, 9) = NetoT
       

       objExcel.ActiveSheet.Columns("A").ColumnWidth = 1
       objExcel.ActiveSheet.Columns("B").ColumnWidth = 4
        objExcel.ActiveSheet.Columns("C").ColumnWidth = 4
        objExcel.ActiveSheet.Columns("D").ColumnWidth = 2
        objExcel.ActiveSheet.Columns("G").ColumnWidth = 2
        objExcel.ActiveSheet.Columns("H").ColumnWidth = 2
        objExcel.ActiveSheet.Columns("I").ColumnWidth = 13
         objExcel.ActiveSheet.Columns("J").ColumnWidth = 3
         objExcel.ActiveSheet.Columns("K").ColumnWidth = 30
         objExcel.ActiveSheet.Columns("L").ColumnWidth = 1
         objExcel.ActiveSheet.Columns("M").ColumnWidth = 30
         
 
 Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto

Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub PushButton2_Click()
Dim SQlReportes As String, Directorio As String, Respuesta As String

CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
SQLNomina = "SELECT Nomina.* From Nomina WHERE Nomina.Activa=1 AND Nomina.CodTipoNomina= '" & CodTipoNomina & "'"
DtaNomina.RecordSource = SQLNomina
DtaNomina.Refresh

NumNomina = DtaNomina.Recordset("NumNomina")

SQlReportes = "SELECT Nomina.NumNomina, Empleado.NumeroInss, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.TarifaHoraria, DetalleNomina.SalarioBasico, DetalleNomina.SeptimoDia, DetalleNomina.Destajo, DetalleNomina.HE, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.OtrosIngresos, DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion AS TotalDevengado, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Deducciones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Prestamo, " & _
               "DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir, (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.FechaNomina, Cargo.Cargo, DetalleNomina.IRPatronal, Empleado.NumCedula, departamento.departamento FROM Nomina INNER JOIN Grupo INNER JOIN Cargo INNER JOIN TipoNomina INNER JOIN Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN " & _
               "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento  " & _
               "WHERE (Nomina.NumNomina = " & NumNomina & ") AND ((DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.SeptimoDia) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) <> 0) ORDER BY Nomina.NumNomina, Empleado.Nombre1"

'SQlReportes = "SELECT Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS,Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13,Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2,Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo, Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo,DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos,DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas, DetalleNomina.BonoProduccion, DetalleNomina.Prestamo, " & _
'              "DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre,DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo,Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres,DetalleNomina.HE,DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasTurno+ DetalleNomina.Antiguedad AS TotalDevengado,DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir, " & _
'              "(DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.HorasTurno+ DetalleNomina.Antiguedad) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar, DetalleNomina.TarifaHoraria, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1, Empleado.NumeroInss, Empleado.NumCedula, departamento.departamento,DetalleNomina.HorasTurno,DetalleNomina.HTurno,DetalleNomina.Antiguedad,DetalleNomina.AñoAntiguedad " & _
'              "FROM  Nomina INNER JOIN Grupo INNER JOIN Cargo INNER JOIN TipoNomina INNER JOIN Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON " & _
'              "TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento  " & _
'              "WHERE (Nomina.NumNomina = " & NumNomina & ") AND ((DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) <> 0) ORDER BY Nomina.NumNomina, Empleado.CodEmpleado1"

Me.CommonDialog1.ShowSave
Directorio = ""
Directorio = Me.CommonDialog1.FileName + ".xls"
Respuesta = Exportar_ADO_Excel(Conexion, SQlReportes, Directorio)

End Sub
