VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmImportar 
   Caption         =   "AdoRegistros"
   ClientHeight    =   6315
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   ScaleHeight     =   6917.841
   ScaleMode       =   0  'User
   ScaleWidth      =   38407.89
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc DtaHorasExtra 
      Height          =   495
      Left            =   7440
      Top             =   7080
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc DtaHrasExtras 
      Height          =   375
      Left            =   360
      Top             =   8640
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc AdoSolicitud 
      Height          =   375
      Left            =   9480
      Top             =   7800
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc AdoConsecutivo 
      Height          =   375
      Left            =   7800
      Top             =   7800
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Adodc2"
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
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   13095
      _Version        =   786432
      _ExtentX        =   23098
      _ExtentY        =   10398
      _StockProps     =   68
      Color           =   64
      ItemCount       =   7
      Item(0).Caption =   "Inicial"
      Item(0).ControlCount=   16
      Item(0).Control(0)=   "TDBTipo"
      Item(0).Control(1)=   "CmdIniciar2"
      Item(0).Control(2)=   "SkinLabel1"
      Item(0).Control(3)=   "DTPFechaIni"
      Item(0).Control(4)=   "Frame1"
      Item(0).Control(5)=   "CmdIniciar"
      Item(0).Control(6)=   "CmdSalir"
      Item(0).Control(7)=   "CmdBuscarLogo"
      Item(0).Control(8)=   "TxtRutaLogo"
      Item(0).Control(9)=   "SkinLabel25"
      Item(0).Control(10)=   "TDBGridNominas"
      Item(0).Control(11)=   "osProgress1"
      Item(0).Control(12)=   "TDBGrid1"
      Item(0).Control(13)=   "DTPFechaFin"
      Item(0).Control(14)=   "SkinLabel2"
      Item(0).Control(15)=   "Label34"
      Item(1).Caption =   "Saldo Vacaciciones"
      Item(1).ControlCount=   10
      Item(1).Control(0)=   "Label2"
      Item(1).Control(1)=   "gridSaldos"
      Item(1).Control(2)=   "dtpFin"
      Item(1).Control(3)=   "btnIniciar"
      Item(1).Control(4)=   "listSaldos"
      Item(1).Control(5)=   "txtCodigo"
      Item(1).Control(6)=   "txtNombre"
      Item(1).Control(7)=   "pbSaldos"
      Item(1).Control(8)=   "Command3"
      Item(1).Control(9)=   "ChkSegunNomina"
      Item(2).Caption =   "Movimiento Salarial"
      Item(2).ControlCount=   6
      Item(2).Control(0)=   "Label3"
      Item(2).Control(1)=   "Command4"
      Item(2).Control(2)=   "Command5"
      Item(2).Control(3)=   "gridMovimientoSalarial"
      Item(2).Control(4)=   "Label4"
      Item(2).Control(5)=   "listMovimientoSalaria"
      Item(3).Caption =   "Horas Extra"
      Item(3).ControlCount=   8
      Item(3).Control(0)=   "Label7"
      Item(3).Control(1)=   "btnExcelHE"
      Item(3).Control(2)=   "pbHE"
      Item(3).Control(3)=   "listHE"
      Item(3).Control(4)=   "btnImportarHE"
      Item(3).Control(5)=   "gridHorasExtra"
      Item(3).Control(6)=   "lblNombre"
      Item(3).Control(7)=   "lblCodigo"
      Item(4).Caption =   "Deducciones"
      Item(4).ControlCount=   9
      Item(4).Control(0)=   "TDBCombo1"
      Item(4).Control(1)=   "Command1"
      Item(4).Control(2)=   "Text1"
      Item(4).Control(3)=   "SkinLabel3"
      Item(4).Control(4)=   "Label1"
      Item(4).Control(5)=   "TDBGrid2"
      Item(4).Control(6)=   "CmdIniciar3"
      Item(4).Control(7)=   "ProgressBar1"
      Item(4).Control(8)=   "Command2"
      Item(5).Caption =   "Ingresos"
      Item(5).ControlCount=   9
      Item(5).Control(0)=   "Command6"
      Item(5).Control(1)=   "Command8"
      Item(5).Control(2)=   "SkinLabel4"
      Item(5).Control(3)=   "ProgressBar2"
      Item(5).Control(4)=   "Label5"
      Item(5).Control(5)=   "TDBGridIngresos"
      Item(5).Control(6)=   "CmdIniciarIngresos"
      Item(5).Control(7)=   "TxtRutaIngresos"
      Item(5).Control(8)=   "TDBComboIngresos"
      Item(6).Caption =   "Solicitud"
      Item(6).ControlCount=   0
      Begin VB.CheckBox ChkSegunNomina 
         Caption         =   "Segun Nomina Activa"
         Height          =   255
         Left            =   -61480
         TabIndex        =   62
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton Command8 
         Height          =   375
         Left            =   -63040
         Picture         =   "FrmImportarEmpleados.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TxtRutaIngresos 
         Height          =   375
         Left            =   -68440
         TabIndex        =   55
         Top             =   480
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.CommandButton CmdIniciarIngresos 
         Caption         =   "Iniciar"
         Height          =   495
         Left            =   -58360
         TabIndex        =   54
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Salir"
         Height          =   495
         Left            =   -58360
         TabIndex        =   53
         Top             =   3840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton btnExcelHE 
         Caption         =   "Buscar Excel"
         Height          =   375
         Left            =   -63760
         TabIndex        =   52
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin XtremeSuiteControls.ProgressBar pbHE 
         Height          =   375
         Left            =   -61960
         TabIndex        =   50
         Top             =   4200
         Visible         =   0   'False
         Width           =   3255
         _Version        =   786432
         _ExtentX        =   5741
         _ExtentY        =   661
         _StockProps     =   93
      End
      Begin VB.ListBox listHE 
         Height          =   2010
         Left            =   -61960
         TabIndex        =   49
         Top             =   1680
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CommandButton btnImportarHE 
         Caption         =   "Importar"
         Height          =   375
         Left            =   -61960
         TabIndex        =   46
         Top             =   3720
         Visible         =   0   'False
         Width           =   3255
      End
      Begin TrueOleDBGrid80.TDBGrid gridHorasExtra 
         Height          =   3495
         Left            =   -69040
         TabIndex        =   45
         Top             =   1080
         Visible         =   0   'False
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6165
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Codigo"
         Columns(0).DataField=   "Codigo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nombre"
         Columns(1).DataField=   "Nombre"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Horas"
         Columns(2).DataField=   "Horas"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         DirectionAfterTab=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=86,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Named:id=33:Normal"
         _StyleDefs(43)  =   ":id=33,.parent=0"
         _StyleDefs(44)  =   "Named:id=34:Heading"
         _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(46)  =   ":id=34,.wraptext=-1"
         _StyleDefs(47)  =   "Named:id=35:Footing"
         _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(49)  =   "Named:id=36:Selected"
         _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(51)  =   "Named:id=37:Caption"
         _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(53)  =   "Named:id=38:HighlightRow"
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(55)  =   "Named:id=39:EvenRow"
         _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(57)  =   "Named:id=40:OddRow"
         _StyleDefs(58)  =   ":id=40,.parent=33"
         _StyleDefs(59)  =   "Named:id=41:RecordSelector"
         _StyleDefs(60)  =   ":id=41,.parent=34"
         _StyleDefs(61)  =   "Named:id=42:FilterBar"
         _StyleDefs(62)  =   ":id=42,.parent=33"
      End
      Begin VB.ListBox listMovimientoSalaria 
         Height          =   1815
         Left            =   -58840
         TabIndex        =   44
         Top             =   3000
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Iniciar"
         Height          =   495
         Left            =   -58600
         TabIndex        =   42
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Buscar Archivo"
         Height          =   495
         Left            =   -58600
         TabIndex        =   41
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin TrueOleDBGrid80.TDBGrid gridMovimientoSalarial 
         Height          =   3975
         Left            =   -69280
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   7011
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Codigo"
         Columns(0).DataField=   "Codigo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nombres"
         Columns(1).DataField=   "Nombres"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "SalarioBasicoMasNivelacion"
         Columns(2).DataField=   "SalarioBasicoMasNivelacion"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "AumentoCordobas"
         Columns(3).DataField=   "AumentoCordobas"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         DirectionAfterTab=   1
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Named:id=33:Normal"
         _StyleDefs(47)  =   ":id=33,.parent=0"
         _StyleDefs(48)  =   "Named:id=34:Heading"
         _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   ":id=34,.wraptext=-1"
         _StyleDefs(51)  =   "Named:id=35:Footing"
         _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   "Named:id=36:Selected"
         _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(55)  =   "Named:id=37:Caption"
         _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(57)  =   "Named:id=38:HighlightRow"
         _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=39:EvenRow"
         _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(61)  =   "Named:id=40:OddRow"
         _StyleDefs(62)  =   ":id=40,.parent=33"
         _StyleDefs(63)  =   "Named:id=41:RecordSelector"
         _StyleDefs(64)  =   ":id=41,.parent=34"
         _StyleDefs(65)  =   "Named:id=42:FilterBar"
         _StyleDefs(66)  =   ":id=42,.parent=33"
      End
      Begin MSComDlg.CommonDialog dialogSaldos 
         Left            =   840
         Top             =   5520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Iniciar"
         Height          =   375
         Left            =   -61720
         TabIndex        =   38
         Top             =   4200
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.ListBox listSaldos 
         Height          =   2595
         Left            =   -61720
         TabIndex        =   35
         Top             =   1560
         Visible         =   0   'False
         Width           =   3375
      End
      Begin MSComctlLib.ProgressBar pbSaldos 
         Height          =   495
         Left            =   -68560
         TabIndex        =   34
         Top             =   4200
         Visible         =   0   'False
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton btnIniciar 
         Caption         =   "Buscar Excel"
         Height          =   375
         Left            =   -63040
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   375
         Left            =   -64600
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   43105
      End
      Begin TrueOleDBGrid70.TDBGrid gridSaldos 
         Height          =   3135
         Left            =   -68560
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5530
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Codigo"
         Columns(0).DataField=   "Codigo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nombre"
         Columns(1).DataField=   "Nombre"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Saldo"
         Columns(2).DataField=   "Saldo"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1693"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=5636"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=5556"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Named:id=33:Normal"
         _StyleDefs(49)  =   ":id=33,.parent=0"
         _StyleDefs(50)  =   "Named:id=34:Heading"
         _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   ":id=34,.wraptext=-1"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   495
         Left            =   -58360
         TabIndex        =   29
         Top             =   3840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton CmdIniciar3 
         Caption         =   "Iniciar"
         Height          =   495
         Left            =   -58360
         TabIndex        =   26
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   -68440
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   -63040
         Picture         =   "FrmImportarEmpleados.frx":04B6
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TxtRutaLogo 
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   480
         Width           =   5295
      End
      Begin VB.CommandButton CmdBuscarLogo 
         Height          =   375
         Left            =   6840
         Picture         =   "FrmImportarEmpleados.frx":096C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   495
         Left            =   11400
         TabIndex        =   11
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton CmdIniciar 
         Caption         =   "Iniciar"
         Height          =   495
         Left            =   11400
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   7440
         TabIndex        =   7
         Top             =   360
         Width           =   3855
         Begin VB.OptionButton Option1 
            Caption         =   "Empleados"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Nominas"
            Height          =   255
            Left            =   2040
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CommandButton CmdIniciar2 
         Caption         =   "Iniciar"
         Height          =   495
         Left            =   11400
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin TrueOleDBList80.TDBCombo TDBTipo 
         Bindings        =   "FrmImportarEmpleados.frx":0E22
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   960
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         _EDITHEIGHT     =   556
         _GAPHEIGHT      =   53
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
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits.Count    =   1
         Appearance      =   1
         BorderStyle     =   1
         ComboStyle      =   0
         AutoCompletion  =   0   'False
         LimitToList     =   0   'False
         ColumnHeaders   =   -1  'True
         ColumnFooters   =   0   'False
         DataMode        =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   0
         Caption         =   ""
         EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         AutoSize        =   -1  'True
         ListField       =   "Nomina"
         BoundColumn     =   ""
         IntegralHeight  =   0   'False
         CellTipsWidth   =   0
         CellTipsDelay   =   1000
         AutoDropdown    =   0   'False
         RowTracking     =   -1  'True
         RightToLeft     =   0   'False
         RowMember       =   ""
         MouseIcon       =   0
         MouseIcon.vt    =   3
         MousePointer    =   0
         MatchEntryTimeout=   2000
         OLEDragMode     =   0
         OLEDropMode     =   0
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         AddItemSeparator=   ";"
         _PropDict       =   $"FrmImportarEmpleados.frx":0E3E
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(42)  =   ":id=34,.wraptext=-1"
         _StyleDefs(43)  =   "Named:id=35:Footing"
         _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(45)  =   "Named:id=36:Selected"
         _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(47)  =   "Named:id=37:Caption"
         _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(49)  =   "Named:id=38:HighlightRow"
         _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(51)  =   "Named:id=39:EvenRow"
         _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(53)  =   "Named:id=40:OddRow"
         _StyleDefs(54)  =   ":id=40,.parent=33"
         _StyleDefs(55)  =   "Named:id=41:RecordSelector"
         _StyleDefs(56)  =   ":id=41,.parent=34"
         _StyleDefs(57)  =   "Named:id=42:FilterBar"
         _StyleDefs(58)  =   ":id=42,.parent=33"
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   11400
         OleObjectBlob   =   "FrmImportarEmpleados.frx":0EE8
         TabIndex        =   5
         Top             =   2640
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPFechaIni 
         Height          =   375
         Left            =   11400
         TabIndex        =   6
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   41829
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmImportarEmpleados.frx":0F58
         TabIndex        =   14
         Top             =   525
         Width           =   1335
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGridNominas 
         Height          =   3015
         Left            =   13560
         TabIndex        =   15
         Top             =   1320
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5318
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "CodEmpleado"
         Columns(0).DataField=   "CodEmpleado"
         Columns(0).NumberFormat=   "000#"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nombres"
         Columns(1).DataField=   "Nombres"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "SalarioBasico"
         Columns(2).DataField=   "SalarioBasico"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Destajo"
         Columns(3).DataField=   "Destajo"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "HorasExtras"
         Columns(4).DataField=   "HorasExtras"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Comisiones"
         Columns(5).DataField=   "Comisiones"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Incentivos"
         Columns(6).DataField=   "Incentivos"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "VacacionesPagadas"
         Columns(7).DataField=   "VacacionesPagadas"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "SeptimoDia"
         Columns(8).DataField=   "SeptimoDia"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "IncetivoProduccion"
         Columns(9).DataField=   "IncetivoProduccion"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "OtrosIngresos"
         Columns(10).DataField=   "OtrosIngresos"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "TotalDevengado"
         Columns(11).DataField=   "TotalDevengado"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Prestamo"
         Columns(12).DataField=   "Prestamo"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "MontoINSS"
         Columns(13).DataField=   "MontoINSS"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "MontoIr"
         Columns(14).DataField=   "MontoIr"
         Columns(14).NumberFormat=   "General Number"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "Deducciones"
         Columns(15).DataField=   "Deducciones"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "INSSPatronal"
         Columns(16).DataField=   "INSSPatronal"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   17
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=17"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1693"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1773"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1693"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=2646"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2566"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=1773"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1693"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=1773"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1693"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=1773"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1693"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=1773"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=1693"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=1773"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=1693"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=2646"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2566"
         Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(45)=   "Column(11).Width=873"
         Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=794"
         Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(49)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(53)=   "Column(13).Width=2725"
         Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=2646"
         Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(57)=   "Column(14).Width=2725"
         Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=2646"
         Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(61)=   "Column(15).Width=2725"
         Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=2646"
         Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(65)=   "Column(16).Width=2725"
         Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=2646"
         Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         DirectionAfterTab=   1
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=70,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=74,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=78,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=82,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=86,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=83,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=84,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=85,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=90,.parent=13"
         _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=87,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=88,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=89,.parent=17"
         _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=94,.parent=13"
         _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=91,.parent=14"
         _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=92,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=93,.parent=17"
         _StyleDefs(78)  =   "Splits(0).Columns(12).Style:id=98,.parent=13"
         _StyleDefs(79)  =   "Splits(0).Columns(12).HeadingStyle:id=95,.parent=14"
         _StyleDefs(80)  =   "Splits(0).Columns(12).FooterStyle:id=96,.parent=15"
         _StyleDefs(81)  =   "Splits(0).Columns(12).EditorStyle:id=97,.parent=17"
         _StyleDefs(82)  =   "Splits(0).Columns(13).Style:id=102,.parent=13"
         _StyleDefs(83)  =   "Splits(0).Columns(13).HeadingStyle:id=99,.parent=14"
         _StyleDefs(84)  =   "Splits(0).Columns(13).FooterStyle:id=100,.parent=15"
         _StyleDefs(85)  =   "Splits(0).Columns(13).EditorStyle:id=101,.parent=17"
         _StyleDefs(86)  =   "Splits(0).Columns(14).Style:id=106,.parent=13"
         _StyleDefs(87)  =   "Splits(0).Columns(14).HeadingStyle:id=103,.parent=14"
         _StyleDefs(88)  =   "Splits(0).Columns(14).FooterStyle:id=104,.parent=15"
         _StyleDefs(89)  =   "Splits(0).Columns(14).EditorStyle:id=105,.parent=17"
         _StyleDefs(90)  =   "Splits(0).Columns(15).Style:id=110,.parent=13"
         _StyleDefs(91)  =   "Splits(0).Columns(15).HeadingStyle:id=107,.parent=14"
         _StyleDefs(92)  =   "Splits(0).Columns(15).FooterStyle:id=108,.parent=15"
         _StyleDefs(93)  =   "Splits(0).Columns(15).EditorStyle:id=109,.parent=17"
         _StyleDefs(94)  =   "Splits(0).Columns(16).Style:id=32,.parent=13"
         _StyleDefs(95)  =   "Splits(0).Columns(16).HeadingStyle:id=29,.parent=14"
         _StyleDefs(96)  =   "Splits(0).Columns(16).FooterStyle:id=30,.parent=15"
         _StyleDefs(97)  =   "Splits(0).Columns(16).EditorStyle:id=31,.parent=17"
         _StyleDefs(98)  =   "Named:id=33:Normal"
         _StyleDefs(99)  =   ":id=33,.parent=0"
         _StyleDefs(100) =   "Named:id=34:Heading"
         _StyleDefs(101) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(102) =   ":id=34,.wraptext=-1"
         _StyleDefs(103) =   "Named:id=35:Footing"
         _StyleDefs(104) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(105) =   "Named:id=36:Selected"
         _StyleDefs(106) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(107) =   "Named:id=37:Caption"
         _StyleDefs(108) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(109) =   "Named:id=38:HighlightRow"
         _StyleDefs(110) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(111) =   "Named:id=39:EvenRow"
         _StyleDefs(112) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(113) =   "Named:id=40:OddRow"
         _StyleDefs(114) =   ":id=40,.parent=33"
         _StyleDefs(115) =   "Named:id=41:RecordSelector"
         _StyleDefs(116) =   ":id=41,.parent=34"
         _StyleDefs(117) =   "Named:id=42:FilterBar"
         _StyleDefs(118) =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
         Height          =   3015
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5318
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Codigo"
         Columns(0).DataField=   "Codigo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nombre1"
         Columns(1).DataField=   "Nombre1"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nombre2"
         Columns(2).DataField=   "Nombre2"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Apellido1"
         Columns(3).DataField=   "Apellido1"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Apellido2"
         Columns(4).DataField=   "Apellido2"
         Columns(4).NumberFormat=   "hh:mm:ss"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Direccion"
         Columns(5).DataField=   "Direccion"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Nacionalidad"
         Columns(6).DataField=   "Nacionalidad"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Sexo"
         Columns(7).DataField=   "Sexo"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Fecha Ingreso"
         Columns(8).DataField=   "FechaIngreso"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Cargo"
         Columns(9).DataField=   "Cargo"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Departamento"
         Columns(10).DataField=   "Departamento"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "RUC"
         Columns(11).DataField=   "RUC"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "INSS"
         Columns(12).DataField=   "INSS"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "CEDULA"
         Columns(13).DataField=   "CEDULA"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Hijos"
         Columns(14).DataField=   "Hijos"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "TipoNomina"
         Columns(15).DataField=   "TipoNomina"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "GrupoNomina"
         Columns(16).DataField=   "GrupoNomina"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "SueldoPeriodo"
         Columns(17).DataField=   "SueldoPeriodo"
         Columns(17).NumberFormat=   "General Number"
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   0
         Columns(18)._MaxComboItems=   5
         Columns(18).Caption=   "TarifaHoraria"
         Columns(18).DataField=   "TarifaHoraria"
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(19)._VlistStyle=   0
         Columns(19)._MaxComboItems=   5
         Columns(19).Caption=   "FechaNacimiento"
         Columns(19).DataField=   "FechaNacimiento"
         Columns(19).NumberFormat=   "Short Date"
         Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(20)._VlistStyle=   0
         Columns(20)._MaxComboItems=   5
         Columns(20).Caption=   "CuentaBanco"
         Columns(20).DataField=   "CuentaBanco"
         Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(21)._VlistStyle=   0
         Columns(21)._MaxComboItems=   5
         Columns(21).Caption=   "Turnos"
         Columns(21).DataField=   "Turnos"
         Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(22)._VlistStyle=   0
         Columns(22)._MaxComboItems=   5
         Columns(22).Caption=   "FechaNacimiento"
         Columns(22).DataField=   "FechaNacimiento"
         Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(23)._VlistStyle=   0
         Columns(23)._MaxComboItems=   5
         Columns(23).Caption=   "Numerocelular"
         Columns(23).DataField=   "Numerocelular"
         Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(24)._VlistStyle=   0
         Columns(24)._MaxComboItems=   5
         Columns(24).Caption=   "CelularEmergencia"
         Columns(24).DataField=   "CelularEmergencia"
         Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(25)._VlistStyle=   0
         Columns(25)._MaxComboItems=   5
         Columns(25).Caption=   "Profesion"
         Columns(25).DataField=   "Profesion"
         Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(26)._VlistStyle=   0
         Columns(26)._MaxComboItems=   5
         Columns(26).Caption=   "EstadoCivil"
         Columns(26).DataField=   "EstadoCivil"
         Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(27)._VlistStyle=   0
         Columns(27)._MaxComboItems=   5
         Columns(27).Caption=   "JefeInmediato"
         Columns(27).DataField=   "JefeInmediato"
         Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(28)._VlistStyle=   0
         Columns(28)._MaxComboItems=   5
         Columns(28).Caption=   "Incentivo"
         Columns(28).DataField=   "Incentivo"
         Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   29
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=29"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=873"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=794"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1773"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1693"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=1773"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1693"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=1773"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1693"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2646"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2566"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=1773"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=1693"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=1773"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=1693"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=1773"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=1693"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(45)=   "Column(11).Width=1773"
         Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=1693"
         Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(49)=   "Column(12).Width=1773"
         Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=1693"
         Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(53)=   "Column(13).Width=2646"
         Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=2566"
         Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(57)=   "Column(14).Width=873"
         Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=794"
         Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(61)=   "Column(15).Width=2725"
         Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=2646"
         Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(65)=   "Column(16).Width=2725"
         Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=2646"
         Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(69)=   "Column(17).Width=2725"
         Splits(0)._ColumnProps(70)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(71)=   "Column(17)._WidthInPix=2646"
         Splits(0)._ColumnProps(72)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(73)=   "Column(18).Width=2725"
         Splits(0)._ColumnProps(74)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(75)=   "Column(18)._WidthInPix=2646"
         Splits(0)._ColumnProps(76)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(77)=   "Column(19).Width=2725"
         Splits(0)._ColumnProps(78)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(79)=   "Column(19)._WidthInPix=2646"
         Splits(0)._ColumnProps(80)=   "Column(19).Order=20"
         Splits(0)._ColumnProps(81)=   "Column(20).Width=2725"
         Splits(0)._ColumnProps(82)=   "Column(20).DividerColor=0"
         Splits(0)._ColumnProps(83)=   "Column(20)._WidthInPix=2646"
         Splits(0)._ColumnProps(84)=   "Column(20).Order=21"
         Splits(0)._ColumnProps(85)=   "Column(21).Width=2725"
         Splits(0)._ColumnProps(86)=   "Column(21).DividerColor=0"
         Splits(0)._ColumnProps(87)=   "Column(21)._WidthInPix=2646"
         Splits(0)._ColumnProps(88)=   "Column(21).Order=22"
         Splits(0)._ColumnProps(89)=   "Column(22).Width=2725"
         Splits(0)._ColumnProps(90)=   "Column(22).DividerColor=0"
         Splits(0)._ColumnProps(91)=   "Column(22)._WidthInPix=2646"
         Splits(0)._ColumnProps(92)=   "Column(22).Order=23"
         Splits(0)._ColumnProps(93)=   "Column(23).Width=2725"
         Splits(0)._ColumnProps(94)=   "Column(23).DividerColor=0"
         Splits(0)._ColumnProps(95)=   "Column(23)._WidthInPix=2646"
         Splits(0)._ColumnProps(96)=   "Column(23).Order=24"
         Splits(0)._ColumnProps(97)=   "Column(24).Width=2725"
         Splits(0)._ColumnProps(98)=   "Column(24).DividerColor=0"
         Splits(0)._ColumnProps(99)=   "Column(24)._WidthInPix=2646"
         Splits(0)._ColumnProps(100)=   "Column(24).Order=25"
         Splits(0)._ColumnProps(101)=   "Column(25).Width=2725"
         Splits(0)._ColumnProps(102)=   "Column(25).DividerColor=0"
         Splits(0)._ColumnProps(103)=   "Column(25)._WidthInPix=2646"
         Splits(0)._ColumnProps(104)=   "Column(25).Order=26"
         Splits(0)._ColumnProps(105)=   "Column(26).Width=2725"
         Splits(0)._ColumnProps(106)=   "Column(26).DividerColor=0"
         Splits(0)._ColumnProps(107)=   "Column(26)._WidthInPix=2646"
         Splits(0)._ColumnProps(108)=   "Column(26).Order=27"
         Splits(0)._ColumnProps(109)=   "Column(27).Width=2725"
         Splits(0)._ColumnProps(110)=   "Column(27).DividerColor=0"
         Splits(0)._ColumnProps(111)=   "Column(27)._WidthInPix=2646"
         Splits(0)._ColumnProps(112)=   "Column(27).Order=28"
         Splits(0)._ColumnProps(113)=   "Column(28).Width=2725"
         Splits(0)._ColumnProps(114)=   "Column(28).DividerColor=0"
         Splits(0)._ColumnProps(115)=   "Column(28)._WidthInPix=2646"
         Splits(0)._ColumnProps(116)=   "Column(28).Order=29"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         DirectionAfterTab=   1
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
         _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
         _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
         _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
         _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
         _StyleDefs(78)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
         _StyleDefs(79)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
         _StyleDefs(80)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
         _StyleDefs(81)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
         _StyleDefs(82)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
         _StyleDefs(83)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
         _StyleDefs(84)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
         _StyleDefs(85)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
         _StyleDefs(86)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
         _StyleDefs(87)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
         _StyleDefs(88)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
         _StyleDefs(89)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
         _StyleDefs(90)  =   "Splits(0).Columns(15).Style:id=98,.parent=13"
         _StyleDefs(91)  =   "Splits(0).Columns(15).HeadingStyle:id=95,.parent=14"
         _StyleDefs(92)  =   "Splits(0).Columns(15).FooterStyle:id=96,.parent=15"
         _StyleDefs(93)  =   "Splits(0).Columns(15).EditorStyle:id=97,.parent=17"
         _StyleDefs(94)  =   "Splits(0).Columns(16).Style:id=102,.parent=13"
         _StyleDefs(95)  =   "Splits(0).Columns(16).HeadingStyle:id=99,.parent=14"
         _StyleDefs(96)  =   "Splits(0).Columns(16).FooterStyle:id=100,.parent=15"
         _StyleDefs(97)  =   "Splits(0).Columns(16).EditorStyle:id=101,.parent=17"
         _StyleDefs(98)  =   "Splits(0).Columns(17).Style:id=106,.parent=13"
         _StyleDefs(99)  =   "Splits(0).Columns(17).HeadingStyle:id=103,.parent=14"
         _StyleDefs(100) =   "Splits(0).Columns(17).FooterStyle:id=104,.parent=15"
         _StyleDefs(101) =   "Splits(0).Columns(17).EditorStyle:id=105,.parent=17"
         _StyleDefs(102) =   "Splits(0).Columns(18).Style:id=110,.parent=13"
         _StyleDefs(103) =   "Splits(0).Columns(18).HeadingStyle:id=107,.parent=14"
         _StyleDefs(104) =   "Splits(0).Columns(18).FooterStyle:id=108,.parent=15"
         _StyleDefs(105) =   "Splits(0).Columns(18).EditorStyle:id=109,.parent=17"
         _StyleDefs(106) =   "Splits(0).Columns(19).Style:id=114,.parent=13"
         _StyleDefs(107) =   "Splits(0).Columns(19).HeadingStyle:id=111,.parent=14"
         _StyleDefs(108) =   "Splits(0).Columns(19).FooterStyle:id=112,.parent=15"
         _StyleDefs(109) =   "Splits(0).Columns(19).EditorStyle:id=113,.parent=17"
         _StyleDefs(110) =   "Splits(0).Columns(20).Style:id=118,.parent=13"
         _StyleDefs(111) =   "Splits(0).Columns(20).HeadingStyle:id=115,.parent=14"
         _StyleDefs(112) =   "Splits(0).Columns(20).FooterStyle:id=116,.parent=15"
         _StyleDefs(113) =   "Splits(0).Columns(20).EditorStyle:id=117,.parent=17"
         _StyleDefs(114) =   "Splits(0).Columns(21).Style:id=122,.parent=13"
         _StyleDefs(115) =   "Splits(0).Columns(21).HeadingStyle:id=119,.parent=14"
         _StyleDefs(116) =   "Splits(0).Columns(21).FooterStyle:id=120,.parent=15"
         _StyleDefs(117) =   "Splits(0).Columns(21).EditorStyle:id=121,.parent=17"
         _StyleDefs(118) =   "Splits(0).Columns(22).Style:id=126,.parent=13"
         _StyleDefs(119) =   "Splits(0).Columns(22).HeadingStyle:id=123,.parent=14"
         _StyleDefs(120) =   "Splits(0).Columns(22).FooterStyle:id=124,.parent=15"
         _StyleDefs(121) =   "Splits(0).Columns(22).EditorStyle:id=125,.parent=17"
         _StyleDefs(122) =   "Splits(0).Columns(23).Style:id=130,.parent=13"
         _StyleDefs(123) =   "Splits(0).Columns(23).HeadingStyle:id=127,.parent=14"
         _StyleDefs(124) =   "Splits(0).Columns(23).FooterStyle:id=128,.parent=15"
         _StyleDefs(125) =   "Splits(0).Columns(23).EditorStyle:id=129,.parent=17"
         _StyleDefs(126) =   "Splits(0).Columns(24).Style:id=134,.parent=13"
         _StyleDefs(127) =   "Splits(0).Columns(24).HeadingStyle:id=131,.parent=14"
         _StyleDefs(128) =   "Splits(0).Columns(24).FooterStyle:id=132,.parent=15"
         _StyleDefs(129) =   "Splits(0).Columns(24).EditorStyle:id=133,.parent=17"
         _StyleDefs(130) =   "Splits(0).Columns(25).Style:id=138,.parent=13"
         _StyleDefs(131) =   "Splits(0).Columns(25).HeadingStyle:id=135,.parent=14"
         _StyleDefs(132) =   "Splits(0).Columns(25).FooterStyle:id=136,.parent=15"
         _StyleDefs(133) =   "Splits(0).Columns(25).EditorStyle:id=137,.parent=17"
         _StyleDefs(134) =   "Splits(0).Columns(26).Style:id=142,.parent=13"
         _StyleDefs(135) =   "Splits(0).Columns(26).HeadingStyle:id=139,.parent=14"
         _StyleDefs(136) =   "Splits(0).Columns(26).FooterStyle:id=140,.parent=15"
         _StyleDefs(137) =   "Splits(0).Columns(26).EditorStyle:id=141,.parent=17"
         _StyleDefs(138) =   "Splits(0).Columns(27).Style:id=146,.parent=13"
         _StyleDefs(139) =   "Splits(0).Columns(27).HeadingStyle:id=143,.parent=14"
         _StyleDefs(140) =   "Splits(0).Columns(27).FooterStyle:id=144,.parent=15"
         _StyleDefs(141) =   "Splits(0).Columns(27).EditorStyle:id=145,.parent=17"
         _StyleDefs(142) =   "Splits(0).Columns(28).Style:id=150,.parent=13"
         _StyleDefs(143) =   "Splits(0).Columns(28).HeadingStyle:id=147,.parent=14"
         _StyleDefs(144) =   "Splits(0).Columns(28).FooterStyle:id=148,.parent=15"
         _StyleDefs(145) =   "Splits(0).Columns(28).EditorStyle:id=149,.parent=17"
         _StyleDefs(146) =   "Named:id=33:Normal"
         _StyleDefs(147) =   ":id=33,.parent=0"
         _StyleDefs(148) =   "Named:id=34:Heading"
         _StyleDefs(149) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(150) =   ":id=34,.wraptext=-1"
         _StyleDefs(151) =   "Named:id=35:Footing"
         _StyleDefs(152) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(153) =   "Named:id=36:Selected"
         _StyleDefs(154) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(155) =   "Named:id=37:Caption"
         _StyleDefs(156) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(157) =   "Named:id=38:HighlightRow"
         _StyleDefs(158) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(159) =   "Named:id=39:EvenRow"
         _StyleDefs(160) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(161) =   "Named:id=40:OddRow"
         _StyleDefs(162) =   ":id=40,.parent=33"
         _StyleDefs(163) =   "Named:id=41:RecordSelector"
         _StyleDefs(164) =   ":id=41,.parent=34"
         _StyleDefs(165) =   "Named:id=42:FilterBar"
         _StyleDefs(166) =   ":id=42,.parent=33"
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   375
         Left            =   11400
         TabIndex        =   17
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   41829
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   11400
         OleObjectBlob   =   "FrmImportarEmpleados.frx":0FD6
         TabIndex        =   18
         Top             =   1680
         Width           =   1335
      End
      Begin TrueOleDBList80.TDBCombo TDBCombo1 
         Bindings        =   "FrmImportarEmpleados.frx":104C
         Height          =   315
         Left            =   -68440
         TabIndex        =   20
         Top             =   960
         Visible         =   0   'False
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         _EDITHEIGHT     =   556
         _GAPHEIGHT      =   53
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
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits.Count    =   1
         Appearance      =   1
         BorderStyle     =   1
         ComboStyle      =   0
         AutoCompletion  =   0   'False
         LimitToList     =   0   'False
         ColumnHeaders   =   -1  'True
         ColumnFooters   =   0   'False
         DataMode        =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   0
         Caption         =   ""
         EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         AutoSize        =   -1  'True
         ListField       =   "Nomina"
         BoundColumn     =   ""
         IntegralHeight  =   0   'False
         CellTipsWidth   =   0
         CellTipsDelay   =   1000
         AutoDropdown    =   0   'False
         RowTracking     =   -1  'True
         RightToLeft     =   0   'False
         RowMember       =   ""
         MouseIcon       =   0
         MouseIcon.vt    =   3
         MousePointer    =   0
         MatchEntryTimeout=   2000
         OLEDragMode     =   0
         OLEDropMode     =   0
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         AddItemSeparator=   ";"
         _PropDict       =   $"FrmImportarEmpleados.frx":1068
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(42)  =   ":id=34,.wraptext=-1"
         _StyleDefs(43)  =   "Named:id=35:Footing"
         _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(45)  =   "Named:id=36:Selected"
         _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(47)  =   "Named:id=37:Caption"
         _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(49)  =   "Named:id=38:HighlightRow"
         _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(51)  =   "Named:id=39:EvenRow"
         _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(53)  =   "Named:id=40:OddRow"
         _StyleDefs(54)  =   ":id=40,.parent=33"
         _StyleDefs(55)  =   "Named:id=41:RecordSelector"
         _StyleDefs(56)  =   ":id=41,.parent=34"
         _StyleDefs(57)  =   "Named:id=42:FilterBar"
         _StyleDefs(58)  =   ":id=42,.parent=33"
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   -69760
         OleObjectBlob   =   "FrmImportarEmpleados.frx":1112
         TabIndex        =   23
         Top             =   525
         Visible         =   0   'False
         Width           =   1335
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGrid2 
         Height          =   2895
         Left            =   -69760
         TabIndex        =   25
         Top             =   1560
         Visible         =   0   'False
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5106
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NoINSS"
         Columns(0).DataField=   "NoINSS"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "CodEmpleado"
         Columns(1).DataField=   "CodEmpleado"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nombres"
         Columns(2).DataField=   "Nombres"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Monto"
         Columns(3).DataField=   "Monto"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "TipoDeduccion"
         Columns(4).DataField=   "TipoDeduccion"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2646"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2566"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2646"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2566"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=4419"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4339"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         DirectionAfterTab=   1
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Named:id=33:Normal"
         _StyleDefs(51)  =   ":id=33,.parent=0"
         _StyleDefs(52)  =   "Named:id=34:Heading"
         _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   ":id=34,.wraptext=-1"
         _StyleDefs(55)  =   "Named:id=35:Footing"
         _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=36:Selected"
         _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=37:Caption"
         _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(61)  =   "Named:id=38:HighlightRow"
         _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(63)  =   "Named:id=39:EvenRow"
         _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(65)  =   "Named:id=40:OddRow"
         _StyleDefs(66)  =   ":id=40,.parent=33"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33"
      End
      Begin XtremeSuiteControls.ProgressBar osProgress1 
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   4440
         Width           =   11295
         _Version        =   786432
         _ExtentX        =   19923
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   14737632
         Scrolling       =   1
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   -69760
         TabIndex        =   28
         Top             =   4560
         Visible         =   0   'False
         Width           =   11295
         _Version        =   786432
         _ExtentX        =   19923
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   14737632
         Scrolling       =   1
         Appearance      =   6
      End
      Begin TrueOleDBList80.TDBCombo TDBComboIngresos 
         Bindings        =   "FrmImportarEmpleados.frx":1190
         Height          =   315
         Left            =   -68440
         TabIndex        =   57
         Top             =   960
         Visible         =   0   'False
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         _EDITHEIGHT     =   556
         _GAPHEIGHT      =   53
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
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits.Count    =   1
         Appearance      =   1
         BorderStyle     =   1
         ComboStyle      =   0
         AutoCompletion  =   0   'False
         LimitToList     =   0   'False
         ColumnHeaders   =   -1  'True
         ColumnFooters   =   0   'False
         DataMode        =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   0
         Caption         =   ""
         EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         AutoSize        =   -1  'True
         ListField       =   "Nomina"
         BoundColumn     =   ""
         IntegralHeight  =   0   'False
         CellTipsWidth   =   0
         CellTipsDelay   =   1000
         AutoDropdown    =   0   'False
         RowTracking     =   -1  'True
         RightToLeft     =   0   'False
         RowMember       =   ""
         MouseIcon       =   0
         MouseIcon.vt    =   3
         MousePointer    =   0
         MatchEntryTimeout=   2000
         OLEDragMode     =   0
         OLEDropMode     =   0
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         AddItemSeparator=   ";"
         _PropDict       =   $"FrmImportarEmpleados.frx":11AC
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(42)  =   ":id=34,.wraptext=-1"
         _StyleDefs(43)  =   "Named:id=35:Footing"
         _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(45)  =   "Named:id=36:Selected"
         _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(47)  =   "Named:id=37:Caption"
         _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(49)  =   "Named:id=38:HighlightRow"
         _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(51)  =   "Named:id=39:EvenRow"
         _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(53)  =   "Named:id=40:OddRow"
         _StyleDefs(54)  =   ":id=40,.parent=33"
         _StyleDefs(55)  =   "Named:id=41:RecordSelector"
         _StyleDefs(56)  =   ":id=41,.parent=34"
         _StyleDefs(57)  =   "Named:id=42:FilterBar"
         _StyleDefs(58)  =   ":id=42,.parent=33"
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   -69760
         OleObjectBlob   =   "FrmImportarEmpleados.frx":1256
         TabIndex        =   58
         Top             =   525
         Visible         =   0   'False
         Width           =   1335
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGridIngresos 
         Height          =   2895
         Left            =   -69760
         TabIndex        =   59
         Top             =   1560
         Visible         =   0   'False
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5106
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NoINSS"
         Columns(0).DataField=   "NoINSS"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "CodEmpleado"
         Columns(1).DataField=   "CodEmpleado"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nombres"
         Columns(2).DataField=   "Nombres"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Monto"
         Columns(3).DataField=   "Monto"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "TipoIncentivo"
         Columns(4).DataField=   "TipoIncentivo"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2646"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2566"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2646"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2566"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=4419"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4339"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         DirectionAfterTab=   1
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Named:id=33:Normal"
         _StyleDefs(51)  =   ":id=33,.parent=0"
         _StyleDefs(52)  =   "Named:id=34:Heading"
         _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   ":id=34,.wraptext=-1"
         _StyleDefs(55)  =   "Named:id=35:Footing"
         _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=36:Selected"
         _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=37:Caption"
         _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(61)  =   "Named:id=38:HighlightRow"
         _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(63)  =   "Named:id=39:EvenRow"
         _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(65)  =   "Named:id=40:OddRow"
         _StyleDefs(66)  =   ":id=40,.parent=33"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33"
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBar2 
         Height          =   375
         Left            =   -69760
         TabIndex        =   60
         Top             =   4560
         Visible         =   0   'False
         Width           =   11295
         _Version        =   786432
         _ExtentX        =   19923
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   14737632
         Scrolling       =   1
         Appearance      =   6
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Nóminas:"
         Height          =   255
         Left            =   -69520
         TabIndex        =   61
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Importando Horas Extra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69040
         TabIndex        =   51
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   -61960
         TabIndex        =   48
         Top             =   1320
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblCodigo 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   -61960
         TabIndex        =   47
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "No existen o inactivos"
         Height          =   255
         Left            =   -58840
         TabIndex        =   43
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Movimiento Salarial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69280
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label txtNombre 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   -61720
         TabIndex        =   37
         Top             =   1200
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label txtCodigo 
         Caption         =   "Codigo:"
         Height          =   255
         Left            =   -61720
         TabIndex        =   36
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Importando saldo de vacaciones al:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68560
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Nóminas:"
         Height          =   255
         Left            =   -69520
         TabIndex        =   24
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label34 
         Caption         =   "Tipo Nóminas:"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5C1A1&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   13335
      TabIndex        =   0
      Top             =   -120
      Width           =   13335
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "IMPORTAR REGISTROS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   4200
         TabIndex        =   1
         Top             =   360
         Width           =   3840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   13200
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   240
         Picture         =   "FrmImportarEmpleados.frx":12D4
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1200
      End
   End
   Begin MSComDlg.CommonDialog CMRutaFoto 
      Left            =   120
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   256
   End
   Begin MSAdodcLib.Adodc AdoRegistros 
      Height          =   330
      Left            =   0
      Top             =   7200
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "AdoRegistros"
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
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   330
      Left            =   3840
      Top             =   8040
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "AdoConsulta"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3840
      Top             =   7680
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "AdoRegistros"
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
      Left            =   3360
      Top             =   7440
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
   Begin MSAdodcLib.Adodc DtaHorarioEmpleado 
      Height          =   375
      Left            =   0
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
      Left            =   0
      Top             =   8280
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
   Begin MSAdodcLib.Adodc DtaHistorico 
      Height          =   375
      Left            =   0
      Top             =   7560
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
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   330
      Left            =   3240
      Top             =   8400
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
   Begin MSAdodcLib.Adodc DtaConsecutivos 
      Height          =   330
      Left            =   8760
      Top             =   7080
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
   Begin MSAdodcLib.Adodc DtaNomina 
      Height          =   330
      Left            =   8760
      Top             =   7440
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
   Begin MSAdodcLib.Adodc DtaTipoNomina 
      Height          =   330
      Left            =   8760
      Top             =   7800
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
   Begin MSAdodcLib.Adodc DtaDetalleNomina 
      Height          =   330
      Left            =   8640
      Top             =   7440
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
   Begin MSAdodcLib.Adodc DtaDeduccion 
      Height          =   375
      Left            =   8640
      Top             =   8280
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
   Begin MSAdodcLib.Adodc DtaDetalleDeduccion2 
      Height          =   330
      Left            =   9480
      Top             =   7440
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "DtaDetalleDeduccion2"
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
   Begin MSAdodcLib.Adodc AdoUserInfo 
      Height          =   375
      Left            =   6240
      Top             =   7800
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
      Caption         =   "AdoUserInfo"
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
   Begin MSAdodcLib.Adodc DtaIncentivo 
      Height          =   330
      Left            =   3960
      Top             =   6960
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "AdoIncentivos"
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
      Height          =   330
      Left            =   1080
      Top             =   6840
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
Attribute VB_Name = "FrmImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExcelHE_Click()
Dim retval
Dim OPENFILENAME As String, Directorio As String
Dim Rango As String, Hoja As String, ruta As String

    On Error Resume Next
  
    dialogSaldos.FileName = ""
    dialogSaldos.Filter = "Archivo xls |*.xls"
    ' Display common dialog box
    dialogSaldos.ShowOpen
    Dim RutaArchivo As String
    RutaArchivo = dialogSaldos.FileName
    
  
    ruta = RutaArchivo 'ruta del archivo excel
    Rango = "A1:G8"    'Text2 & ":" & Text3 'Rango de datos (opcional)
    Hoja = "Hoja1" 'Nombre de la hoja
'    ruta = "C:\"
'    Set Me.TDBGrid1.DataSource = LeerTxt(ruta)
    
    Set Me.gridHorasExtra.DataSource = Leer_Excel(ruta, "Hoja1")
   ' Set Me.TDBGridNominas.DataSource = Leer_Excel(ruta, "Hoja1")
End Sub

Private Sub btnImportarHE_Click()
Dim CodigoEmpleado As String
Dim CodigoEmpleadoAutonumerico As Integer
Dim ConsHE As Integer
Dim SQLHorasExtras As String
Dim CantidadHoras As Double
Dim NombreEmpleado As String
Dim TipoNomina As String
Dim SQLNomina As String
Dim NumeroNomina As Integer
Dim Consecutivo As Integer

    With Me.DtaHrasExtras
       '.DatabaseName = Ruta
       .ConnectionString = Conexion
    End With

     gridHorasExtra.MoveFirst
     Do While Not gridHorasExtra.EOF
        
        CodigoEmpleado = gridHorasExtra.Columns(0).Text
        NombreEmpleado = gridHorasExtra.Columns(1).Text
        CantidadHoras = CDbl(gridHorasExtra.Columns(2).Text)  '//// SACO LOS DATOS DEL GRID
        
        Me.lblCodigo.Caption = CodigoEmpleado
        Me.lblNombre.Caption = NombreEmpleado
        
        AdoConsulta.RecordSource = "SELECT    CodEmpleado, CodTipoNomina AS TipoNomina FROM         Empleado WHERE     (CodEmpleado1 = '" & CodigoEmpleado & "') AND (Activo = 'True')"
        AdoConsulta.Refresh
        
        If AdoConsulta.Recordset.EOF Then
            MsgBox ("El Empleado con el codigo " & CodigoEmpleado & " no existe o esta inactivo")
            Me.listHE.AddItem (CodigoEmpleado & ", Inactivo o no existe")
            Exit Sub
        Else
            TipoNomina = AdoConsulta.Recordset("TipoNomina")
            CodigoEmpleadoAutonumerico = AdoConsulta.Recordset("CodEmpleado")
        End If
        
        
        
        SQLNomina = "SELECT Nomina.*, TipoNomina.TipoPago FROM TipoNomina INNER JOIN Nomina ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina Where (((Nomina.Activa) = 1)) And Nomina.CodTipoNomina = '" & TipoNomina & "' "
        DtaNomina.RecordSource = SQLNomina
        DtaNomina.Refresh                 '/////// SACO EL NUMEROM DE LA NOMINA
        
        If DtaNomina.Recordset.EOF Then
            MsgBox ("No Hay nomina activa para el empleado con el codigo: " & CodigoEmpleado & "")
            Me.listHE.AddItem (CodigoEmpleado & ", Nomina Inactiva")
            Exit Sub
        Else
            NumeroNomina = DtaNomina.Recordset("NumNomina")
        End If
        
        'busco en Hrs Extras si ya le fue gravada una hora extra
        SQLHorasExtras = "SELECT HorasExtras.CodEmpleado, HorasExtras.NumNomina, HorasExtras.CantHoras, HorasExtras.Pagada From HorasExtras WHERE HorasExtras.CodEmpleado=" & CodigoEmpleadoAutonumerico & " AND HorasExtras.NumNomina= " & NumeroNomina & ""
        DtaHrasExtras.RecordSource = SQLHorasExtras
        
        DtaHrasExtras.Refresh
        If Not DtaHrasExtras.Recordset.EOF Then
           MsgBox "Ya le fueron agregadas horas extras a este empleado, las horas anteriores serán reemplazadas"
           DtaHrasExtras.Recordset.Fields("canthoras") = CantidadHoras
           DtaHrasExtras.Recordset.Fields("Pagada") = 0
           DtaHrasExtras.Recordset.Update
           DtaHrasExtras.Refresh
           Me.listHE.AddItem (CodigoEmpleado & ", horas actualizadas")
        Else
        
            'Si no las encontro grabo las horas extras
            Me.DtaHrasExtras.RecordSource = "HorasExtras"
            Me.DtaHrasExtras.Refresh
             If Not Me.DtaHrasExtras.Recordset.EOF Then
               Me.DtaHrasExtras.Recordset.MoveLast
               Consecutivo = Me.DtaHrasExtras.Recordset("ID") + 1
             Else
               Consecutivo = 1
             End If
            
              Me.DtaHrasExtras.Recordset.AddNew
              DtaHrasExtras.Recordset.Fields("Id") = Consecutivo
              DtaHrasExtras.Recordset.Fields("CodEmpleado") = CodigoEmpleadoAutonumerico
              DtaHrasExtras.Recordset.Fields("NumNomina") = NumeroNomina
              DtaHrasExtras.Recordset.Fields("canthoras") = CantidadHoras
              DtaHrasExtras.Recordset.Fields("Pagada") = 0
              DtaHrasExtras.Recordset.Update
            
        End If
       
     gridHorasExtra.MoveNext
     Loop
        
        MsgBox ("Importado con exito")
End Sub

Private Sub btnIniciar_Click()
Dim retval
Dim OPENFILENAME As String, Directorio As String
Dim Rango As String, Hoja As String, ruta As String

    On Error Resume Next
  
    dialogSaldos.FileName = ""
    dialogSaldos.Filter = "Archivo xls |*.xls"
    ' Display common dialog box
    dialogSaldos.ShowOpen
    Dim RutaArchivo As String
    RutaArchivo = dialogSaldos.FileName
    
  
    ruta = RutaArchivo 'ruta del archivo excel
    Rango = "A1:G8"    'Text2 & ":" & Text3 'Rango de datos (opcional)
    Hoja = "Hoja1" 'Nombre de la hoja
'    ruta = "C:\"
'    Set Me.TDBGrid1.DataSource = LeerTxt(ruta)
    
    Set Me.gridSaldos.DataSource = Leer_Excel(ruta, "Hoja1")
   ' Set Me.TDBGridNominas.DataSource = Leer_Excel(ruta, "Hoja1")
    
End Sub

Private Sub CmdBuscarLogo_Click()
Dim retval
Dim OPENFILENAME As String, Directorio As String
Dim Rango As String, Hoja As String, ruta As String

    On Error Resume Next
    ' Set the commom dialog properties we need
    If Me.TxtRutaLogo.Text <> "" Then
       CMRutaFoto.InitDir = Me.TxtRutaLogo.Text
    End If
    CMRutaFoto.FileName = ""
    ' We will load BMP, JPG, and TIF files
    
    CMRutaFoto.Filter = "Archivo xls |*.xls"
    ' Display common dialog box
    CMRutaFoto.ShowOpen
    Me.TxtRutaLogo.Text = CMRutaFoto.FileName
   
    
  
    ruta = Me.TxtRutaLogo.Text 'ruta del archivo excel
    Rango = "A1:G8"    'Text2 & ":" & Text3 'Rango de datos (opcional)
    Hoja = "Hoja1" 'Nombre de la hoja
'    ruta = "C:\"
'    Set Me.TDBGrid1.DataSource = LeerTxt(ruta)
    
    Set Me.TDBGrid1.DataSource = Leer_Excel(ruta, "Hoja1")
    Set Me.TDBGridNominas.DataSource = Leer_Excel(ruta, "Hoja1")
    
End Sub

Private Sub CmdIniciar_Click()
 Dim sql As String, CodigoEmpleado As Double, Dia As Date, Hora As String, Fecha As Date, Fecha2 As Date, FechaNacimiento As Date
  Dim Nombres  As String, Nombre1 As String, Nombre2 As String, Apellido As String, Apellido2 As String, Direccion As String
  Dim Nacionalidad As String, NumeroCedula As String, Sexo As String, NumeroInss As String, NumeroRuc As String, CuentaBanco As String
  Dim NHijos As Double, CodDepartamento As String, departamento As String, CodCargo As String, Cargo As String
  Dim CodGrupo As String, DescripcionGrupo As String, CodTipoNomina As String, DescripcionTipo As String
  Dim CodTurno As String, CodigoEmpleado1 As String, SueldoPeriodo As Double, TarifaHorario As Double
  Dim Id As Double, IdCard As String, NumeroCelular As String, NumeroCelularEmergencia As String, Turno As String
  Dim rs As New ADODB.Recordset, Profesion As String, EstadoCivil As String, JefeInmediato As String
  Dim Incentivo As Double
  
          Me.TDBGrid1.MoveFirst
          

          Do While Not Me.TDBGrid1.EOF
          
            '--------------------CARGO LAS VARIABLES ------------------------------------------
            

            
            CodigoEmpleado1 = Format(Me.TDBGrid1.Columns("Codigo").Text, "000#")
            Nombre1 = Me.TDBGrid1.Columns("Nombre1").Text
            Nombre2 = Me.TDBGrid1.Columns("Nombre2").Text
            Apellido = Me.TDBGrid1.Columns("Apellido1").Text
            Apellido2 = Me.TDBGrid1.Columns("Apellido2").Text
            Direccion = Me.TDBGrid1.Columns("Direccion").Text
            Nacionalidad = Me.TDBGrid1.Columns("Nacionalidad").Text
            Sexo = Me.TDBGrid1.Columns("Sexo").Text
            If Me.TDBGrid1.Columns("FechaIngreso").Text = "" Then
              Fecha = Format(Now, "dd-MM-yyyy")
            Else
             Fecha = Me.TDBGrid1.Columns("FechaIngreso").Text
            End If
            Cargo = Me.TDBGrid1.Columns("Cargo").Text
            departamento = Me.TDBGrid1.Columns("Departamento").Text
            NumeroRuc = Me.TDBGrid1.Columns("RUC").Text
            NumeroInss = Me.TDBGrid1.Columns("INSS").Text
            NumeroCedula = Me.TDBGrid1.Columns("Cedula").Text
            If Me.TDBGrid1.Columns("Hijos").Text <> "" Then
              NHijos = Me.TDBGrid1.Columns("Hijos").Text
            End If
            DescripcionTipo = Me.TDBGrid1.Columns("TipoNomina").Text
            DescripcionGrupo = Me.TDBGrid1.Columns("GrupoNomina").Text
            
           
            
            
            EstadoCivil = Me.TDBGrid1.Columns("EstadoCivil").Text
            Incentivo = Me.TDBGrid1.Columns("Incentivo").Text
            Turno = Me.TDBGrid1.Columns("Turnos").Text
            Profesion = Me.TDBGrid1.Columns("Profesion").Text
            EstadoCivil = Me.TDBGrid1.Columns("EstadoCivil").Text
            JefeInmediato = Me.TDBGrid1.Columns("JefeInmediato").Text
            NumeroCelular = Me.TDBGrid1.Columns("Numerocelular").Text
            NumeroCelularEmergencia = Me.TDBGrid1.Columns("CelularEmergencia").Text
            
            If Me.TDBGrid1.Columns("SueldoPeriodo").Text = "" Then
                SueldoPeriodo = 0
            Else
                SueldoPeriodo = Me.TDBGrid1.Columns("SueldoPeriodo").Text
            End If
            
            If Me.TDBGrid1.Columns("TarifaHoraria").Text = "" Then
                TarifaHorario = 0
            Else
                TarifaHorario = Me.TDBGrid1.Columns("TarifaHoraria").Text
            End If
            
            If Not IsDate(Me.TDBGrid1.Columns("FechaNacimiento").Text) Then
                FechaNacimiento = Format(Now, "dd-MM-yyyy")
            Else
                FechaNacimiento = Me.TDBGrid1.Columns("FechaNacimiento").Text
            End If
            
            CuentaBanco = Me.TDBGrid1.Columns("CuentaBanco").Text
            

            
            
            
            
            Nombres = Nombre1 & " " & Nombre2 & " " & Apellido & " " & Apellido2
            
            
            
            
            
            CodTipoNomina = BuscaCodigo(DescripcionTipo, "TipoNomina", "CodTipoNomina", "Nomina")
            
            CodDepartamento = BuscaCodigo(departamento, "Departamento", "CodDepartamento", "Departamento")
            If CodDepartamento = "00" Then
              rs.Open "INSERT INTO Departamento ([CodDepartamento] ,[Departamento]) Values ('" & UltimoCodigo & "'  ,'" & departamento & "')", Conexion
             CodDepartamento = UltimoCodigo
            End If
            
            CodCargo = BuscaCodigo(Cargo, "Cargo", "CodCargo", "Cargo")
            If CodCargo = "00" Then
              rs.Open "INSERT INTO Cargo ([CodCargo] ,[Cargo]) Values ('" & UltimoCodigo & "'  ,'" & Cargo & "')", Conexion
              CodCargo = UltimoCodigo
            End If
            
            CodGrupo = BuscaCodigo(DescripcionGrupo, "Grupo", "CodGrupo", "Grupo")
            If CodGrupo = "00" Then
              rs.Open "INSERT INTO Grupo ([CodGrupo] ,[Grupo]) Values ('" & UltimoCodigo & "'  ,'" & DescripcionGrupo & "')", Conexion
              CodGrupo = UltimoCodigo
            End If
            
            '----------------------------------------------------------------------------------------------
            '-----------------------------------GRABO DATOS GENERALES DEL EMPLEADO ------------------------
            '----------------------------------------------------------------------------------------------
      
         Me.AdoConsulta.RecordSource = "SELECT  Empleado.* From Empleado WHERE  (CodEmpleado1 = '" & CodigoEmpleado1 & "')"
         Me.AdoConsulta.Refresh
         If Me.AdoConsulta.Recordset.EOF Then
         

                    DtaEmpleado.Recordset.AddNew
                        DtaEmpleado.Recordset("CodEmpleado1") = CodigoEmpleado1
                        DtaEmpleado.Recordset("Nombre1") = Nombre1
                        DtaEmpleado.Recordset("Nombre2") = Nombre2
                        DtaEmpleado.Recordset("Apellido1") = Apellido
                        DtaEmpleado.Recordset("Apellido2") = Apellido2
                        DtaEmpleado.Recordset("Direccion") = Direccion
                        DtaEmpleado.Recordset("numcedula") = NumeroCedula
                        DtaEmpleado.Recordset("sexo") = Sexo
                        DtaEmpleado.Recordset("NumeroInss") = NumeroInss
                        DtaEmpleado.Recordset("numeroruc") = NumeroRuc
                        DtaEmpleado.Recordset("Nacionalidad") = Nacionalidad
                        If CodDepartamento <> "" Then
                         DtaEmpleado.Recordset("CodDepartamento") = CodDepartamento
                        End If
                        If CodCargo <> "" Then
                          DtaEmpleado.Recordset("CodCargo") = CodCargo
                        End If
                        DtaEmpleado.Recordset("Codgrupo") = CodGrupo
                        DtaEmpleado.Recordset("Sindicalista") = "No"
                        DtaEmpleado.Recordset("CodTipoNomina") = CodTipoNomina
                        DtaEmpleado.Recordset("numhijos") = NHijos
                        DtaEmpleado.Recordset("CuentaBanco") = CuentaBanco
                        
                        DtaEmpleado.Recordset("Turno") = Turno
                        DtaEmpleado.Recordset("Numerocelular") = NumeroCelular
                        DtaEmpleado.Recordset("CelularEmergencia") = NumeroCelularEmergencia
                        DtaEmpleado.Recordset("Profesion") = Profesion
                        DtaEmpleado.Recordset("EstadoCivil") = EstadoCivil
                        DtaEmpleado.Recordset("JefeInmediato") = JefeInmediato
                        DtaEmpleado.Recordset("Incentivo") = Incentivo
            
                        'grabar los nuevos datos de la nómina
                          DtaEmpleado.Recordset("SueldoPeriodo") = Format(SueldoPeriodo, "##,##0.00")
                          DtaEmpleado.Recordset("TarifaHoraria") = Format(TarifaHorario, "##,##0.00")
                          DtaEmpleado.Recordset("PorcentajeComision") = CDbl(0)
                          DtaEmpleado.Recordset("Dolarizado") = False
                          DtaEmpleado.Recordset("PorcientoIncentivo") = 20
                          DtaEmpleado.Recordset("salariominimo") = False
                          DtaEmpleado.Recordset("ExentoInss") = False
                          DtaEmpleado.Recordset("ExentoIr") = False
                          DtaEmpleado.Recordset("PagoInssPatronal") = True
                          
                     DtaEmpleado.Recordset.Update
                     
                     If IsNumeric(CodigoEmpleado1) Then
                       CodigoEmpleado = CodigoEmpleado1
                      
                     Else
                       Me.AdoUserInfo.Refresh
                       If Not Me.AdoUserInfo.Recordset.EOF Then
                         Me.AdoUserInfo.Recordset.MoveLast
                         CodigoEmpleado = Me.AdoUserInfo.Recordset("Userid") + 1
                       End If
                     End If
                     
                     IdCard = CodigoEmpleado1
                     
                      Me.AdoConsulta.RecordSource = "SELECT   * From Userinfo WHERE  (Userid = '" & CodigoEmpleado & "')"
                      Me.AdoConsulta.Refresh
                      If Me.AdoConsulta.Recordset.EOF Then
                            Me.AdoUserInfo.Recordset.AddNew
                            Me.AdoUserInfo.Recordset("Userid") = CodigoEmpleado
                            Me.AdoUserInfo.Recordset("Name") = Nombre1 & " " & Nombre2 & " " & Apellido & " " & Apellido2
                            Me.AdoUserInfo.Recordset("Sex") = Sexo
                            Me.AdoUserInfo.Recordset("IDCard") = IdCard
                            Me.AdoUserInfo.Recordset.Update
                      Else
                            Me.AdoConsulta.Recordset("Userid") = CodigoEmpleado
                            Me.AdoConsulta.Recordset("Name") = Nombre1 & " " & Nombre2 & " " & Apellido & " " & Apellido2
                            Me.AdoConsulta.Recordset("Sex") = Sexo
                            Me.AdoConsulta.Recordset("IDCard") = IdCard
                            Me.AdoConsulta.Recordset.Update
                      
                      End If
                     
                     
        
                     
                     
                     
                     
                         CodigoEmpleado1 = DtaEmpleado.Recordset("CodEmpleado1")
                         CodigoEmpleado = DtaEmpleado.Recordset("CodEmpleado")
                     
                       Me.DtaHorarioEmpleado.RecordSource = "SELECT CodEmpleado, LEntrada, LSalida, MEntrada, MSalida, MCEntrada, MCSalida, JEntrada, JSalida, VEntrada, VSalida, TComida, TurnoLunes,TurnoMartes , TurnoMiercoles, TurnoJueves, TurnoViernes, TurnoSabado, TurnoDomingo, SEntrada, SSalida, DEntrada, DSalida From dbo.HorarioEmpleado WHERE(CodEmpleado = '" & CodigoEmpleado & "')"
                       Me.DtaHorarioEmpleado.Refresh
                       If Me.DtaHorarioEmpleado.Recordset.EOF Then
                         Me.DtaTurnos.Refresh
                         If Not Me.DtaTurnos.Recordset.EOF Then
                           CodTurno = Me.DtaTurnos.Recordset("CodTurno")
                         Me.DtaHorarioEmpleado.Recordset.AddNew
                           Me.DtaHorarioEmpleado.Recordset("CodEmpleado") = CodigoEmpleado
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
                           Me.DtaHorarioEmpleado.Recordset("TurnoLunes") = CodTurno
                           Me.DtaHorarioEmpleado.Recordset("TurnoMartes") = CodTurno
                           Me.DtaHorarioEmpleado.Recordset("TurnoMiercoles") = CodTurno
                           Me.DtaHorarioEmpleado.Recordset("TurnoJueves") = CodTurno
                           Me.DtaHorarioEmpleado.Recordset("TurnoViernes") = CodTurno
                           Me.DtaHorarioEmpleado.Recordset("TurnoSabado") = CodTurno
                           Me.DtaHorarioEmpleado.Recordset("TurnoDomingo") = CodTurno
                           Me.DtaHorarioEmpleado.Recordset("SEntrada") = Me.DtaTurnos.Recordset("SEntrada")
                           Me.DtaHorarioEmpleado.Recordset("SSalida") = Me.DtaTurnos.Recordset("SEntrada")
                           Me.DtaHorarioEmpleado.Recordset("DEntrada") = Me.DtaTurnos.Recordset("SEntrada")
                           Me.DtaHorarioEmpleado.Recordset("DSalida") = Me.DtaTurnos.Recordset("SEntrada")
                    
                         Me.DtaHorarioEmpleado.Recordset.Update
                         End If
                       End If
        
                          Me.DtaHistorico.RecordSource = "SELECT Id, Codempleado, FechaBaja, MotivoBaja, FechaAumento, MotivoAumento, FechaInicialSusp, FechaFinalSusp, MotivoSuspencion, FechaNacimiento,FechaContrato,FechaContratoVac , CargoInicial, CargoActual, CargoAnterior, SueldoInicial, SueldoAnterior, SueldoActual, CuentaDebito, CuentaCredito,CuentaPrestamo,CuentaOtrosIngresos,CuentaINSS,CuentaIR From Historico WHERE (Codempleado = " & CodigoEmpleado & ")"
                          DtaHistorico.Refresh
                          If Not Me.DtaHistorico.Recordset.EOF Then
                    
                          
                               If DtaHistorico.Recordset("CodEmpleado") = CodigoEmpleado Then
                    '                Historico = False
                    
                                    DtaHistorico.Recordset("CodEmpleado") = CodigoEmpleado
                    '                DtaHistorico.Recordset.Fields("FechaNacimiento") = FechaNacimiento
                                    DtaHistorico.Recordset.Fields("FechaContrato") = Format(Fecha, "dd/MM/yyyy")
                                    DtaHistorico.Recordset.Fields("FechaContratoVac") = Format(Fecha, "dd/MM/yyyy")
                                    DtaHistorico.Recordset.Update
                                    Valida = 1
                                End If
                     
                           Else
                         
                                        Me.DtaConsulta.RecordSource = "SELECT Id From Historico"
                                        Me.DtaConsulta.Refresh
                                        If DtaConsulta.Recordset.EOF Then
                                          Id = 1
                                        Else
                                          Me.DtaConsulta.Recordset.MoveLast
                                          Id = Me.DtaConsulta.Recordset("id") + 1
                                        End If
                    
                                        DtaHistorico.Recordset.AddNew
                                              Me.DtaHistorico.Recordset("id") = Id
                                              DtaHistorico.Recordset("CodEmpleado") = CodigoEmpleado
                    '                          DtaHistorico.Recordset.Fields("FechaNacimiento") = FechaNacimiento
                                              DtaHistorico.Recordset.Fields("FechaContrato") = Format(Fecha, "dd/MM/yyyy")
                                              DtaHistorico.Recordset.Fields("FechaContratoVac") = Format(Fecha, "dd/MM/yyyy")
                                        DtaHistorico.Recordset.Update
                    
                               
                           End If

                Else

                        AdoConsulta.Recordset("Nombre1") = Nombre1
                        AdoConsulta.Recordset("Nombre2") = Nombre2
                        AdoConsulta.Recordset("Apellido1") = Apellido
                        AdoConsulta.Recordset("Apellido2") = Apellido2
                        AdoConsulta.Recordset("Direccion") = Direccion
                        AdoConsulta.Recordset("numcedula") = NumeroCedula
                        AdoConsulta.Recordset("sexo") = Sexo
                        AdoConsulta.Recordset("NumeroInss") = NumeroInss
                        AdoConsulta.Recordset("numeroruc") = NumeroRuc
                        If CodDepartamento <> "" Then
                         AdoConsulta.Recordset("CodDepartamento") = CodDepartamento
                        End If
                        If CodCargo <> "" Then
                          AdoConsulta.Recordset("CodCargo") = CodCargo
                        End If
                        DtaEmpleado.Recordset("Nacionalidad") = Nacionalidad
                        AdoConsulta.Recordset("Codgrupo") = CodGrupo
                        AdoConsulta.Recordset("Sindicalista") = "No"
                        AdoConsulta.Recordset("CodTipoNomina") = CodTipoNomina
                        AdoConsulta.Recordset("numhijos") = NHijos
                        AdoConsulta.Recordset("CuentaBanco") = CuentaBanco
            
                        'grabar los nuevos datos de la nómina
                          AdoConsulta.Recordset("SueldoPeriodo") = Format(SueldoPeriodo, "##,##0.00")
                          AdoConsulta.Recordset("TarifaHoraria") = Format(TarifaHorario, "##,##0.00")
                          AdoConsulta.Recordset("PorcentajeComision") = CDbl(0)
                          AdoConsulta.Recordset("Dolarizado") = False
                          AdoConsulta.Recordset("PorcientoIncentivo") = 20
                          AdoConsulta.Recordset("salariominimo") = False
                          AdoConsulta.Recordset("ExentoInss") = False
                          AdoConsulta.Recordset("ExentoIr") = False
                          AdoConsulta.Recordset("PagoInssPatronal") = True
                          
                     AdoConsulta.Recordset.Update
                        
                End If
             Me.Caption = "Procesando " & Nombres
             DoEvents
             Me.TDBGrid1.MoveNext
          Loop
End Sub

Private Sub CmdIniciar2_Click()
Dim CodTipoNomina As String, NumeroNomina As Double, CodEmpleado As Double

DtaNomina.RecordSource = "Nomina"
DtaNomina.Refresh

DtaConsecutivos.Refresh

CodTipoNomina = Me.TDBTipo.Columns(0).Text

DtaNomina.Recordset.AddNew
NumeroNomina = DtaConsecutivos.Recordset("nominas")
DtaNomina.Recordset("NumNomina") = DtaConsecutivos.Recordset("nominas")
DtaNomina.Recordset("CodTipoNomina") = CodTipoNomina
DtaNomina.Recordset("FechaNomina") = Format(CDate(Me.DtpFechaFin.Value), "DD/MM/YYYY")
DtaNomina.Recordset("FechaNominaINI") = Format(CDate(Me.DTPFechaIni.Value), "DD/MM/YYYY")
DtaNomina.Recordset("Activa") = 0
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
DtaNomina.Recordset("Anulada") = 0
DtaNomina.Recordset("Cerrada") = 0
DtaNomina.Recordset("Procesada") = 0
DtaNomina.Recordset("Mes") = Month(Me.DtpFechaFin.Value)
DtaNomina.Recordset("Ano") = Year(Me.DtpFechaFin.Value)
DtaNomina.Recordset.Update


DtaConsecutivos.Refresh
DtaConsecutivos.Recordset("nominas") = DtaConsecutivos.Recordset("nominas") + 1
DtaConsecutivos.Recordset.Update

  DtaDetalleNomina.RecordSource = "SELECT DetalleNomina.id, DetalleNomina.BonoProduccion ,DetalleNomina.IncetivoProduccion,DetalleNomina.NumNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HE, DetalleNomina.DD, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.INATEC, DetalleNomina.Mes13, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.TotalSubsidio, DetalleNomina.VacacionesPagadas, DetalleNomina.DiasVacaciones,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.TarifaHoraria,DetalleNomina.produjo,DetalleNomina.AjusteINSS,HTurno, HorasTurno,Antiguedad, AñoAntiguedad  From DetalleNomina"
  DtaDetalleNomina.Refresh


   Me.TDBGridNominas.MoveFirst
          

   Do While Not Me.TDBGridNominas.EOF
   
       CodEmpleado = BuscaCodigoInterno(Me.TDBGridNominas.Columns(0).Text)
   
       If CodEmpleado <> -1 Then
        DtaDetalleNomina.Recordset.AddNew
        DtaDetalleNomina.Recordset("CodEmpleado") = CodEmpleado
        DtaDetalleNomina.Recordset("NumNomina") = NumeroNomina
        DtaDetalleNomina.Recordset("SeptimoDia") = Me.TDBGridNominas.Columns(8).Text
        DtaDetalleNomina.Recordset("BonoProduccion") = 0
        DtaDetalleNomina.Recordset("TarifaHoraria") = 0
        DtaDetalleNomina.Recordset("IncetivoProduccion") = Me.TDBGridNominas.Columns(6).Text
        DtaDetalleNomina.Recordset("HTrabajada") = 0
        DtaDetalleNomina.Recordset("produjo") = "N"
        DtaDetalleNomina.Recordset("SalarioBasico") = Me.TDBGridNominas.Columns(2).Text
        DtaDetalleNomina.Recordset("destajo") = Me.TDBGridNominas.Columns(3).Text
        DtaDetalleNomina.Recordset("HE") = 0
        DtaDetalleNomina.Recordset("HorasExtras") = Me.TDBGridNominas.Columns(4).Text
        DtaDetalleNomina.Recordset("Comisiones") = Me.TDBGridNominas.Columns(5).Text
        DtaDetalleNomina.Recordset("incentivos") = Me.TDBGridNominas.Columns(6).Text
        DtaDetalleNomina.Recordset("OtrosIngresos") = Me.TDBGridNominas.Columns(10).Text
        DtaDetalleNomina.Recordset("DescripOtrIngre") = "xx"
        DtaDetalleNomina.Recordset("Deducciones") = Me.TDBGridNominas.Columns(15).Text
        DtaDetalleNomina.Recordset("Prestamo") = Me.TDBGridNominas.Columns(12).Text
        DtaDetalleNomina.Recordset("MontoInss") = Me.TDBGridNominas.Columns(13).Text
        DtaDetalleNomina.Recordset("MontoIR") = Me.TDBGridNominas.Columns(14).Text
        DtaDetalleNomina.Recordset("Vacaciones") = Me.TDBGridNominas.Columns(7).Text
        DtaDetalleNomina.Recordset("Mes13") = 0
        DtaDetalleNomina.Recordset("INSSPatronal") = Me.TDBGridNominas.Columns(16).Text
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
        DtaDetalleNomina.Recordset.Update
      End If
   
      Me.Caption = "Procesando " & Me.TDBGridNominas.Columns(1).Text
      DoEvents
      Me.TDBGridNominas.MoveNext
   Loop
   
   
   Me.DtaDetalleNomina.RecordSource = "SELECT SUM(SalarioBasico) AS SalarioBasico, SUM(Destajo) AS Destajo, SUM(HorasExtras) AS HorasExtras, SUM(Comisiones) AS Comisiones, SUM(OtrosIngresos) AS OtrosIngresos, SUM(Incentivos) AS Incentivos, SUM(Deducciones) AS Deducciones, SUM(Prestamo) AS Prestamo, SUM(MontoINSS) AS MontoINSS, SUM(MontoIR) AS MontoIR, SUM(VacacionesPagadas) AS VacacionesPagadas, SUM(SeptimoDia) AS SeptimoDia, SUM(IncetivoProduccion) AS IncetivoProduccion From DetalleNomina Where (NumNomina = " & NumeroNomina & " )"
   Me.DtaDetalleNomina.Refresh
   If Not Me.DtaDetalleNomina.Recordset.EOF Then
   

    DtaNomina.RecordSource = "SELECT NumNomina, TotalSalarioBasico, TotalDestajo, TotalHorasExtras, TotalComisiones, TotalIncentivos, TotalDeducciones, TotalOtrosIngresos, TotalPrestamo, TotalMontoINSS , TotalMontoIR, TotalVacaciones, TotalINSSPatronal From Nomina Where (NumNomina = " & NumeroNomina & ")"
    DtaNomina.Refresh
      If Not Me.DtaNomina.Recordset.EOF Then
        DtaNomina.Recordset("TotalSalarioBasico") = DtaDetalleNomina.Recordset("SalarioBasico")
        DtaNomina.Recordset("TotalDestajo") = DtaDetalleNomina.Recordset("destajo")
        DtaNomina.Recordset("TotalHorasExtras") = DtaDetalleNomina.Recordset("HorasExtras")
        DtaNomina.Recordset("TotalComisiones") = DtaDetalleNomina.Recordset("Comisiones")
        DtaNomina.Recordset("TotalIncentivos") = DtaDetalleNomina.Recordset("incentivos")
        DtaNomina.Recordset("TotalOtrosIngresos") = DtaDetalleNomina.Recordset("OtrosIngresos")
        DtaNomina.Recordset("TotalDeducciones") = DtaDetalleNomina.Recordset("Deducciones")
        DtaNomina.Recordset("TotalPrestamo") = DtaDetalleNomina.Recordset("Prestamo")
        DtaNomina.Recordset("TotalMontoInss") = DtaDetalleNomina.Recordset("MontoInss")
        DtaNomina.Recordset("TotalMontoIR") = DtaDetalleNomina.Recordset("MontoIR")
        DtaNomina.Recordset.Update
      End If
   End If

End Sub

Private Sub CmdIniciar3_Click()
On Error GoTo TipoErrs

  Dim sql As String, CodigoEmpleado As String, Dia As Date, Hora As String, Fecha As Date, Fecha2 As Date
  Dim Nombres  As String, Nombre1 As String, Nombre2 As String, Apellido As String, Apellido2 As String, Direccion As String
  Dim Nacionalidad As String, NumeroCedula As String, Sexo As String, NumeroInss As String, NumeroRuc As String
  Dim NHijos As Double, CodDepartamento As String, departamento As String, CodCargo As String, Cargo As String
  Dim CodGrupo As String, DescripcionGrupo As String, CodTipoNomina As String, DescripcionTipo As String
  Dim CodTurno As String, CodigoEmpleado1 As String, SueldoPeriodo As Double, TarifaHorario As Double
  Dim Id As Double, Monto As Double, CodDeduccion As String, CodEmpleado As Double, NumeroNomina As Double, Numdeduccion As Double
  
          Me.TDBGrid2.MoveFirst
          
            CodTipoNomina = Me.TDBCombo1.Columns(0).Text

         

          Do While Not Me.TDBGrid2.EOF
            NumeroInss = Me.TDBGrid2.Columns(0).Text
            CodigoEmpleado = Me.TDBGrid2.Columns(1).Text
            Nombres = Me.TDBGrid2.Columns(2).Text
            Monto = Me.TDBGrid2.Columns(3).Text
            CodDeduccion = Me.TDBGrid2.Columns(4).Text
            
            
         '------------------------BUSCO EL CODIGO DEL EMPLEADO ----------------------------------------------------------
            Me.DtaConsulta.RecordSource = "SELECT  * From Empleado WHERE  (CodEmpleado1 = '" & CodigoEmpleado & "') AND (Activo = 1)"
            Me.DtaConsulta.Refresh
            If Not Me.DtaConsulta.Recordset.EOF Then
              CodEmpleado = Me.DtaConsulta.Recordset("CodEmpleado")
              CodTipoNomina = Me.DtaConsulta.Recordset("CodTipoNomina")
            Else
              CodEmpleado = 0
            End If

        
            Me.DtaConsulta.RecordSource = "SELECT * FROM TipoNomina INNER JOIN Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina INNER JOIN Nomina ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina Where (Empleado.CodEmpleado = " & CodEmpleado & ") And (Nomina.Activa = 1) And (Empleado.Activo = 1) "
            Me.DtaConsulta.Refresh
            If Me.DtaConsulta.Recordset.EOF Then
              MsgBox "No Existe, Nomina Activa para este Empleado Codigo:" & CodigoEmpleado, vbCritical, "Sistema de Nominas"
              
            Else
              NumeroNomina = Me.DtaConsulta.Recordset("NumNomina")
              
         
          

           
                    Me.DtaConsulta.RecordSource = "SELECT  * From Deduccion Where (CodEmpleado = " & CodEmpleado & ") And (NumNomina = " & NumeroNomina & ") AND (Deduccion.CodTipoDeduccion =  '" & CodDeduccion & "')"
                    Me.DtaConsulta.Refresh
                    If Me.DtaConsulta.Recordset.EOF Then
                    
                         Me.DtaDeduccion.ConnectionString = Conexion
                         Me.DtaDeduccion.RecordSource = "SELECT NumDeduccion, CodEmpleado, CodTipoDeduccion, NumVeces, Pagado, NUmNomina From Deduccion"
                         Me.DtaDeduccion.Refresh
                         If Me.DtaDeduccion.Recordset.EOF Then
                          Numdeduccion = 0
                         Else
                           Me.DtaDeduccion.Recordset.MoveLast
                           Numdeduccion = Me.DtaDeduccion.Recordset("NumDeduccion") + 1
                        End If
        
                        DtaDeduccion.Recordset.AddNew
                        DtaDeduccion.Recordset("NumDeduccion") = Numdeduccion
                        DtaDeduccion.Recordset("CodEmpleado") = val(CodEmpleado)
                        DtaDeduccion.Recordset("codtipodeduccion") = CodDeduccion
                        DtaDeduccion.Recordset("numveces") = 1
                        DtaDeduccion.Recordset("pagado") = False
                        DtaDeduccion.Recordset("NumNomina") = NumeroNomina
                        DtaDeduccion.Recordset.Update
                        
                   Else
                        
                       Numdeduccion = Me.DtaConsulta.Recordset("NumDeduccion")
                    End If
                    
                    
                    
                    Me.DtaConsulta.RecordSource = "SELECT Id, NumDeduccion, Valor, NumVez, Pagado, NumNomina From DetalleDeduccion"
                    Me.DtaConsulta.Refresh
                    If Me.DtaConsulta.Recordset.EOF Then
                       Id = 1
                    Else
                       Me.DtaConsulta.Recordset.MoveLast
                      ' Id = Me.DtaConsulta.Recordset("Id") + 1
   
                    End If
                    
                    
                    DtaDetalleDeduccion2.Refresh
        
                    Me.DtaConsulta.RecordSource = "SELECT  * From  DetalleDeduccion Where (NumDeduccion = " & Numdeduccion & ") And (NumNomina = " & NumeroNomina & ") And Pagado = 'False'"
                    Me.DtaConsulta.Refresh
                    If Me.DtaConsulta.Recordset.EOF Then
                    DtaDetalleDeduccion2.Recordset.AddNew
                     'DtaDetalleDeduccion2.Recordset("ID") = Id
                     DtaDetalleDeduccion2.Recordset("NumDeduccion") = Numdeduccion
                     DtaDetalleDeduccion2.Recordset("valor") = val(Monto)
                     DtaDetalleDeduccion2.Recordset("NumVez") = 1
                     DtaDetalleDeduccion2.Recordset("pagado") = False
                     DtaDetalleDeduccion2.Recordset("NumNomina") = NumeroNomina
                     DtaDetalleDeduccion2.Recordset.Update
                    Else
                     DtaDetalleDeduccion2.Recordset("NumDeduccion") = Numdeduccion
                     DtaDetalleDeduccion2.Recordset("valor") = val(Monto)
                     DtaDetalleDeduccion2.Recordset.Update
                    End If
             
          
             End If
             Me.Caption = "Procesando " & Nombres
             DoEvents
             Me.TDBGrid2.MoveNext
          Loop
          
TipoErrs:
If Err.Number = 0 Then
Else
MsgBox Err.Number, vbCritical, "Zeus Nominas"
End If



Exit Sub
          
End Sub

Private Sub CmdIniciarIngresos_Click()
On Error GoTo TipoErrs

  Dim sql As String, CodigoEmpleado As String, Dia As Date, Hora As String, Fecha As Date, Fecha2 As Date
  Dim Nombres  As String, Nombre1 As String, Nombre2 As String, Apellido As String, Apellido2 As String, Direccion As String
  Dim Nacionalidad As String, NumeroCedula As String, Sexo As String, NumeroInss As String, NumeroRuc As String
  Dim NHijos As Double, CodDepartamento As String, departamento As String, CodCargo As String, Cargo As String
  Dim CodGrupo As String, DescripcionGrupo As String, CodTipoNomina As String, DescripcionTipo As String
  Dim CodTurno As String, CodigoEmpleado1 As String, SueldoPeriodo As Double, TarifaHorario As Double
  Dim Id As Double, Monto As Double, CodIncentivo As String, CodEmpleado As Double, NumeroNomina As Double, NumIncentivo As Double
  
          Me.TDBGridIngresos.MoveFirst
          
            CodTipoNomina = Me.TDBComboIngresos.Columns(0).Text

         

          Do While Not Me.TDBGridIngresos.EOF
            NumeroInss = Me.TDBGridIngresos.Columns(0).Text
            CodigoEmpleado = Format(Me.TDBGridIngresos.Columns(1).Text, "000#") 'Me.TDBGridIngresos.Columns(1).Text
            Nombres = Me.TDBGridIngresos.Columns(2).Text
            Monto = Me.TDBGridIngresos.Columns(3).Text
            CodIncentivo = Me.TDBGridIngresos.Columns(4).Text
            
            
         '------------------------BUSCO EL CODIGO DEL EMPLEADO ----------------------------------------------------------
            Me.DtaConsulta.RecordSource = "SELECT  * From Empleado WHERE  (CodEmpleado1 = '" & CodigoEmpleado & "') AND (Activo = 1)"
            Me.DtaConsulta.Refresh
            If Not Me.DtaConsulta.Recordset.EOF Then
              CodEmpleado = Me.DtaConsulta.Recordset("CodEmpleado")
              CodTipoNomina = Me.DtaConsulta.Recordset("CodTipoNomina")
            Else
              CodEmpleado = 0
            End If

          '/////////////////////////BUSCO SI EXISTE EN LA NOMINA ACTIVO /////////////////////////////
            Me.DtaConsulta.RecordSource = "SELECT * FROM TipoNomina INNER JOIN Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina INNER JOIN Nomina ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina Where (Empleado.CodEmpleado = " & CodEmpleado & ") And (Nomina.Activa = 1) And (Empleado.Activo = 1) "
            Me.DtaConsulta.Refresh
            If Me.DtaConsulta.Recordset.EOF Then
              MsgBox "No Existe, Nomina Activa para este Empleado Codigo:" & CodigoEmpleado, vbCritical, "Sistema de Nominas"
              
            Else
              NumeroNomina = Me.DtaConsulta.Recordset("NumNomina")
              
         
          

           
                    Me.DtaConsulta.RecordSource = "SELECT  * From Incentivo Where (CodEmpleado = " & CodEmpleado & ") AND (CodTipoIncentivo=  '" & CodIncentivo & "')"
                    Me.DtaConsulta.Refresh
                    If Me.DtaConsulta.Recordset.EOF Then
                    
                    
                         Me.DtaIncentivo.ConnectionString = Conexion
                         Me.DtaIncentivo.RecordSource = "SELECT NumIncentivo, CodEmpleado, CodTipoIncentivo, NumVeces, Pagado  From Incentivo"
                         Me.DtaIncentivo.Refresh
                         If Me.DtaIncentivo.Recordset.EOF Then
                          NumIncentivo = 0
                         Else
                           Me.DtaIncentivo.Recordset.MoveLast
                           NumIncentivo = Me.DtaIncentivo.Recordset("NumIncentivo") + 1
                        End If
        
                        DtaIncentivo.Recordset.AddNew
                        DtaIncentivo.Recordset("NumIncentivo") = NumIncentivo
                        DtaIncentivo.Recordset("CodEmpleado") = val(CodEmpleado)
                        DtaIncentivo.Recordset("CodTipoIncentivo") = CodIncentivo
                        DtaIncentivo.Recordset("numveces") = 1
                        DtaIncentivo.Recordset("pagado") = False
'                        DtaIncentivo.Recordset("NumNomina") = NumeroNomina
                        DtaIncentivo.Recordset.Update
                        
                   Else
                        
                      NumIncentivo = Me.DtaConsulta.Recordset("NumIncentivo")
                    End If
                    
                    
                    
                    Me.DtaConsulta.RecordSource = "SELECT NumIncentivo, Valor, NumVez, Pagado, NumNomina, Id FROM DetalleIncentivo"
                    Me.DtaConsulta.Refresh
                    If Me.DtaConsulta.Recordset.EOF Then
                       Id = 1
                    Else
                       Me.DtaConsulta.Recordset.MoveLast
                       Id = Me.DtaConsulta.Recordset("Id") + 1
   
                    End If
                    
                    
                    DtaDetalleIncentivo.Refresh
        
                    Me.DtaConsulta.RecordSource = "SELECT  * From  DetalleIncentivo Where (NumIncentivo = " & NumIncentivo & ") And (NumNomina = " & NumeroNomina & ") And Pagado = 'False'"
                    Me.DtaConsulta.Refresh
                    If Me.DtaConsulta.Recordset.EOF Then
                    DtaConsulta.Recordset.AddNew
                     DtaConsulta.Recordset("ID") = Id
                     DtaConsulta.Recordset("NumIncentivo") = NumIncentivo
                     DtaConsulta.Recordset("valor") = val(Monto)
                     DtaConsulta.Recordset("NumVez") = 1
                     DtaConsulta.Recordset("pagado") = False
                     DtaConsulta.Recordset("NumNomina") = NumeroNomina
                     DtaConsulta.Recordset.Update
                    Else
'                     DtaDetalleIncentivo.Recordset("NumIncentivo") = NumIncentivo
                     DtaConsulta.Recordset("valor") = val(Monto)
                     DtaConsulta.Recordset.Update
                    End If
             
          
             End If
             Me.Caption = "Procesando " & Nombres
             DoEvents
             Me.TDBGridIngresos.MoveNext
          Loop
          
          MsgBox "importacion correcta!!!", vbExclamation
          
TipoErrs:
If Err.Number = 0 Then
Else
MsgBox Err.Number, vbCritical, "Zeus Nominas"
End If



Exit Sub

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim retval
Dim OPENFILENAME As String, Directorio As String
Dim Rango As String, Hoja As String, ruta As String

    On Error Resume Next
    ' Set the commom dialog properties we need
    If Me.TxtRutaLogo.Text <> "" Then
       CMRutaFoto.InitDir = Me.TxtRutaLogo.Text
    End If
    CMRutaFoto.FileName = ""
    ' We will load BMP, JPG, and TIF files
    
    CMRutaFoto.Filter = "Archivo xls |*.xls"
    ' Display common dialog box
    CMRutaFoto.ShowOpen
    Me.Text1.Text = CMRutaFoto.FileName
   
    
    
  
    ruta = Me.Text1.Text 'ruta del archivo excel
    Rango = "A1:G8"    'Text2 & ":" & Text3 'Rango de datos (opcional)
    Hoja = "Hoja1" 'Nombre de la hoja
'    ruta = "C:\"
'    Set Me.TDBGrid1.DataSource = LeerTxt(ruta)
    
    Set Me.TDBGrid2.DataSource = Leer_Excel(ruta, "Hoja1")
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim TotalExcel As Double, TotalSistema As Double, TotalFicha As Double, TotalSolicitud As Double
Dim CodigoEmpleado As String
Dim NumeroSolicitud As String


With Me.AdoConsecutivo
   .ConnectionString = Conexion
   .RecordSource = "SELECT  * From Consecutivos"
   .Refresh
End With

With Me.AdoSolicitud
   .ConnectionString = Conexion
End With
        
        NumeroSolicitud = Format(ConsecutivoSolicitud, "0000#")
        gridSaldos.MoveFirst
          

        Do While Not gridSaldos.EOF
          TotalFicha = 0
          CodigoEmpleado = gridSaldos.Columns(0).Text
          
          If CodigoEmpleado = "S120080138" Then
           CodigoEmpleado = "S120080138"
          End If
          
          txtCodigo.Caption = CodigoEmpleado
          txtNombre.Caption = gridSaldos.Columns(1).Text
          
          Me.AdoConsulta.RecordSource = "Select * from Empleado where CodEmpleado1 = '" & CodigoEmpleado & "' and Activo = 'True'"
          Me.AdoConsulta.Refresh
          If Not Me.AdoConsulta.Recordset.EOF Then
            
                TotalExcel = gridSaldos.Columns(2).Text
                TotalSistema = CalculoDiasVacaciones(CodigoEmpleado, Me.dtpFin.Value)
                
                
                TotalSolicitud = 0
                 
                 

                If Me.ChkSegunNomina.Value = 1 Then
                   Dim FechaIni As Date, FechaFin As Date, FechaContrato As Date
                        '//////////////////////////////// Saco datos generales del empleado ///////////////////////
                         MDIPrimero.DtaConsulta.RecordSource = "SELECT     TOP (1) Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombre, Historico.FechaContratoVac, Empleado.CodEmpleado, DATEADD(month,    (YEAR(Historico.FechaContratoVac) - 1900) * 12 + MONTH(Historico.FechaContratoVac), - 1) AS UdMes  FROM         Empleado INNER JOIN   Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (Empleado.CodEmpleado1 = '" & CodigoEmpleado & "') and Empleado.Activo = 'True'"
                         MDIPrimero.DtaConsulta.Refresh
                         If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
                             FechaContrato = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
                         End If


'                        '/////////////////////////////////BUSCO LOS DIAS SEGUN LA NOMINA /////////////////////
                         MDIPrimero.DtaConsulta.RecordSource = "SELECT  DetalleNomVaca.NumNomVaca, DetalleNomVaca.CodEmpleado, NomVaca.FechaIni, NomVaca.FechaFin, NomVaca.Activa, Empleado.CodEmpleado1 FROM  DetalleNomVaca INNER JOIN  NomVaca ON DetalleNomVaca.NumNomVaca = NomVaca.NumNomVaca INNER JOIN  Empleado ON DetalleNomVaca.CodEmpleado = Empleado.CodEmpleado  " & _
                                                               "WHERE   (NomVaca.Activa = 1) AND (Empleado.CodEmpleado1 = '" & CodigoEmpleado & "')"
                         MDIPrimero.DtaConsulta.Refresh
                         If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
                           FechaIni = MDIPrimero.DtaConsulta.Recordset("FechaIni")
                           FechaFin = MDIPrimero.DtaConsulta.Recordset("FechaFin")
                         End If

                          TotalSistema = DateDiff("d", FechaContrato, FechaFin) * 0.0833
                          TotalSistema = CalculoDiasVacaciones(CodigoEmpleado, Me.dtpFin.Value)

                          If TotalSistema > 15 Then
                                '/////////////////////////////////////////////////////////////////////////////////////////////////
                                '///////////////////////BUSCO LOS DIAS DE VACACIONES /////////////////////////////////////////////
                                '/////////////////////////////////////////////////////////////////////////////////////////////////
                                 MDIPrimero.DtaConsulta.RecordSource = "SELECT CodigoEmpleado, SUM(DiasDisfrutar) AS Dias From SolicitudVacaciones WHERE (TipoSolicitud = 'Vacaciones') AND (Anulado = 0) AND (FechaInicio >= CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102)) AND (FechaFin <= CONVERT(DATETIME,'" & Format(FechaFin, "yyyy-MM-dd") & "', 102)) GROUP BY CodigoEmpleado, TipoSolicitud HAVING  (SolicitudVacaciones.CodigoEmpleado = '" & CodigoEmpleado & "')"
                                 MDIPrimero.DtaConsulta.Refresh
                                 If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
                                     TotalSolicitud = MDIPrimero.DtaConsulta.Recordset("Dias")
                                 Else
                                     TotalSolicitud = 0
                                 End If


                            TotalSistema = 15
                          End If
                End If

                '//////////////////////////////////////////RESTO LOS DIAS GANADOS DE VACACIONES A LAS SOLICITUDES GRABADAS ///////////////
                TotalSistema = TotalSistema - TotalSolicitud
          
          If TotalExcel < 0 Then
            TotalFicha = (TotalSistema + Math.Abs(TotalExcel))
          Else
            TotalFicha = Format(TotalSistema - TotalExcel, "####0.00")
          End If
          
          
          If TotalFicha <> 0 Then
                Me.AdoSolicitud.RecordSource = "SELECT  * From SolicitudVacaciones WHERE (NumeroSolicitud = '" & NumeroSolicitud & "')"
                Me.AdoSolicitud.Refresh
                   If Me.AdoSolicitud.Recordset.EOF Then
                      Me.AdoSolicitud.Recordset.AddNew
                               
                               
                                Me.AdoSolicitud.Recordset("FechaSolicitud") = Format(Me.dtpFin.Value, "dd/mm/yyyy")
                                Me.AdoSolicitud.Recordset("NumeroSolicitud") = NumeroSolicitud
                                Me.AdoSolicitud.Recordset("TipoSolicitud") = "Vacaciones"
                                Me.AdoSolicitud.Recordset("CodigoEmpleado") = CodigoEmpleado
                                Me.AdoSolicitud.Recordset("DiasVacaciones") = 0
                                Me.AdoSolicitud.Recordset("DiasDisfrutados") = 0
                                Me.AdoSolicitud.Recordset("FechaInicio") = Me.dtpFin.Value
                                Me.AdoSolicitud.Recordset("FechaFin") = dtpFin.Value
                                Me.AdoSolicitud.Recordset("DiasDisfrutar") = (CDbl(TotalFicha))
                                Me.AdoSolicitud.Recordset("Observaciones") = "AJUSTE AUTOMATICO"
                                Me.AdoSolicitud.Recordset.Update
                                
                                
                                
                                NumeroSolicitud = Format(CInt(NumeroSolicitud) + 1, "0000#")
                    
                    End If
                    
            End If
                
          Else
            Me.listSaldos.AddItem ("El empleado " & CodigoEmpleado & gridSaldos.Columns(1).Text & " no existe o esta inactivo")
          End If
          
          
          
          
          
          DoEvents
          gridSaldos.MoveNext
        Loop
        
        Me.AdoConsecutivo.Recordset("Solicitud") = ConsecutivoSolicitud
        Me.AdoConsecutivo.Recordset.Update
End Sub

Private Sub Command4_Click()
Dim retval
Dim OPENFILENAME As String, Directorio As String
Dim Rango As String, Hoja As String, ruta As String

    On Error Resume Next
  
    dialogSaldos.FileName = ""
    dialogSaldos.Filter = "Archivo xls |*.xls"
    ' Display common dialog box
    dialogSaldos.ShowOpen
    Dim RutaArchivo As String
    RutaArchivo = dialogSaldos.FileName
    
  
    ruta = RutaArchivo 'ruta del archivo excel
    Rango = "A1:G8"    'Text2 & ":" & Text3 'Rango de datos (opcional)
    Hoja = "Hoja1" 'Nombre de la hoja
'    ruta = "C:\"
'    Set Me.TDBGrid1.DataSource = LeerTxt(ruta)
    
    Set Me.gridMovimientoSalarial.DataSource = Leer_Excel(ruta, "Hoja1")
   ' Set Me.TDBGridNominas.DataSource = Leer_Excel(ruta, "Hoja1")
    
End Sub

Private Sub Command5_Click()
Dim CodigoEmpleado As String
Dim NombreEmpleado As String
Dim SalarioBasicoMasNivelacion As Double
Dim AumentoCordobas As Double
  gridMovimientoSalarial.MoveFirst
        Do While Not gridMovimientoSalarial.EOF
        CodigoEmpleado = gridMovimientoSalarial.Columns(0).Text
            If CodigoEmpleado = "" Then
                Exit Do
            End If

          
          NombreEmpleado = gridMovimientoSalarial.Columns(1).Text
          SalarioBasicoMasNivelacion = gridMovimientoSalarial.Columns(2).Text
          AumentoCordobas = gridMovimientoSalarial.Columns(3).Text
          
          Me.AdoConsulta.RecordSource = "Select * from Empleado where CodEmpleado1 = '" & CodigoEmpleado & "' and Activo = 'True'"
          Me.AdoConsulta.Refresh
          
          If Not Me.AdoConsulta.Recordset.EOF Then
            AdoConsulta.Recordset("OtrosIngresos") = AumentoCordobas / 2
            Me.AdoConsulta.Recordset("SueldoPeriodo") = (AumentoCordobas / 2) + (SalarioBasicoMasNivelacion / 2)
            AdoConsulta.Recordset.Update
          Else
            listMovimientoSalaria.AddItem (CodigoEmpleado + " " + NombreEmpleado)
          End If
          
          DoEvents
          gridMovimientoSalarial.MoveNext
        Loop
          
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
On Error GoTo TipoErrs

  Dim sql As String, CodigoEmpleado As Double, Dia As Date, Hora As String, Fecha As Date, Fecha2 As Date
  Dim Nombres  As String, Nombre1 As String, Nombre2 As String, Apellido As String, Apellido2 As String, Direccion As String
  Dim Nacionalidad As String, NumeroCedula As String, Sexo As String, NumeroInss As String, NumeroRuc As String
  Dim NHijos As Double, CodDepartamento As String, departamento As String, CodCargo As String, Cargo As String
  Dim CodGrupo As String, DescripcionGrupo As String, CodTipoNomina As String, DescripcionTipo As String
  Dim CodTurno As String, CodigoEmpleado1 As String, SueldoPeriodo As Double, TarifaHorario As Double
  Dim Id As Double, Monto As Double, CodDeduccion As String, CodEmpleado As Double, NumeroNomina As Double, Numdeduccion As Double
  
          Me.TDBGrid2.MoveFirst
          
            CodTipoNomina = Me.TDBCombo1.Columns(0).Text

         

          Do While Not Me.TDBGrid2.EOF
            NumeroInss = Me.TDBGrid2.Columns(0).Text
            CodigoEmpleado = Me.TDBGrid2.Columns(1).Text
            Nombres = Me.TDBGrid2.Columns(2).Text
            Monto = Me.TDBGrid2.Columns(3).Text
            CodDeduccion = Me.TDBGrid2.Columns(4).Text
            
            
         '------------------------BUSCO EL CODIGO DEL EMPLEADO ----------------------------------------------------------
            Me.DtaConsulta.RecordSource = "SELECT  * From Empleado WHERE  (CodEmpleado1 = '" & CodigoEmpleado & "') AND (Activo = 1)"
            Me.DtaConsulta.Refresh
            If Not Me.DtaConsulta.Recordset.EOF Then
              CodEmpleado = Me.DtaConsulta.Recordset("CodEmpleado")
            Else
              CodEmpleado = 0
            End If

        
            Me.DtaConsulta.RecordSource = "SELECT * FROM TipoNomina INNER JOIN Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina INNER JOIN Nomina ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina Where (Empleado.CodEmpleado = " & CodEmpleado & ") And (TipoNomina.Activa = 1) And (Empleado.Activo = 1) "
            Me.DtaConsulta.Refresh
            If Me.DtaConsulta.Recordset.EOF Then
              MsgBox "No Existe, Nomina Activa para este Empleado Codigo:" & CodigoEmpleado, vbCritical, "Sistema de Nominas"
              
            Else
              NumeroNomina = Me.DtaConsulta.Recordset("NumNomina")
              
         
          

           
                    Me.DtaConsulta.RecordSource = "SELECT  * From Deduccion Where (CodEmpleado = " & CodEmpleado & ") And (NumNomina = " & NumeroNomina & ") AND (Deduccion.CodTipoDeduccion =  '" & CodDeduccion & "')"
                    Me.DtaConsulta.Refresh
                    If Me.DtaConsulta.Recordset.EOF Then
                    
                         Me.DtaDeduccion.ConnectionString = Conexion
                         Me.DtaDeduccion.RecordSource = "SELECT NumDeduccion, CodEmpleado, CodTipoDeduccion, NumVeces, Pagado, NUmNomina From Deduccion"
                         Me.DtaDeduccion.Refresh
                         If Me.DtaDeduccion.Recordset.EOF Then
                          Numdeduccion = 0
                         Else
                           Me.DtaDeduccion.Recordset.MoveLast
                           Numdeduccion = Me.DtaDeduccion.Recordset("NumDeduccion") + 1
                        End If
        
                        DtaDeduccion.Recordset.AddNew
                        DtaDeduccion.Recordset("NumDeduccion") = Numdeduccion
                        DtaDeduccion.Recordset("CodEmpleado") = val(CodEmpleado)
                        DtaDeduccion.Recordset("codtipodeduccion") = CodDeduccion
                        DtaDeduccion.Recordset("numveces") = 1
                        DtaDeduccion.Recordset("pagado") = False
                        DtaDeduccion.Recordset("NumNomina") = NumeroNomina
                        DtaDeduccion.Recordset.Update
                        
                   Else
                        
                       Numdeduccion = Me.DtaConsulta.Recordset("NumDeduccion")
                    End If
                    
                    
                    
                    Me.DtaConsulta.RecordSource = "SELECT Id, NumDeduccion, Valor, NumVez, Pagado, NumNomina From DetalleDeduccion"
                    Me.DtaConsulta.Refresh
                    If Me.DtaConsulta.Recordset.EOF Then
                       Id = 1
                    Else
                       Me.DtaConsulta.Recordset.MoveLast
                      ' Id = Me.DtaConsulta.Recordset("Id") + 1
   
                    End If
                    
                    
                    DtaDetalleDeduccion2.Refresh
        
                    Me.DtaConsulta.RecordSource = "SELECT  * From  DetalleDeduccion Where (NumDeduccion = " & Numdeduccion & ") And (NumNomina = " & NumeroNomina & ") And Pagado = 'False'"
                    Me.DtaConsulta.Refresh
                    If Me.DtaConsulta.Recordset.EOF Then
                    DtaDetalleDeduccion2.Recordset.AddNew
                     'DtaDetalleDeduccion2.Recordset("ID") = Id
                     DtaDetalleDeduccion2.Recordset("NumDeduccion") = Numdeduccion
                     DtaDetalleDeduccion2.Recordset("valor") = val(Monto)
                     DtaDetalleDeduccion2.Recordset("NumVez") = 1
                     DtaDetalleDeduccion2.Recordset("pagado") = False
                     DtaDetalleDeduccion2.Recordset("NumNomina") = NumeroNomina
                     DtaDetalleDeduccion2.Recordset.Update
                    Else
                     DtaDetalleDeduccion2.Recordset("NumDeduccion") = Numdeduccion
                     DtaDetalleDeduccion2.Recordset("valor") = val(Monto)
                     DtaDetalleDeduccion2.Recordset.Update
                    End If
             
          
             End If
             Me.Caption = "Procesando " & Nombres
             DoEvents
             Me.TDBGrid2.MoveNext
          Loop
          
TipoErrs:
If Err.Number = 0 Then
Else
MsgBox Err.Number, vbCritical, "Zeus Nominas"
End If



Exit Sub
End Sub

Private Sub Command8_Click()
Dim retval
Dim OPENFILENAME As String, Directorio As String
Dim Rango As String, Hoja As String, ruta As String

    On Error Resume Next
    ' Set the commom dialog properties we need
    If Me.TxtRutaLogo.Text <> "" Then
       CMRutaFoto.InitDir = Me.TxtRutaLogo.Text
    End If
    CMRutaFoto.FileName = ""
    ' We will load BMP, JPG, and TIF files
    
    CMRutaFoto.Filter = "Archivo xls |*.xls"
    ' Display common dialog box
    CMRutaFoto.ShowOpen
    Me.TxtRutaIngresos.Text = CMRutaFoto.FileName
   
    
    
  
    ruta = Me.TxtRutaIngresos.Text  'ruta del archivo excel
    Rango = "A1:G8"    'Text2 & ":" & Text3 'Rango de datos (opcional)
    Hoja = "Hoja1" 'Nombre de la hoja
'    ruta = "C:\"
'    Set Me.TDBGrid1.DataSource = LeerTxt(ruta)
    
    Set Me.TDBGridIngresos.DataSource = Leer_Excel(ruta, "Hoja1")
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd

 Me.TDBGrid1.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.TDBGrid1.OddRowStyle.BackColor = &H80000005
 Me.TDBGrid1.AlternatingRowStyle = True
 
 Me.TDBGridNominas.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.TDBGridNominas.OddRowStyle.BackColor = &H80000005
 Me.TDBGridNominas.AlternatingRowStyle = True
 
 With Me.DtaDetalleNomina
   .ConnectionString = Conexion
End With

With Me.DtaDetalleDeduccion2
   .ConnectionString = Conexion
   .RecordSource = "DetalleDeduccion"
   .Refresh
End With

 
 With Me.DtaTipoNomina
   .ConnectionString = Conexion
   .RecordSource = "TipoNomina"
   .Refresh
End With

With Me.DtaDeduccion
   .ConnectionString = Conexion
End With

With Me.DtaIncentivo
   .ConnectionString = Conexion
End With
 
With Me.DtaDetalleIncentivo
   .ConnectionString = Conexion
   .RecordSource = "DetalleIncentivo"
   .Refresh
End With

With Me.DtaNomina
   .ConnectionString = Conexion
End With
 
With Me.DtaConsecutivos
   .ConnectionString = Conexion
   .RecordSource = "Consecutivos"
End With
 
With Me.DtaHistorico
   .ConnectionString = Conexion
   .RecordSource = "Historico"
   .Refresh
End With

With Me.DtaHorarioEmpleado
   .ConnectionString = Conexion
End With
 
With Me.DtaTurnos
   .ConnectionString = Conexion
   .RecordSource = "Turno"
   .Refresh
End With
 
With Me.AdoConsulta
   .ConnectionString = Conexion
End With

With Me.DtaConsulta
   .ConnectionString = Conexion
End With

 With Me.AdoRegistros
   .ConnectionString = Conexion
End With

With Me.DtaEmpleado
   .ConnectionString = Conexion
   .RecordSource = "Select * From Empleado Where (Activo = 1) ORDER BY CodEmpleado1"
   .Refresh
End With

With Me.AdoUserInfo
   .ConnectionString = Conexion
   .RecordSource = "SELECT * From Userinfo"
   .Refresh
End With
End Sub


Public Function LeerTxt(Directorio As String) As ADODB.Recordset
      On Error GoTo ErrorFunction
      Dim rs As ADODB.Recordset
      Set rs = New ADODB.Recordset
      Dim cn As ADODB.Connection
      Dim Texto1 As String, Texto2 As String
      Set cn = New ADODB.Connection
      
'
'      cn.Open "DRIVER={Microsoft Text Driver (*.txt; *.csv)};" & _
'                         "DBQ=" & Directorio & ";", "", ""
'      rs.Open "select * from [records#csv]", cn, adOpenStatic, adLockReadOnly, adCmdText

      cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
                     & "Data Source=" & Directorio & ";" _
                    & "Extended Properties='text;HDR=YES;FMT=CSVDelimited'"
                    
'      cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
'                     & "Data Source=" & Directorio & ";" _
'                    & "Extended Properties='text;HDR=YES;FMT=CSVDelimited(,)'"

       
                    
     rs.Open "select * from [records#csv]", cn, adOpenStatic, adLockReadOnly, adCmdText
     
'     Do Until rs.EOF
'        Texto1 = rs.Fields.item("ID")
'        Texto2 = rs.Fields.item("Name")
'    rs.MoveNext
'     Loop
       
      
      Set LeerTxt = rs
      
      Set rs = Nothing
      Set cn = Nothing
      
      Exit Function
ErrorFunction:
      MsgBox Err.Description, vbCritical
      Err.Clear
End Function

'devuelve un objeto Recordset con los datos de la hoja
Public Function Leer_Excel(ByVal PathXls As String, Hoja As String) As ADODB.Recordset

      On Error GoTo ErrorFunction
      Dim rs As ADODB.Recordset
      Set rs = New ADODB.Recordset
      Dim cs As String

      rs.CursorLocation = adUseClient
      rs.CursorType = adOpenKeyset
      rs.LockType = adLockBatchOptimistic

      cs = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & PathXls
      
      Hoja = "[" & Hoja & "$" & "]"
      
      rs.Open "SELECT * FROM " & Hoja, cs
      Set Leer_Excel = rs
      Set rs = Nothing
      Exit Function
      
ErrorFunction:
      MsgBox Err.Description, vbCritical
      Err.Clear
End Function


Private Sub Option1_Click()
 If Me.Option1.Value = True Then
   Me.TDBGridNominas.Visible = False
   Me.TDBGrid1.Visible = True
   Me.CmdIniciar.Visible = True
   Me.CmdIniciar2.Visible = False
   Me.SkinLabel1.Visible = False
   Me.SkinLabel2.Visible = False
   Me.DTPFechaIni.Visible = False
   Me.DtpFechaFin.Visible = False
 End If
End Sub

Private Sub Option2_Click()
 If Me.Option2.Value = True Then
   Me.TDBGridNominas.Visible = True
   Me.TDBGrid1.Visible = False
   Me.CmdIniciar.Visible = False
   Me.CmdIniciar2.Visible = True
   Me.SkinLabel1.Visible = True
   Me.SkinLabel2.Visible = True
   Me.DTPFechaIni.Visible = True
   Me.DtpFechaFin.Visible = True
 End If
End Sub


