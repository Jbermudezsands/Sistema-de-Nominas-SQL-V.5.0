VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmMovimientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de la Nomina"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   347
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   507
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "FrmMovimientos.frx":0000
      TabIndex        =   27
      Top             =   1440
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "FrmMovimientos.frx":006A
      TabIndex        =   26
      Top             =   960
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "FrmMovimientos.frx":00D6
      TabIndex        =   25
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      DownPicture     =   "FrmMovimientos.frx":0140
      Height          =   375
      Left            =   5880
      Picture         =   "FrmMovimientos.frx":1C22
      TabIndex        =   24
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox TxtNumNomina 
      Height          =   375
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox TxtNomina 
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox TxtNombres 
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox TxtApellidos 
      Height          =   375
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox TxtCodNomina 
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox TxtCodigoEmpleado 
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Top             =   480
      Width           =   1815
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
      Left            =   3240
      Picture         =   "FrmMovimientos.frx":3704
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin MSAdodcLib.Adodc DtaHorasExtra 
      Height          =   375
      Left            =   360
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
      Caption         =   "DtaHorasExtra"
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
      Left            =   3240
      Top             =   5880
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
      Caption         =   "DtaHrasExtras"
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
   Begin MSDataListLib.DataCombo DbCCodEmpleado 
      Bindings        =   "FrmMovimientos.frx":3852
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodEmpleado1"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaTipoDestajo 
      Height          =   495
      Left            =   3240
      Top             =   6960
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaTipoDestajo"
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
   Begin MSAdodcLib.Adodc DtaTipoComision 
      Height          =   375
      Left            =   360
      Top             =   6840
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
      Caption         =   "DtaTipoComision"
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
      Left            =   3240
      Top             =   6360
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
   Begin MSAdodcLib.Adodc DtaComisiones 
      Height          =   375
      Left            =   240
      Top             =   6240
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
   Begin MSAdodcLib.Adodc DtaDestajos 
      Height          =   375
      Left            =   240
      Top             =   5760
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
      Caption         =   "DtaDestajos"
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
      Left            =   2880
      Top             =   5400
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
   Begin MSAdodcLib.Adodc DtaNomina 
      Height          =   375
      Left            =   120
      Top             =   5400
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4895
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Horas Extras"
      TabPicture(0)   =   "FrmMovimientos.frx":386C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Comisiones"
      TabPicture(1)   =   "FrmMovimientos.frx":3888
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CmdBorrarComi"
      Tab(1).Control(1)=   "TxtTotalComision"
      Tab(1).Control(2)=   "ListTipoComision"
      Tab(1).Control(3)=   "CmdAplicaComi"
      Tab(1).Control(4)=   "DbgrComisiones"
      Tab(1).Control(5)=   "Label3"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Otros Destajos"
      TabPicture(2)   =   "FrmMovimientos.frx":38A4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CmdBorrarDestajo"
      Tab(2).Control(1)=   "TxtTotalDestajo"
      Tab(2).Control(2)=   "LstTipoDestajo"
      Tab(2).Control(3)=   "CmdAplicaDesta"
      Tab(2).Control(4)=   "DbgrDestajos"
      Tab(2).Control(5)=   "Label6"
      Tab(2).ControlCount=   6
      Begin VB.CommandButton CmdBorrarDestajo 
         Caption         =   "Borrar"
         DownPicture     =   "FrmMovimientos.frx":38C0
         Height          =   375
         Left            =   -74880
         Picture         =   "FrmMovimientos.frx":53A2
         TabIndex        =   16
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox TxtTotalDestajo 
         Alignment       =   1  'Right Justify
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
         Left            =   -72720
         TabIndex        =   15
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ListBox LstTipoDestajo 
         Height          =   840
         Left            =   -74640
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CommandButton CmdBorrarComi 
         Caption         =   "Borrar"
         DownPicture     =   "FrmMovimientos.frx":6E84
         Height          =   375
         Left            =   -74880
         Picture         =   "FrmMovimientos.frx":8966
         TabIndex        =   13
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox TxtTotalComision 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2880
         Width           =   1815
      End
      Begin VB.ListBox ListTipoComision 
         Height          =   840
         Left            =   -72840
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CommandButton CmdAplicaDesta 
         DownPicture     =   "FrmMovimientos.frx":A448
         Height          =   375
         Left            =   -70560
         Picture         =   "FrmMovimientos.frx":BF2A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton CmdAplicaComi 
         DownPicture     =   "FrmMovimientos.frx":D82C
         Height          =   375
         Left            =   -70800
         Picture         =   "FrmMovimientos.frx":F30E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Caption         =   "Horas Extras"
         Height          =   1815
         Left            =   1440
         TabIndex        =   6
         Top             =   480
         Width           =   3135
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmMovimientos.frx":10C10
            TabIndex        =   28
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox TxtHrasExtras 
            Height          =   375
            Left            =   1200
            TabIndex        =   8
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton CmdAgregar 
            Caption         =   "Agregar"
            DownPicture     =   "FrmMovimientos.frx":10C84
            Height          =   375
            Left            =   960
            Picture         =   "FrmMovimientos.frx":12766
            TabIndex        =   7
            Top             =   1200
            Width           =   1455
         End
      End
      Begin TrueOleDBGrid70.TDBGrid DbgrDestajos 
         Bindings        =   "FrmMovimientos.frx":14248
         Height          =   1815
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3201
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=110,.bold=0,.fontsize=825,.italic=0"
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
      Begin TrueOleDBGrid70.TDBGrid DbgrComisiones 
         Bindings        =   "FrmMovimientos.frx":14262
         Height          =   1815
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3201
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
         AllowAddNew     =   -1  'True
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
      Begin VB.Label Label6 
         Caption         =   "Total Destajo"
         Height          =   255
         Left            =   -73920
         TabIndex        =   18
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Total Comisión"
         Height          =   255
         Left            =   -74520
         TabIndex        =   17
         Top             =   3000
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const ColIndex_ListTipoComision = 0
Const ColIndex_ListTipoDestajo = 0

Public Sub CmdAgregar_Click()
Dim CodEmpleado As Double
Dim consecutivo As Double
'On Error GoTo TipoErrs

If Not IsNumeric(TxtHrasExtras.Text) Then
   MsgBox "La Cantidad de Horas Extras es errónea"
   TxtHrasExtras.SetFocus
   Exit Sub
End If

CodEmpleado = Me.TxtCodigoEmpleado.Text

'busco en Hrs Extras si ya le fue gravada una hora extra
SQLHorasExtras = "SELECT HorasExtras.CodEmpleado, HorasExtras.NumNomina, HorasExtras.CantHoras, HorasExtras.Pagada From HorasExtras WHERE HorasExtras.CodEmpleado=" & CodEmpleado & " AND HorasExtras.NumNomina= " & NumNomina & ""
DtaHrasExtras.RecordSource = SQLHorasExtras
DtaHrasExtras.Refresh
Do While Not DtaHrasExtras.Recordset.EOF
'If DtaHrasExtras.Recordset.Fields("CodEmpleado") = DbCCodEmpleado.Text And DtaHrasExtras.Recordset.NumNomina = DtaNomina.Recordset("NumNomina") Then
   MsgBox "Ya le fueron agregadas horas extras a este empleado, las horas anteriores serán reemplazadas"
   'DtaHrasExtras.Recordset.Edit
   DtaHrasExtras.Recordset.Fields("canthoras") = CDbl(TxtHrasExtras.Text)
   DtaHrasExtras.Recordset.Fields("Pagada") = 0
   DtaHrasExtras.Recordset.Update
   DtaHrasExtras.Refresh
   'TxtHrasExtras.Text = "0"
   DbCCodEmpleado.Text = ""
   TxtNombres.Text = ""
   TxtApellidos.Text = ""
   TxtCodNomina.Text = ""
   txtNomina.Text = ""
   TxtNumNomina.Text = ""
   Exit Sub
'End If

DtaHrasExtras.Recordset.MoveNext
Loop

'si no las encontro grabo las horas extras
  Me.DtaHrasExtras.RecordSource = "HorasExtras"
  Me.DtaHrasExtras.Refresh
  If Not Me.DtaHrasExtras.Recordset.EOF Then
   Me.DtaHrasExtras.Recordset.MoveLast
   consecutivo = Me.DtaHrasExtras.Recordset("ID") + 1
  Else
   consecutivo = 1
  End If

   Me.DtaHrasExtras.Recordset.AddNew
   DtaHrasExtras.Recordset.Fields("Id") = consecutivo
   DtaHrasExtras.Recordset.Fields("CodEmpleado") = CodEmpleado
   DtaHrasExtras.Recordset.Fields("NumNomina") = DtaNomina.Recordset("NumNomina")
   DtaHrasExtras.Recordset.Fields("canthoras") = CDbl(TxtHrasExtras.Text)
   DtaHrasExtras.Recordset.Fields("Pagada") = 0
   DtaHrasExtras.Recordset.Update
   
   TxtHrasExtras.Text = "0"
   DbCCodEmpleado.Text = ""
   TxtNombres.Text = ""
   TxtApellidos.Text = ""
   TxtCodNomina.Text = ""
   
Exit Sub

TipoErrs:

ControlErrores
Unload Me


End Sub

Private Sub CmdAgregarComi_Click()

End Sub

Private Sub CmdAgregarComision_Click()
On Error GoTo TipoErrs

If Not IsNumeric(TxtComision.Text) Then
   MsgBox "La Cantidad Digitada es errónea"
   TxtComision.SetFocus
   Exit Sub
End If

'pregunto si la nómina de este empleado está activada
DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("CodTipoNomina") = TxtCodNomina.Text And DtaTipoNomina.Recordset("Activa") = False Then
   MsgBox "La Nómina de ese empleado no ha sido Activada"
   Exit Sub
End If
DtaTipoNomina.Recordset.MoveNext
Loop

'averiguo si a esta empleado se le puede pagar Comisiones
DtaTipoNomina.Refresh

Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("CodTipoNomina") = TxtCodNomina.Text Then
MsgBox DtaTipoNomina.Recordset("TipoPago")
If DtaTipoNomina.Recordset("TipoPago") <> "Salario Destajo y Comision" And DtaTipoNomina.Recordset("TipoPago") <> "Salario Fijo y Comision" Then
   MsgBox "A este Empleado no se le paga Comisión"
   Exit Sub
End If
End If
DtaTipoNomina.Recordset.MoveNext
Loop



'busco en Hrs Extras si ya le fue gravada una hora extra
DtaComisiones.Refresh
Do While Not DtaComisiones.Recordset.EOF

If DtaComisiones.Recordset("CodEmpleado") = DbCCodEmpleado.Text And DtaComisiones.Recordset("NumNomina") = DtaNomina.Recordset("NumNomina") Then
   MsgBox "Ya le fue gravado la Comision a este empleado, la cantidad anterior será reemplazada"
'   DtaComisiones.Recordset.Edit
   DtaComisiones.Recordset("cantidad") = val(TxtComision.Text)
   DtaComisiones.Recordset.Update
   DtaComisiones.Refresh
   
   TxtHrasExtras.Text = "0"
   DbCCodEmpleado.Text = ""
   TxtNombres.Text = ""
   TxtApellidos.Text = ""
   TxtCodNomina.Text = ""
   Exit Sub
End If

DtaComisiones.Recordset.MoveNext
Loop

'si no las encontro grabo las horas extras

   DtaComisiones.Recordset.AddNew
   DtaComisiones.Recordset("CodEmpleado") = DbCCodEmpleado.Text
   DtaComisiones.Recordset("NumNomina") = DtaNomina.Recordset("NumNomina")
   DtaComisiones.Recordset("cantidad") = val(TxtComision.Text)
   DtaComisiones.Recordset.Update
   
   TxtHrasExtras.Text = "0"
   DbCCodEmpleado.Text = ""
   TxtNombres.Text = ""
   TxtApellidos.Text = ""
   TxtCodNomina.Text = ""
   
Exit Sub

TipoErrs:

ControlErrores
Unload Me



End Sub

Private Sub CmdAgregarDestajo_Click()
On Error GoTo TipoErrs

If Not IsNumeric(TxtDestajo.Text) Then
   MsgBox "La Cantidad Digitada es errónea"
   TxtDestajo.SetFocus
   Exit Sub
End If

'pregunto si la nómina de este empleado está activada
DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("CodTipoNomina") = TxtCodNomina.Text And DtaTipoNomina.Recordset("Activa") = False Then
   MsgBox "La Nómina de ese empleado no ha sido Activada"
   Exit Sub
End If
DtaTipoNomina.Recordset.MoveNext
Loop


'averiguo si a esta empleado se le puede pagar destajo
DtaTipoNomina.Refresh

Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("CodTipoNomina") = TxtCodNomina.Text Then
If DtaTipoNomina.Recordset("TipoPago") <> "Salario Destajo y Comision" And DtaTipoNomina.Recordset("TipoPago") <> "Salario Destajo" Then
   MsgBox "A este Empleado no se le paga al Destajo"
   Exit Sub
End If
End If
DtaTipoNomina.Recordset.MoveNext
Loop

'busco en Hrs Extras si ya le fue gravada una hora extra
DtaDestajos.Refresh
Do While Not DtaDestajos.Recordset.EOF
If DtaDestajos.Recordset("CodEmpleado") = DbCCodEmpleado.Text And DtaDestajos.Recordset("NumNomina") = DtaNomina.Recordset("NumNomina") Then
   MsgBox "Ya le fue gravado el destajo a este empleado, la cantidad anterior será reemplazada"
   'DtaDestajos.Recordset.Edit
   DtaDestajos.Recordset("cantidad") = val(TxtDestajo.Text)
   DtaDestajos.Recordset.Update
   DtaDestajos.Refresh
   
   TxtHrasExtras.Text = "0"
   DbCCodEmpleado.Text = ""
   TxtNombres.Text = ""
   TxtApellidos.Text = ""
   TxtCodNomina.Text = ""
   Exit Sub
End If

DtaDestajos.Recordset.MoveNext
Loop

'si no las encontro grabo las horas extras

   DtaDestajos.Recordset.AddNew
   DtaDestajos.Recordset("CodEmpleado") = DbCCodEmpleado.Text
   DtaDestajos.Recordset("NumNomina") = DtaNomina.Recordset("NumNomina")
   DtaDestajos.Recordset("cantidad") = val(TxtDestajo.Text)
   DtaDestajos.Recordset.Update
   
   DbCCodEmpleado.Text = ""
   TxtNombres.Text = ""
   TxtApellidos.Text = ""
   TxtCodNomina.Text = ""
   TxtDestajo.Text = "0"
   
   
Exit Sub

TipoErrs:

ControlErrores
Unload Me


End Sub

Private Sub CmdAplicaComi_Click()
DtaEmpleado.Refresh
'Busco el codigo del empleado para que automaticamente ubique el nombre
 'aunque no existe en la data consulta
    Do While Not DtaEmpleado.Recordset.EOF
     If DtaEmpleado.Recordset("CodEmpleado") = DbCCodEmpleado.Text Then
       'DtaEmpleado.Recordset.Edit
       DtaEmpleado.Recordset("PorcentajeComision") = CDbl(TxtTotalComision.Text)
       DtaEmpleado.Recordset.Update
       Exit Do
     End If
       DtaEmpleado.Recordset.MoveNext
   Loop

End Sub

Private Sub CmdAplicaDesta_Click()
DtaEmpleado.Refresh
'Busco el codigo del empleado para que automaticamente ubique el nombre
 'aunque no existe en la data consulta
    Do While Not DtaEmpleado.Recordset.EOF
     If DtaEmpleado.Recordset("CodEmpleado") = DbCCodEmpleado.Text Then
       'DtaEmpleado.Recordset.Edit
       DtaEmpleado.Recordset("TarifaHoraria") = CDbl(TxtTotalDestajo.Text)
       DtaEmpleado.Recordset.Update
       Exit Do
     End If
       DtaEmpleado.Recordset.MoveNext
   Loop
End Sub

Private Sub CmdBorrarComi_Click()
On Error GoTo TipoErr
Dim Total As Double
DtaComisiones.Recordset.Delete

'coloco el total de las comisiones
        Total = 0
        DtaComisiones.Refresh
        Do While Not DtaComisiones.Recordset.EOF
        Total = Total + DtaComisiones.Recordset("Total")
        DtaComisiones.Recordset.MoveNext
        Loop
        
        TxtTotalComision.Text = Format(Total, "###,##0.00")
        DbgrComisiones.Columns(4).Visible = False
        DbgrComisiones.Columns(5).Visible = False
        DbgrComisiones.Columns(6).Visible = False
        DbgrComisiones.Columns(0).Width = 2000
        DbgrComisiones.Columns(1).Width = 1200
        DbgrComisiones.Columns(2).Width = 1200
        DbgrComisiones.Columns(3).Width = 1200
        DbgrComisiones.Columns(1).NumberFormat = "#0.00%"
        DbgrComisiones.Columns(3).NumberFormat = "##,##0.00"
        DbgrComisiones.Columns(1).Locked = True
        DbgrComisiones.Columns(3).Locked = True
        DbgrComisiones.Columns(0).Button = True

Exit Sub

TipoErr:
ControlErrores

End Sub

Private Sub CmdBorrarDestajo_Click()
On Error GoTo TipoErr
Dim Total As Double
DtaDestajos.Recordset.Delete

'coloco el total de los destajos
        Total = 0
        DtaDestajos.Refresh
        Do While Not DtaDestajos.Recordset.EOF
        Total = Total + DtaDestajos.Recordset("Total")
        DtaDestajos.Recordset.MoveNext
        Loop
        TxtTotalDestajo.Text = Format(Total, "###,##0.00")
        DbgrDestajos.Columns(4).Visible = False
        DbgrDestajos.Columns(5).Visible = False
        DbgrDestajos.Columns(6).Visible = False
        DbgrDestajos.Columns(0).Width = 2000
        DbgrDestajos.Columns(1).Width = 1200
        DbgrDestajos.Columns(2).Width = 1200
        DbgrDestajos.Columns(3).Width = 1200
        DbgrDestajos.Columns(1).NumberFormat = "##,##0.00"
        DbgrDestajos.Columns(3).NumberFormat = "##,##0.00"
        DbgrDestajos.Columns(1).Locked = True
        DbgrDestajos.Columns(3).Locked = True
        DbgrDestajos.Columns(0).Button = True
Exit Sub

TipoErr:
ControlErrores
End Sub

Private Sub cmdBuscar_Click()
FrmBuscaEmpleado.Show 1
End Sub

Private Sub CmdBuscarEmpleado_Click()
FrmBuscaEmpleado.Show 1
End Sub

Private Sub CmdHistoDeducciones_Click()


End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DBCCodEmpleado_Change()
'On Error GoTo TipoErr
Dim SQLNomina As String
Dim TipoNomina As String
Dim SQLHorasExtras As String
Dim SQLDestajos As String
Dim SQlComisiones As String
Dim Total As Double
Dim Encontrado As Boolean

Encontrado = False
DtaEmpleado.Refresh
'Busco el codigo del empleado para que automaticamente ubique el nombre
 'aunque no existe en la data consulta
    Do While Not DtaEmpleado.Recordset.EOF
     If DtaEmpleado.Recordset("CodEmpleado1") = DbCCodEmpleado.Text Then
        CodEmpleado = DtaEmpleado.Recordset("CodEmpleado")
        Me.TxtCodigoEmpleado.Text = DtaEmpleado.Recordset("CodEmpleado")
        If Not IsNull(DtaEmpleado.Recordset("Nombre2")) Then
          TxtNombres.Text = DtaEmpleado.Recordset("Nombre1") + " " + DtaEmpleado.Recordset("Nombre2")
        End If
        TxtApellidos.Text = DtaEmpleado.Recordset("Apellido1") + " " + DtaEmpleado.Recordset("Apellido2")
        TxtCodNomina.Text = DtaEmpleado.Recordset("CodTipoNomina")
        Encontrado = True
        
        Exit Do
     End If
       DtaEmpleado.Recordset.MoveNext
   Loop
   
If Encontrado = False Then
        TxtNombres.Text = ""
        TxtApellidos.Text = ""
        TxtCodNomina.Text = ""
        txtNomina.Text = ""
        SSTab1.Enabled = False
        Exit Sub
End If

TipoNomina = TxtCodNomina.Text
SQLNomina = "SELECT Nomina.*, TipoNomina.TipoPago FROM TipoNomina INNER JOIN Nomina ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina Where (((Nomina.Activa) = 1)) And Nomina.CodTipoNomina = '" & TipoNomina & "' "
DtaNomina.RecordSource = SQLNomina
DtaNomina.Refresh

If DtaNomina.Recordset.EOF Then
   MsgBox "Este Empleado no tiene su nómina activa"
   SSTab1.Enabled = False
Else
   SSTab1.Enabled = True
   TxtNumNomina.Text = DtaNomina.Recordset("NumNomina")
   NumNomina = DtaNomina.Recordset("NumNomina")
'coloco los sqls para este empleado

SQLDestajos = "SELECT TipoDestajo.Destajo, TipoDestajo.Monto, DestalleDestajos.Cantidad, [DestalleDestajos].[Cantidad]*[TipoDestajo].[Monto] AS Total, DestalleDestajos.NUmNomina, DestalleDestajos.CodEmpleado,DestalleDestajos.CodTipoDestajo FROM TipoDestajo INNER JOIN DestalleDestajos ON TipoDestajo.COdTipoDestajo = DestalleDestajos.CodTipoDestajo WHERE DestalleDestajos.CodEmpleado='" & CodEmpleado & "' AND DestalleDestajos.NUmNomina=" & NumNomina & ""
DtaDestajos.RecordSource = SQLDestajos
DtaDestajos.Refresh

SQLHorasExtras = "SELECT HorasExtras.CodEmpleado, HorasExtras.NumNomina, HorasExtras.CantHoras, HorasExtras.Pagada From HorasExtras WHERE HorasExtras.CodEmpleado='" & CodEmpleado & "' AND HorasExtras.NumNomina= " & NumNomina & ""
DtaHrasExtras.RecordSource = SQLHorasExtras
DtaHrasExtras.Refresh

If Not DtaHrasExtras.Recordset.EOF Then
   TxtHrasExtras.Text = Format(DtaHrasExtras.Recordset.Fields("canthoras"), "###,##0.00")
Else
   TxtHrasExtras.Text = "0.00"
End If

SQlComisiones = "SELECT TipoComision.Comision, TipoComision.Porcentaje, DetalleComisiones.Cantidad, [DetalleComisiones].[Cantidad]*[TipoComision].[Porcentaje] AS Total, DetalleComisiones.NUmNomina, DetalleComisiones.CodEmpleado, DetalleComisiones.CodTipoComision  FROM TipoComision INNER JOIN DetalleComisiones ON TipoComision.CodTipoComision = DetalleComisiones.CodTipoComision WHERE DetalleComisiones.CodEmpleado= '" & CodEmpleado & "' AND DetalleComisiones.NUmNomina= " & NumNomina & ""
DtaComisiones.RecordSource = SQlComisiones
DtaComisiones.Refresh

'coloco las propiedades de los dbgrids
'DBGridFactura.Columns(4).NumberFormat = "##,##0.00"
'DBGridFactura.Columns(2).Locked = True
'DBGridFactura.Columns(1).Visible = False
'DBGridFactura.Columns(4).Text = Valor
'Dbgrdetalle_ButtonClick (1)
'Dbgrdetalle.Columns(4).Width = TextWidth(String$(10, "O"))
'Dbgrdetalle.Columns(1).Button = True


   'evaluo el tipo de nómina para ver si activo las pestañas correspondientes
   If DtaNomina.Recordset("TipoPago") = "Salario Fijo" Then
   SSTab1.TabEnabled(1) = False
   SSTab1.TabEnabled(2) = False
   ElseIf DtaNomina.Recordset("TipoPago") = "Salario Destajo" Then
   SSTab1.TabEnabled(1) = False
   SSTab1.TabEnabled(2) = True
   ElseIf DtaNomina.Recordset("TipoPago") = "Salario Fijo y Comision" Then
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = False
   ElseIf DtaNomina.Recordset("TipoPago") = "Salario Destajo y Comision" Then
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = True
   ElseIf DtaNomina.Recordset("TipoPago") = "Salario Fijo,Destajo y Comision" Then
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = True
   End If
   
   If DtaEmpleado.Recordset("SalarioFijo") = "S" Then
         SSTab1.TabEnabled(1) = False
         SSTab1.TabEnabled(2) = False
        Else
         'SSTab1.TabEnabled(1) = True
         'SSTab1.TabEnabled(2) = False
        End If
'coloco el total de las comisiones
        Total = 0
        DtaComisiones.Refresh
        Do While Not DtaComisiones.Recordset.EOF
        Total = Total + DtaComisiones.Recordset("Total")
        DtaComisiones.Recordset.MoveNext
        Loop
        
        TxtTotalComision.Text = Format(Total, "###,##0.00")
        
        DbgrComisiones.Columns(4).Visible = False
        DbgrComisiones.Columns(5).Visible = False
        DbgrComisiones.Columns(6).Visible = False
        DbgrComisiones.Columns(0).Width = 2000
        DbgrComisiones.Columns(1).Width = 1200
        DbgrComisiones.Columns(2).Width = 1200
        DbgrComisiones.Columns(3).Width = 1200
        DbgrComisiones.Columns(1).NumberFormat = "#0.00%"
        DbgrComisiones.Columns(3).NumberFormat = "##,##0.00"
        DbgrComisiones.Columns(1).Locked = True
        DbgrComisiones.Columns(3).Locked = True
        DbgrComisiones.Columns(0).Button = True
      
'coloco el total de los destajos
        Total = 0
        DtaDestajos.Refresh
        Do While Not DtaDestajos.Recordset.EOF
        Total = Total + DtaDestajos.Recordset("Total")
        DtaDestajos.Recordset.MoveNext
        Loop
        
        TxtTotalDestajo.Text = Format(Total, "###,##0.00")
        
        DbgrDestajos.Columns(4).Visible = False
        DbgrDestajos.Columns(5).Visible = False
        DbgrDestajos.Columns(6).Visible = False
        DbgrDestajos.Columns(1).NumberFormat = "##,##0.00"
        DbgrDestajos.Columns(3).NumberFormat = "##,##0.00"
        DbgrDestajos.Columns(0).Width = 2000
        DbgrDestajos.Columns(1).Width = 1200
        DbgrDestajos.Columns(2).Width = 1200
        DbgrDestajos.Columns(3).Width = 1200
        DbgrDestajos.Columns(1).Locked = True
        DbgrDestajos.Columns(3).Locked = True
        DbgrDestajos.Columns(0).Button = True
      


End If
Exit Sub
TipoErr:
ControlErrores
End Sub

Private Sub DBCTipoComision_Click(Area As Integer)

End Sub

Private Sub DBCTipoComision_KeyPress(KeyAscii As Integer)
    '//Habilita el uso de teclado para asignar a la celda o Escape...
    Select Case KeyAscii
        Case vbKeyReturn
             '//Asigma el texto seleccionado a la celda
             DbgrComisiones.Columns(ColIndex_List1).Text = List1.Text
             List1.Visible = False
        Case vbKeyEscape
             List1.Visible = False
    End Select

End Sub

Private Sub DbCCodEmpleado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = "13" Then
    Me.TxtHrasExtras.SetFocus
     Me.TxtHrasExtras.SelStart = 0
     Me.TxtHrasExtras.SelLength = Len(TxtHrasExtras.Text)
End If

End Sub

Private Sub DbgrComisiones_AfterUpdate()
Dim Total As Double
Total = 0
DtaComisiones.Recordset.MoveFirst
Do While Not DtaComisiones.Recordset.EOF
Total = Total + DtaComisiones.Recordset("Total")
DtaComisiones.Recordset.MoveNext
Loop


Criterio = "CodEmpleado='" & Me.DbCCodEmpleado.Text & "'"
Me.DtaEmpleado.Recordset.Find Criterio
If Not Me.DtaEmpleado.Recordset.EOF Then
 'Me.'DtaEmpleado.Recordset.Edit
  Me.DtaEmpleado.Recordset("PorcentajeComision") = Format(Total, "###,##0.00")
 
 Me.DtaEmpleado.Recordset.Update
End If


TxtTotalComision.Text = Format(Total, "###,##0.00")

DbgrComisiones.Columns(4).Visible = False
DbgrComisiones.Columns(5).Visible = False
DbgrComisiones.Columns(6).Visible = False
DbgrComisiones.Columns(1).NumberFormat = "#0.00%"
DbgrComisiones.Columns(3).NumberFormat = "##,##0.00"
        DbgrComisiones.Columns(0).Width = 2000
        DbgrComisiones.Columns(1).Width = 1200
        DbgrComisiones.Columns(2).Width = 1200
        DbgrComisiones.Columns(3).Width = 1200
DbgrComisiones.Columns(1).Locked = True
DbgrComisiones.Columns(3).Locked = True
DbgrComisiones.Columns(0).Button = True


End Sub

Private Sub DbgrComisiones_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
   
   '// Use este evento para que cuando el usuario teclee un caracter sobre la celda
   '// se despliegue la lista. Es decir, se obliga al usuario a usar un ítem de la lista.
   '// En caso de dar al usuario libertad de escribir, elimine las siguientes líneas (If-End If),
   '// o precedales con un comentario
   
    If ColIndex = ListTipoComision_List Then
       '// Se obliga a seleccionar de la lista:
       Cancel = True
       DbgrComisiones_ButtonClick (ColIndex)
    End If

End Sub

Private Sub DbgrComisiones_ButtonClick(ByVal ColIndex As Integer)
 Dim c As Column
    
 Set c = DbgrComisiones.Columns(ColIndex)
   With ListTipoComision
           '// Despliegue de la lista al lado de la celda.
           '// Elimine los comentarios de las dos siguientes líneas
           '// y coloque comentarios a las tres posteriores. A su gusto
           '.Left = DBGrid1.Left + C.Left + C.Width
           '.Top = DBGrid1.Top + DBGrid1.RowTop(DBGrid1.Row)
            
           '// Lista debajo de la celda, al estilo ComboBox (3 líneas)
        .Left = DbgrComisiones.Left + c.Left
        .Top = DbgrComisiones.Top + DbgrComisiones.RowTop(DbgrComisiones.Row) + DbgrComisiones.RowHeight
        .Width = c.Width + 15

        .ListIndex = 0
        .Visible = True
        .ZOrder 0
        .SetFocus

       End With

End Sub

Private Sub DbgrComisiones_Scroll(Cancel As Integer)
    '//Oculta la lista si hace Scroll
    ListTipoComision.Visible = False
End Sub

Private Sub DbgrDestajos_AfterUpdate()
Dim Total As Double
Total = 0
DtaDestajos.Recordset.MoveFirst
Do While Not DtaDestajos.Recordset.EOF
Total = Total + DtaDestajos.Recordset("Total")
DtaDestajos.Recordset.MoveNext
Loop

TxtTotalDestajo.Text = Format(Total, "###,##0.00")

DbgrDestajos.Columns(4).Visible = False
DbgrDestajos.Columns(5).Visible = False
DbgrDestajos.Columns(6).Visible = False
        DbgrDestajos.Columns(0).Width = 2000
        DbgrDestajos.Columns(1).Width = 1200
        DbgrDestajos.Columns(2).Width = 1200
        DbgrDestajos.Columns(3).Width = 1200
DbgrDestajos.Columns(1).NumberFormat = "##,##0.00"
DbgrDestajos.Columns(3).NumberFormat = "##,##0.00"
DbgrDestajos.Columns(1).Locked = True
DbgrDestajos.Columns(3).Locked = True
DbgrDestajos.Columns(0).Button = True
End Sub

Private Sub DbgrDestajos_ButtonClick(ByVal ColIndex As Integer)
 Dim c As Column
    
 Set c = DbgrDestajos.Columns(ColIndex)
   With LstTipoDestajo
           '// Despliegue de la lista al lado de la celda.
           '// Elimine los comentarios de las dos siguientes líneas
           '// y coloque comentarios a las tres posteriores. A su gusto
           '.Left = DBGrid1.Left + C.Left + C.Width
           '.Top = DBGrid1.Top + DBGrid1.RowTop(DBGrid1.Row)
            
           '// Lista debajo de la celda, al estilo ComboBox (3 líneas)
        .Left = DbgrDestajos.Left + c.Left
        .Top = DbgrDestajos.Top + DbgrDestajos.RowTop(DbgrDestajos.Row) + DbgrDestajos.RowHeight
        .Width = c.Width + 15

        .ListIndex = 0
        .Visible = True
        .ZOrder 0
        .SetFocus

       End With

End Sub

Private Sub DbgrDestajos_Scroll(Cancel As Integer)
    '//Oculta la lista si hace Scroll
    LstTipoDestajo.Visible = False
End Sub

Private Sub Form_Activate()
DbCCodEmpleado.Text = CodEmpleado

End Sub

Private Sub Form_Load()
Dim sql As String
SSTab1.Enabled = False
MDIPrimero.Skin1.ApplySkin hWnd


With Me.DtaComisiones
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaHorasExtra
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaDestajos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "TipoDestajo"
   .Refresh
End With

With Me.DtaEmpleado
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT CodEmpleado, CodEmpleado1, Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula, Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria, PorcentajeComision, OtrosIngresos, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones,Liquidado , Ausente, SalarioFijo, SumarSubsidio, PorcientoIncentivo From Empleado Where (Activo = 1) ORDER BY CodEmpleado1"
   .Refresh
                      
End With

With Me.DtaHrasExtras
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "HorasExtras"
   .Refresh
End With

With Me.DtaNomina
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaTipoComision
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "TipoComision"
   .Refresh
End With

With Me.DtaTipoDestajo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "TipoDestajo"
   .Refresh
End With

With Me.DtaTipoNomina
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "TipoNomina"
   .Refresh
End With

DtaTipoComision.Refresh
Do While Not DtaTipoComision.Recordset.EOF
   ListTipoComision.AddItem DtaTipoComision.Recordset("Comision")
DtaTipoComision.Recordset.MoveNext
Loop

DtaTipoDestajo.Refresh
Do While Not DtaTipoDestajo.Recordset.EOF
   LstTipoDestajo.AddItem DtaTipoDestajo.Recordset("destajo")
DtaTipoDestajo.Recordset.MoveNext
Loop


End Sub


Private Sub ListTipoComision_DblClick()
    ListTipoComision_KeyPress vbKeyReturn
End Sub

Private Sub ListTipoComision_KeyPress(KeyAscii As Integer)
Dim CodComision As String

DtaTipoComision.Refresh
Do While Not DtaTipoComision.Recordset.EOF
   If ListTipoComision.Text = DtaTipoComision.Recordset("Comision") Then
   CodComision = DtaTipoComision.Recordset("Codtipocomision")
   End If
DtaTipoComision.Recordset.MoveNext
Loop


'//Habilita el uso de teclado para asignar a la celda o Escape...
    Select Case KeyAscii
        Case vbKeyReturn
             '//Asigna el texto seleccionado a la celda
             'DbgrComisiones.Columns(ColIndex_ListTipoComision).Text = ListTipoComision.Text
             
             DbgrComisiones.Columns(4).Text = NumNomina
             DbgrComisiones.Columns(5).Text = CodEmpleado
             DbgrComisiones.Columns(6).Text = CodComision
             DtaComisiones.Refresh
             DbgrComisiones.Columns(4).Visible = False
             DbgrComisiones.Columns(5).Visible = False
             DbgrComisiones.Columns(6).Visible = False
             DbgrComisiones.Columns(0).Width = 2000
             DbgrComisiones.Columns(1).Width = 1200
             DbgrComisiones.Columns(2).Width = 1200
             DbgrComisiones.Columns(3).Width = 1200
             DbgrComisiones.Columns(1).NumberFormat = "#0.00%"
             DbgrComisiones.Columns(3).NumberFormat = "##,##0.00"
             DbgrComisiones.Columns(1).Locked = True
             DbgrComisiones.Columns(3).Locked = True
             DbgrComisiones.Columns(0).Button = True
             ListTipoComision.Visible = False
        Case vbKeyEscape
             ListTipoComision.Visible = False
    End Select

End Sub

Private Sub ListTipoComision_LostFocus()
    '//Oculta la lista si pierde el enfoque
ListTipoComision.Visible = False

End Sub

Private Sub LstTipoDestajo_DblClick()
 LstTipoDestajo_KeyPress vbKeyReturn
End Sub

Private Sub LstTipoDestajo_KeyPress(KeyAscii As Integer)
Dim CodDestajo As String

DtaTipoDestajo.Refresh
Do While Not DtaTipoDestajo.Recordset.EOF
   If LstTipoDestajo.Text = DtaTipoDestajo.Recordset("destajo") Then
   CodDestajo = DtaTipoDestajo.Recordset("Codtipodestajo")
   End If
DtaTipoDestajo.Recordset.MoveNext
Loop


'//Habilita el uso de teclado para asignar a la celda o Escape...
    Select Case KeyAscii
        Case vbKeyReturn
             '//Asigna el texto seleccionado a la celda
             'DbgrComisiones.Columns(ColIndex_ListTipoComision).Text = ListTipoComision.Text
             
             DbgrDestajos.Columns(4).Text = NumNomina
             DbgrDestajos.Columns(5).Text = CodEmpleado
             DbgrDestajos.Columns(6).Text = CodDestajo
             DtaDestajos.Refresh
             DbgrDestajos.Columns(4).Visible = False
             DbgrDestajos.Columns(5).Visible = False
             DbgrDestajos.Columns(6).Visible = False
        DbgrDestajos.Columns(0).Width = 2000
        DbgrDestajos.Columns(1).Width = 1200
        DbgrDestajos.Columns(2).Width = 1200
        DbgrDestajos.Columns(3).Width = 1200
             DbgrDestajos.Columns(1).NumberFormat = "##,##0.00"
             DbgrDestajos.Columns(3).NumberFormat = "##,##0.00"
             DbgrDestajos.Columns(1).Locked = True
             DbgrDestajos.Columns(3).Locked = True
             DbgrDestajos.Columns(0).Button = True
             LstTipoDestajo.Visible = False
        Case vbKeyEscape
             istTipodestajo.Visible = False
    End Select
End Sub

Private Sub LstTipoDestajo_LostFocus()
    '//Oculta la lista si pierde el enfoque
LstTipoDestajo.Visible = False
End Sub

Private Sub TxtCodNomina_Change()
Me.DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
  If DtaTipoNomina.Recordset("CodTipoNomina") = TxtCodNomina Then
     txtNomina.Text = DtaTipoNomina.Recordset("nomina")
  End If
DtaTipoNomina.Recordset.MoveNext
Loop

End Sub

Private Sub TxtHrasExtras_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = "13" Then
      CmdAgregar_Click
    Me.DbCCodEmpleado.SetFocus
     Me.DbCCodEmpleado.SelStart = 0
     Me.DbCCodEmpleado.SelLength = Len(DbCCodEmpleado.Text)
End If
End Sub

Private Sub TxtHrasExtras_LostFocus()
TxtHrasExtras = Format(TxtHrasExtras, "###,##0.00")
End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub
