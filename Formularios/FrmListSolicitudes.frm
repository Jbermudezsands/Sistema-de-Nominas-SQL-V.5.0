VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmListSolicitudes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registros de Grupos para Fichas"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListSolicitudes.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListSolicitudes.frx":0886
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoNomina 
      Height          =   375
      Left            =   6240
      Top             =   4800
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
      Caption         =   "AdoNomina"
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
      Left            =   1920
      Top             =   4680
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
      Caption         =   "AdoConsecutivo"
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
      Left            =   5040
      Top             =   2880
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
      Caption         =   "AdoSolicitud"
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
      Height          =   375
      Left            =   2160
      Top             =   5160
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
   Begin VB.Frame Frame2 
      Caption         =   "Datos de la Solicitud"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3855
      Begin XtremeSuiteControls.DateTimePicker DtpFechaSolicitud 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   41714.4187384259
      End
      Begin XtremeSuiteControls.ComboBox CmbTipoSolicitud 
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   840
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox CmbClasificado 
         Height          =   315
         Left            =   1320
         TabIndex        =   16
         Top             =   1200
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         Appearance      =   6
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Solicitud"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Solicitud"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Clasificacion"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1695
      Left            =   3960
      TabIndex        =   11
      Top             =   120
      Width           =   6495
      Begin VB.TextBox txtHorasSolicitud 
         Height          =   285
         Left            =   3720
         TabIndex        =   15
         Text            =   "00"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox TxtDiasSolicitud 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "1"
         Top             =   720
         Width           =   735
      End
      Begin XtremeSuiteControls.DateTimePicker DtpFechaInicio 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   41714.4187384259
      End
      Begin XtremeSuiteControls.DateTimePicker DtpFechaFin 
         Height          =   315
         Left            =   3720
         TabIndex        =   14
         Top             =   240
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   41714.4187384259
      End
      Begin XtremeSuiteControls.ProgressBar pbEmpleados 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   6255
         _Version        =   786432
         _ExtentX        =   11033
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   14737632
         Scrolling       =   1
         Appearance      =   6
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Horas a Disfrutar"
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Dias a Disfrutar"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin"
         Height          =   255
         Left            =   2880
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc AdoDepartamento 
      Height          =   375
      Left            =   840
      Top             =   4200
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc AdoEmpleados 
      Height          =   375
      Left            =   840
      Top             =   3600
      Visible         =   0   'False
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
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   10560
      TabIndex        =   1
      Top             =   120
      Width           =   2775
      Begin XtremeSuiteControls.PushButton CmdGrabar 
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   4920
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Procesar"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmListSolicitudes.frx":10B8
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdFiltrar 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   4440
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Filtrar"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmListSolicitudes.frx":341C
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.RadioButton OptDepartamentoTodos 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Departamento"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptButtonTodos 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   735
         _Version        =   786432
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Todos"
         UseVisualStyle  =   -1  'True
      End
      Begin TrueOleDBList80.TDBCombo cboDepartamento 
         Bindings        =   "FrmListSolicitudes.frx":56EC
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   8811
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
         ListField       =   "Departamento"
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
         DropdownPosition=   1
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         AddItemSeparator=   ";"
         _PropDict       =   $"FrmListSolicitudes.frx":570A
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
      Begin XtremeSuiteControls.PushButton CmdSalir 
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   5400
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Salir"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmListSolicitudes.frx":57B4
         ImageAlignment  =   0
      End
      Begin VB.Label LblNumeroSolicitud 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label LblEmpleados 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   2295
      End
   End
   Begin TrueOleDBGrid80.TDBGrid DbgridEmpleados 
      Bindings        =   "FrmListSolicitudes.frx":5CB8
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7435
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
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=8,.bold=0,.fontsize=825,.italic=0"
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
   Begin VB.Label Label8 
      Height          =   375
      Left            =   10800
      TabIndex        =   26
      Top             =   2760
      Width           =   2295
   End
End
Attribute VB_Name = "FrmListSolicitudes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFiltrar_Click()
  LlenarGrid
End Sub

Private Sub CmdGrabar_Click()
            Dim CodigoEmpleadoR As String
            Dim Contador As Integer, Fecha As Date, NumeroSolicitud As String, Horas As Double
            Dim CantDias As Double, FechaIni As Date, FechaFin As Date, Cont As Double
             Dim rs As New ADODB.Recordset
             
  LlenarGrid
  
  
            pbEmpleados.Visible = True
            Me.LblEmpleados.Visible = True
            pbEmpleados.Min = 0
            pbEmpleados.Max = AdoEmpleados.Recordset.RecordCount

            Contador = 0
            Cont = 0
         Do While Not Me.AdoEmpleados.Recordset.EOF
           If AdoEmpleados.Recordset("Seleccion") = "Verdadero" Then
                Cont = Cont + 1
                CodigoEmpleadoR = AdoEmpleados.Recordset("CodEmpleado1")
                pbEmpleados.Value = Contador
                Me.LblEmpleados.Caption = AdoEmpleados.Recordset("Nombres")
                NumeroSolicitud = Format(ConsecutivoSolicitud, "0000#")
                
    '            Me.Caption = "Procesando el Empleado :" & CodigoEmpleadoR & " " & AdoEmpleados.Recordset("Nombres")
                DoEvents
                
                AdoNomina.RecordSource = "SELECT TipoNomina.Horas FROM TipoNomina INNER JOIN  Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina WHERE        (Empleado.CodEmpleado1 = '" & CodigoEmpleadoR & "')"
                AdoNomina.Refresh
                
                If Not Me.AdoNomina.Recordset.EOF Then
                  Horas = Me.AdoNomina.Recordset("Horas")
                End If
                        '///////////////////////////////////////
                        '///////////////////////////////////////
                        
                        
                            Fecha = Format(Me.DtpFechaSolicitud.Value, "yyyy-mm-dd")
                            Me.AdoSolicitud.RecordSource = "SELECT  * From SolicitudVacaciones WHERE (NumeroSolicitud = '" & NumeroSolicitud & "')"
                            Me.AdoSolicitud.Refresh
                            If Me.AdoSolicitud.Recordset.EOF Then
                            
                                  Me.AdoSolicitud.Recordset.AddNew
                              
                                    Me.AdoSolicitud.Recordset("FechaSolicitud") = Format(Me.DtpFechaSolicitud.Value, "dd/mm/yyyy")
                                    Me.AdoSolicitud.Recordset("NumeroSolicitud") = NumeroSolicitud
                                    Me.AdoSolicitud.Recordset("TipoSolicitud") = Me.CmbTipoSolicitud.Text
                                    Me.AdoSolicitud.Recordset("CodigoEmpleado") = CodigoEmpleadoR
                                    Me.AdoSolicitud.Recordset("DiasVacaciones") = 0
                                    Me.AdoSolicitud.Recordset("FechaInicio") = Me.DtpFechaInicio.Value
                                    Me.AdoSolicitud.Recordset("FechaFin") = Me.DtpFechaFin.Value
                                    
                                     If CDbl(Me.txtHorasSolicitud.Text) = 0 And CDbl(Me.TxtDiasSolicitud.Text) >= 1 Then 'Significa que son Dias
                                            Me.AdoSolicitud.Recordset("DiasDisfrutar") = CDbl(Me.TxtDiasSolicitud.Text)
                                            Me.AdoSolicitud.Recordset("DiasDisfrutados") = CDbl(Me.TxtDiasSolicitud.Text)
                                     Else 'Significa que son horas
                                            If CDbl(TxtDiasSolicitud.Text) = 1 And CDbl(txtHorasSolicitud.Text) > 0 Then
                                            Me.AdoSolicitud.Recordset("DiasDisfrutar") = CDbl(Me.txtHorasSolicitud.Text) / Horas
                                            Me.AdoSolicitud.Recordset("DiasDisfrutados") = CDbl(Me.txtHorasSolicitud.Text) / Horas
                                            End If
                                     End If
                                    
                                    Me.AdoSolicitud.Recordset("Observaciones") = "Proceso Automatico Solicitud"
                                    Me.AdoSolicitud.Recordset.Update
                                    
                                    Me.AdoConsecutivo.Recordset("Solicitud") = ConsecutivoSolicitud
                                    Me.AdoConsecutivo.Recordset.Update
                            End If
                            
                            
                            '///////////////////////////////////////ELIMINO LOS REGISTROS ////////////////////////////////////
                            
                            rs.Open "DELETE FROM [DescuentoDiasVacaciones] WHERE (NumeroSolicitud = '" & NumeroSolicitud & "') AND (TipoDescuento = '" & Me.CmbTipoSolicitud.Text & "') AND (CodigoEmpleado = '" & CodigoEmpleadoR & "')", Conexion
                            
                            'Supongo que elimina por si existen registros para luego actualizarlos.
                            
                            'Condicionales del DoWhile
                            FechaIni = Format(Me.DtpFechaInicio.Value, "dd/mm/yyyy")
                            FechaFin = Format(Me.DtpFechaFin.Value, "dd/mm/yyyy")
                            
                            
                            Do While FechaIni <= FechaFin
                              If Me.CmbTipoSolicitud.Text = "Vacaciones Pagadas" Then
                              
                                   If CDbl(Me.txtHorasSolicitud.Text) = 0 And CDbl(TxtDiasSolicitud.Text) >= 1 Then 'Significa que son Dias
                                       CantDias = DateDiff("d", Format(Me.DtpFechaInicio.Value, "dd/mm/yyyy"), Format(Me.DtpFechaFin.Value, "dd/mm/yyyy")) + 1
                                   Else 'Significa que son horas
                                       If CDbl(TxtDiasSolicitud.Text) = 1 And CDbl(txtHorasSolicitud.Text) > 0 Then
                                       CantDias = CDbl(txtHorasSolicitud.Text) / Horas
                                       End If
                                   End If
                                   Resultado = GrabaDescuentoDias(FechaIni, CodigoEmpleadoR, Me.CmbTipoSolicitud.Text, NumeroSolicitud, CantDias)
                                   Exit Do
                              
                               
                                  
                              Else
                               CantDias = 1
                               Resultado = GrabaDescuentoDias(FechaIni, CodigoEmpleadoR, Me.CmbTipoSolicitud.Text, NumeroSolicitud, 1)
                              
                              End If
                              FechaIni = DateAdd("d", 1, FechaIni)
                            Loop
                            
    '                    Me.TxtNumero.Text = Format(CInt(NumeroSolicitud) + 1, "0000#")
                        '///////////////////////////////////////////
                        '///////////////////////////////////////////
                        DoEvents
                        
                End If
            AdoEmpleados.Recordset.MoveNext
            Contador = Contador + 1
         Loop
         
         
         '///////////////////////////////////////////////////CAMBIO EL VALOR DE LA SELECCION //////////////////////
         rs.Open "Update [dbo].[Empleado] Set [HorasTurno] = 0", Conexion
         MsgBox "Proceso Terminado se agregaron " & Cont & " Empleados", vbInformation, "Zeus Nominas"
         
End Sub

Private Sub CmdSalir_Click()
Dim rs As New ADODB.Recordset
rs.Open "Update [dbo].[Empleado] Set [HorasTurno] = 0", Conexion

Unload Me
End Sub

Private Sub DtpFechaFin_Change()
  Me.TxtDiasSolicitud.Text = DateDiff("d", Me.DtpFechaInicio.Value, Me.DtpFechaFin.Value) + 1
End Sub

Private Sub DtpFechaInicio_Change()
  Me.TxtDiasSolicitud.Text = DateDiff("d", Me.DtpFechaInicio.Value, Me.DtpFechaFin.Value) + 1
End Sub

Private Sub Form_Load()
 Me.AdoEmpleados.ConnectionString = Conexion
 
  Me.DbgridEmpleados.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DbgridEmpleados.OddRowStyle.BackColor = &H80000005
 Me.DbgridEmpleados.AlternatingRowStyle = True
 
' Me.Picture1.BackColor = RGB(222, 227, 247)
 Me.Frame1.BackColor = RGB(222, 227, 247)
 Me.OptButtonTodos.BackColor = RGB(222, 227, 247)
 Me.OptDepartamentoTodos.BackColor = RGB(222, 227, 247)
 Me.Frame2.BackColor = RGB(222, 227, 247)
 Me.Frame3.BackColor = RGB(222, 227, 247)
 
 Me.DtpFechaFin.Value = Now
 Me.DtpFechaInicio.Value = Now
 Me.DtpFechaSolicitud.Value = Now
 
 With Me.AdoDepartamento
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT CodDepartamento, Departamento From departamento"
   .Refresh
End With

 With Me.CmbTipoSolicitud
    .AddItem "Vacaciones"
    .AddItem "Vacaciones Pagadas"
    .AddItem "Vacaciones Programadas"
    .AddItem "Subsidio"
    .AddItem "Ausente"
    .AddItem "Feriado"
    .AddItem "Permiso Programado"
    .AddItem "Suspension"
    .Text = "Vacaciones"
 End With
 
With Me.CmbClasificado
    .AddItem "Riesgo Laboral"
    .AddItem "Accidente Comun"
    .AddItem "Embarazo"
    .AddItem "Enfermedades"
    .Text = "Riesgo Laboral"
End With

With Me.AdoConsecutivo
   .ConnectionString = Conexion
   .RecordSource = "SELECT  * From Consecutivos"
   .Refresh
End With



'With Me.AdoAux
'   .ConnectionString = Conexion
'End With
'
With Me.AdoNomina
   .ConnectionString = Conexion
End With
'With Me.AdoAuxiliar
'   .ConnectionString = Conexion
'End With

With Me.AdoSolicitud
   .ConnectionString = Conexion
End With


Me.OptButtonTodos.Value = True
LlenarGrid
 

End Sub

Public Sub LlenarGrid()
Dim SqlString As String, CodDepartamento As String

    Dim item As TrueDBGrid80.ValueItem

    CodDepartamento = Me.cboDepartamento.Columns("CodDepartamento").Text

    If Me.OptButtonTodos.Value = True Then
     SqlString = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 As Nombres, Empleado.CodEmpleado, departamento.CodDepartamento , departamento.departamento, HorasTurno As Seleccion FROM  Empleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento Where (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1"
    ElseIf Me.OptDepartamentoTodos.Value = True Then
     SqlString = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres,  Empleado.CodEmpleado , departamento.CodDepartamento, departamento.departamento, HorasTurno As Seleccion FROM Empleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento WHERE  (Empleado.Activo = 1) AND (Departamento.CodDepartamento = '" & CodDepartamento & "') ORDER BY Empleado.CodEmpleado1"
    End If
    
    Me.AdoEmpleados.RecordSource = SqlString
    Me.AdoEmpleados.ConnectionString = Conexion
    Me.AdoEmpleados.Refresh
    

'    With Me.DbgridEmpleados.Columns("Seleccion").ValueItems
'
'         item = New TrueDBGrid80.ValueItem
'         item.Value = "False"
'         item.DisplayValue = Me.ImageList1.ListImages(1).Picture
'
'         item = New TrueDBGrid80.ValueItem
'         item.Value = "True"
'         item.DisplayValue = Me.ImageList1.ListImages(2).Picture
'
''            item.Value = "False"
''            item.DisplayValue = Me.ImageList.Images(1)
''            .Add (item)
''
''            item = New C1.Win.C1TrueDBGrid.ValueItem()
''            item.Value = "True"
''            item.DisplayValue = Me.ImageList.Images(0)
''            .Add (item)
'
'            Me.DbgridEmpleados.Columns("Seleccion").ValueItems.Translate = True
'    End With
    
    Me.DbgridEmpleados.DataSource = Me.AdoEmpleados
    Me.DbgridEmpleados.Columns("CodEmpleado").Visible = False
    Me.DbgridEmpleados.Columns("CodDepartamento").Visible = False
    Me.DbgridEmpleados.Columns("Nombres").Width = 5500
    Me.DbgridEmpleados.Columns("departamento").Width = 1500
    Me.DbgridEmpleados.Columns("CodEmpleado").Locked = True
    Me.DbgridEmpleados.Columns("Nombres").Locked = True
    Me.DbgridEmpleados.Columns("CodEmpleado1").Locked = True
    Me.DbgridEmpleados.Columns("CodDepartamento").Locked = True
    Me.DbgridEmpleados.Columns("departamento").Locked = True
    Me.DbgridEmpleados.Columns("Seleccion").Locked = False
    
    Me.DbgridEmpleados.Columns("Seleccion").ValueItems.Presentation = dbgCheckBox

    
    
    
    
    
    
End Sub

Private Sub OptButtonTodos_Click()
 If Me.OptDepartamentoTodos.Value = True Then
   Me.cboDepartamento.Visible = True
 Else
   Me.cboDepartamento.Visible = False
 End If
 
 LlenarGrid
End Sub

Private Sub OptDepartamentoTodos_Click()
 If Me.OptDepartamentoTodos.Value = True Then
   Me.cboDepartamento.Visible = True
 Else
   Me.cboDepartamento.Visible = False
 End If
 
 LlenarGrid

End Sub
