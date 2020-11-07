VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.TaskPanel.v12.0.0.Demo.ocx"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#12.0#0"; "Codejock.Calendar.v12.0.0.Demo.ocx"
Begin VB.Form FrmControlVacaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario de Vacaciones"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16935
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   16935
   Begin XtremeTaskPanel.TaskPanel TaskPanel1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   2295
      _Version        =   786432
      _ExtentX        =   4048
      _ExtentY        =   2143
      _StockProps     =   64
      ItemLayout      =   2
      HotTrackStyle   =   1
      MultiColumn     =   -1  'True
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmControlVacaciones.frx":0000
      TabIndex        =   38
      Top             =   4680
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblFechaFin 
      Height          =   255
      Left            =   1320
      OleObjectBlob   =   "FrmControlVacaciones.frx":0076
      TabIndex        =   37
      Top             =   5160
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblFechaInicio 
      Height          =   255
      Left            =   1200
      OleObjectBlob   =   "FrmControlVacaciones.frx":00D4
      TabIndex        =   36
      Top             =   4680
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3675
      ScaleWidth      =   2475
      TabIndex        =   17
      Top             =   840
      Width           =   2535
      Begin VB.TextBox TxtDiasSubsidio 
         Height          =   375
         Left            =   1800
         TabIndex        =   40
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox TxtVacaciones 
         Height          =   375
         Left            =   1800
         TabIndex        =   23
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox TxtAusente 
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox TxtFeriado 
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox TxtVacPagadas 
         Height          =   375
         Left            =   1800
         TabIndex        =   20
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox TxtVacProgramadas 
         Height          =   375
         Left            =   1800
         TabIndex        =   19
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox TxtPermisoProgramado 
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   2640
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderStyle     =   4  'Dash-Dot
         X1              =   2520
         X2              =   0
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label15 
         BackColor       =   &H0080C0FF&
         Height          =   375
         Left            =   1320
         TabIndex        =   42
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Subsidios"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Vacaciones"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   1320
         TabIndex        =   34
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Ausente"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF80FF&
         Height          =   375
         Left            =   1320
         TabIndex        =   32
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Feriado"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H008888FB&
         Height          =   375
         Left            =   1320
         TabIndex        =   30
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Vac.Pagada"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000C000&
         Height          =   375
         Left            =   1320
         TabIndex        =   28
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Vac.Programda"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080FFFF&
         Height          =   375
         Left            =   1320
         TabIndex        =   26
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Perm.Programdo"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Height          =   375
         Left            =   1320
         TabIndex        =   24
         Top             =   2640
         Width           =   375
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   735
      Left            =   2520
      TabIndex        =   7
      Top             =   0
      Width           =   13815
      _Version        =   786432
      _ExtentX        =   24368
      _ExtentY        =   1296
      _StockProps     =   79
      ForeColor       =   -2147483630
      UseVisualStyle  =   -1  'True
      Begin ACTIVESKINLibCtl.SkinLabel LblNombres 
         Height          =   375
         Left            =   5280
         OleObjectBlob   =   "FrmControlVacaciones.frx":0132
         TabIndex        =   11
         Top             =   240
         Width           =   8295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmControlVacaciones.frx":0190
         TabIndex        =   10
         Top             =   240
         Width           =   855
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
         Left            =   4680
         Picture         =   "FrmControlVacaciones.frx":01FC
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin TrueOleDBList80.TDBCombo DBCodigoEmpleado 
         Bindings        =   "FrmControlVacaciones.frx":034A
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   14102
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
         ListField       =   "CodEmpleado1"
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
         _PropDict       =   $"FrmControlVacaciones.frx":0364
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
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   6960
      Width           =   13935
      _Version        =   786432
      _ExtentX        =   24580
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Vacaciones"
      ForeColor       =   -2147483630
      UseVisualStyle  =   -1  'True
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   6240
         OleObjectBlob   =   "FrmControlVacaciones.frx":040E
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "FrmControlVacaciones.frx":0490
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "FrmControlVacaciones.frx":0508
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmControlVacaciones.frx":057C
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtPendientes 
         Height          =   285
         Left            =   7680
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtPorDisfrutar 
         Height          =   285
         Left            =   5160
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtDiasDisfrutados 
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtDiasVacaciones 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin XtremeCalendarControl.DatePicker wndDatePicker 
      Height          =   6015
      Left            =   2760
      TabIndex        =   2
      Top             =   840
      Width           =   13935
      _Version        =   786432
      _ExtentX        =   24580
      _ExtentY        =   10610
      _StockProps     =   64
      AutoSize        =   0   'False
      ShowTodayButton =   0   'False
      ShowNoneButton  =   0   'False
      ShowWeekNumbers =   -1  'True
      ShowNonMonthDays=   0   'False
      RowCount        =   3
      ColumnCount     =   4
      AskDayMetrics   =   -1  'True
      FirstWeekOfYearDays=   2
   End
   Begin MSAdodcLib.Adodc DtaEmpleado 
      Height          =   375
      Left            =   1920
      Top             =   8160
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
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   375
      Left            =   5640
      Top             =   8160
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
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmControlVacaciones.frx":05F8
      TabIndex        =   39
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label1"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "FrmControlVacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FechaInicio As Date, FechaFin As Date, i As Double, DiasVacaciones As Double, DiasVacaPag As Double, DiasFeriado As Double, DiasAusente As Double, PermisoProgramado As Double, VacProgramadas As Double
Public FechaVacaciones As Date, DiasVacaAcumulados As Double, DiasSubsidio As Double

Private Sub CmdBuscarEmpleado_Click()
QueProducto = "CodigoEmpleado"
FrmConsulta.Show 1
Me.DBCodigoEmpleado.Text = FrmConsulta.CodigoEmpleado1
DBCodigoEmpleado_ItemChange
End Sub


Private Sub DBCodigoEmpleado_ItemChange()

'Me.AdoConsulta.RecordSource = "SELECT CodEmpleado,CodEmpleado1, Nombre1 + ' '+ Nombre2 +' '+Apellido1+' '+Apellido2 as Nombres, Nombre1, Nombre2, Apellido1, Apellido2, NumHijos, Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula,Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss, NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria,OtrosIngresos, PorcentajeComision, DescripOtrIngre, ExentoInss, ExentoIr, PagoInssPatronal, SalarioMinimo, Observaciones, Activo, Ausente, SalarioFijo , SumarSubsidio, PorcientoIncentivo,Dolarizado,CuentaBanco,SueldoActualBasico From Empleado WHERE (CodEmpleado1 = '" & DBCodigoEmpleado.Text & "') And (Activo = 1)"
Me.adoConsulta.RecordSource = "SELECT  Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.NumHijos, Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr, Empleado.NumCedula, Empleado.Sindicalista, Empleado.CodDepartamento, Empleado.CodGrupo, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.OtrosIngresos, Empleado.PorcentajeComision, Empleado.DescripOtrIngre, Empleado.ExentoInss, Empleado.ExentoIr, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Observaciones, Empleado.Activo, Empleado.Ausente, Empleado.SalarioFijo,  " & _
                              "Empleado.SumarSubsidio, Empleado.PorcientoIncentivo, Empleado.Dolarizado, Empleado.CuentaBanco, Empleado.SueldoActualBasico, Historico.FechaContrato FROM Empleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE (Empleado.CodEmpleado1 = '" & DBCodigoEmpleado.Text & "') AND (Empleado.Activo = 1)"
Me.adoConsulta.Refresh

If Not Me.adoConsulta.Recordset.EOF Then
       Me.LblNombres.Caption = Me.adoConsulta.Recordset("Nombres")
       FechaInicio = "01/01/1800"
       FechaFin = "01/01/1800"
       FechaVacaciones = Me.adoConsulta.Recordset("FechaContrato")
       i = 0
         DiasVacaPag = 0
         DiasVacaciones = 0
         DiasAusente = 0
         DiasFeriado = 0
         PermisoProgramado = 0
         VacProgramadas = 0
       wndDatePicker.RedrawControl
End If
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd

CreateTaskPanel
wndDatePicker.AutoSize = False
wndDatePicker.FirstDayOfWeek = 2
wndDatePicker.ColumnCount = 4
wndDatePicker.RowCount = 3
wndDatePicker.MaxSelectionCount = -1
wndDatePicker.FirstWeekOfYearDays = 7
wndDatePicker.ShowNonMonthDays = True
wndDatePicker.AskDayMetrics = True
wndDatePicker.RedrawControl

With Me.adoConsulta
   .ConnectionString = Conexion
End With

With Me.DtaEmpleado
   .ConnectionString = Conexion
   .RecordSource = "SELECT CodEmpleado, CodEmpleado1, Nombre1 + ' '+ Nombre2 +' '+Apellido1+' '+Apellido2 as Nombres From Empleado Where (Activo = 1) ORDER BY CodEmpleado1"
   .Refresh
End With

Me.DBCodigoEmpleado.Columns(0).Visible = False

End Sub



Private Sub wndDatePicker_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
  Dim Color As Long
       
    
    If Weekday(Day) = vbSunday Then
        Metrics.ForeColor = vbRed
    End If
        
    If Weekday(Day) = vbSaturday Then
        Metrics.ForeColor = vbBlue
'        Metrics.Font.Bold = True
    End If
    
   
    
    If i = 0 Then
      FechaInicio = Day
         DiasVacaPag = 0
         DiasVacaciones = 0
         DiasAusente = 0
         DiasFeriado = 0
         PermisoProgramado = 0
         VacProgramadas = 0
         DiasAcumulados = 0
         DiasVacaAcumulados = 0
         If Me.DBCodigoEmpleado.Text <> "" Then
           DiasVacaAcumulados = CalcularDiasVaca(CDate(FechaVacaciones), FechaInicio) - DiasVacaDesAcumulados(Me.DBCodigoEmpleado.Text, FechaInicio)
         End If
         
    Else
      FechaFin = Day
    End If
    
   
    
    '/////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////////////////BUSCO LOS DIA DE DESCUENTOS //////////////////////
    '////////////////////////////////////////////////////////////////////////////////////
    If Me.DBCodigoEmpleado.Text <> "" Then
        Me.adoConsulta.RecordSource = "SELECT  * From DescuentoDiasVacaciones WHERE  (FechaDescuento = CONVERT(DATETIME, '" & Format(Day, "yyyy-mm-dd") & "', 102)) AND (CodigoEmpleado = '" & Me.DBCodigoEmpleado.Text & "')"
        Me.adoConsulta.Refresh
        If Not Me.adoConsulta.Recordset.EOF Then
          Color = val(Me.adoConsulta.Recordset("Color"))
          Metrics.BackColor = Color
          TipoVacaciones = Me.adoConsulta.Recordset("TipoDescuento")
          Select Case TipoVacaciones
            Case "Vacaciones Pagadas": DiasVacaPag = DiasVacaDescuentos(Me.DBCodigoEmpleado.Text, FechaInicio, FechaFin, TipoVacaciones)
            Case "Vacaciones": DiasVacaciones = DiasVacaDescuentos(Me.DBCodigoEmpleado.Text, FechaInicio, FechaFin, TipoVacaciones)
            Case "Ausente": DiasAusente = DiasVacaDescuentos(Me.DBCodigoEmpleado.Text, FechaInicio, FechaFin, TipoVacaciones)
            Case "Feriado": DiasFeriado = DiasVacaDescuentos(Me.DBCodigoEmpleado.Text, FechaInicio, FechaFin, TipoVacaciones)
            Case "Permiso Programado": PermisoProgramado = DiasVacaDescuentos(Me.DBCodigoEmpleado.Text, FechaInicio, FechaFin, TipoVacaciones)
            Case "Vacaciones Programadas": VacProgramadas = DiasVacaDescuentos(Me.DBCodigoEmpleado.Text, FechaInicio, FechaFin, TipoVacaciones)
            Case "Subsidio": DiasSubsidio = DiasVacaDescuentos(Me.DBCodigoEmpleado.Text, FechaInicio, FechaFin, TipoVacaciones)
          End Select
          
        End If
     
         Me.TxtVacPagadas.Text = DiasVacaPag
         Me.TxtVacaciones.Text = DiasVacaciones
         Me.TxtAusente.Text = DiasAusente
         Me.TxtFeriado.Text = DiasFeriado
         Me.TxtPermisoProgramado.Text = PermisoProgramado
         Me.TxtVacProgramadas.Text = VacProgramadas
         Me.TxtDiasSubsidio.Text = DiasSubsidio
         
         Me.txtDiasdisfrutados.Text = DiasVacaPag + DiasVacaciones + DiasAusente + DiasFeriado
         Me.TxtDiasVacaciones.Text = DiasVacaAcumulados
         Me.TxtPorDisfrutar.Text = DiasVacaAcumulados - (DiasVacaPag + DiasVacaciones + DiasAusente + DiasFeriado)
         Me.TxtPendientes.Text = PermisoProgramado + VacProgramadas
         
         Me.LblFechaInicio.Caption = FechaInicio
         Me.LblFechaFin.Caption = FechaFin
    
    
    End If
    
    
   i = i + 1
        
'    If Date - Day < 4 And Date - Day > 0 Then
'        Metrics.BackColor = vbGreen
'    End If
'
'    If Day - Date = 6 Then
'        Set Metrics.Picture = pictDay
'    End If

End Sub

Sub CreateTaskPanel()

    Dim Group As TaskPanelGroup
    Dim item As TaskPanelGroupItem
    
    Set Group = Me.TaskPanel1.Groups.Add(100, "Fechas")

'    Set Group = wndTaskPanel.Groups.Add(100, "Procesos")
'    Group.Tooltip = "Sistema de Nominas"
'    Group.Special = True
'    Group.Items.Add 1, "Empleados", xtpTaskItemTypeLink, 1
'    Group.Items.Add 2, "Activar Nominas", xtpTaskItemTypeLink, 2
'    Group.Items.Add 3, "Movimiento de Produccion", xtpTaskItemTypeLink, 3
'    Group.Items.Add 4, "Horas Extras", xtpTaskItemTypeLink, 4
'    Group.Items.Add 5, "Calcular Nomina", xtpTaskItemTypeLink, 5
'    Group.Items.Add 6, "Subsidios", xtpTaskItemTypeLink, 6
    
End Sub



Private Sub wndDatePicker_MonthChanged()
  FechaInicio = "01/01/1800"
  FechaFin = "01/01/1800"
  i = 0
         DiasVacaPag = 0
         DiasVacaciones = 0
         DiasAusente = 0
         DiasFeriado = 0
         PermisoProgramado = 0
         VacProgramadas = 0
End Sub



