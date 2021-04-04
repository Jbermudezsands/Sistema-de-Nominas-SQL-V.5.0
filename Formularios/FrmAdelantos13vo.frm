VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form FrmAdelantos13vo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adelanto de 13vo Mes y Vacaciones"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "FrmAdelantos13vo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6915
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Borrar Linea"
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   4560
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo DBCodigo2 
      Bindings        =   "FrmAdelantos13vo.frx":1E72
      Height          =   315
      Left            =   3600
      TabIndex        =   11
      Top             =   6600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodEmpleado1"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaAdelanto 
      Height          =   375
      Left            =   240
      Top             =   7080
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
   Begin MSAdodcLib.Adodc DtaBusca 
      Height          =   375
      Left            =   240
      Top             =   7560
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
      Caption         =   "DtaBusca"
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
      Left            =   3240
      Top             =   7080
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
   Begin MSComCtl2.MonthView Fecha 
      Height          =   2370
      Left            =   2160
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   81592321
      CurrentDate     =   38423
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   4080
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3975
      Begin TrueOleDBList80.TDBCombo DBCodigo 
         Bindings        =   "FrmAdelantos13vo.frx":1E8C
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
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
         _PropDict       =   $"FrmAdelantos13vo.frx":1EA7
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
      Begin VB.TextBox TxtNombre1 
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox TxtApellido1 
         Height          =   305
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox TxtCargo 
         Height          =   305
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Numero Empleado"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre Empleado"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Apellido Empleado"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Cargo Empleado"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Foto del Empleado"
      Height          =   2055
      Left            =   4080
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1575
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2565
      End
   End
   Begin TrueOleDBGrid70.TDBGrid DBGAdelantos 
      Bindings        =   "FrmAdelantos13vo.frx":1F51
      Height          =   2295
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4048
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
   Begin MSAdodcLib.Adodc AdoEmpleados 
      Height          =   375
      Left            =   240
      Top             =   6480
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
End
Attribute VB_Name = "FrmAdelantos13vo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBorrarLinea_Click()

End Sub

Private Sub CmdCerrar_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim Respuesta As String

Respuesta = MsgBox("Desea Eliminar la Linea?", vbYesNo, "Sistema de Nominas")
 
If Respuesta = "6" Then
 Me.DtaAdelanto.Recordset.Delete
 Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo Where (((Adelanto13vo.CodEmpleado) = '" & DBCodigo.Text & "'))"
   Me.DtaAdelanto.Refresh
   Me.DBGAdelantos.Columns(0).Visible = False
   Me.DBGAdelantos.Columns(2).NumberFormat = "##,##0.00"
   Me.DBGAdelantos.Columns(1).NumberFormat = "dd/mm/yyyy"
   Me.DBGAdelantos.Columns(4).Button = True
   Me.DBGAdelantos.Enabled = True
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DBCodigo_Change()
On Error GoTo TipoErrs
Dim Fecha As Date, HoraEntra As Variant
Destino = ""
DtaBusca.RecordSource = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Cargo.Cargo, Departamento.Departamento FROM Departamento INNER JOIN (Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) ON Departamento.CodDepartamento = Empleado.CodDepartamento WHERE (((Empleado.CodEmpleado)='" & DBCodigo.Columns(0).Text & "'))"
DtaBusca.Refresh
If Not DtaBusca.Recordset.EOF Then
   TxtNombre1.Text = DtaBusca.Recordset("Nombre1")
   TxtApellido1.Text = DtaBusca.Recordset("Apellido1")
   TxtCargo.Text = DtaBusca.Recordset("Cargo")
   Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo Where (((Adelanto13vo.CodEmpleado) = '" & DBCodigo.Columns(0).Text & "'))"
   Me.DtaAdelanto.Refresh
   Me.DBGAdelantos.Columns(0).Visible = False
   Me.DBGAdelantos.Columns(2).NumberFormat = "##,##0.00"
   Me.DBGAdelantos.Columns(1).NumberFormat = "dd/mm/yyyy"
   Me.DBGAdelantos.Columns(4).Button = True
   Me.DBGAdelantos.Columns(1).Button = True
   Me.DBGAdelantos.Enabled = True
    
   If Dir(App.Path + "\Fotos\" & DBCodigo.Text & ".jpg") <> "" Then
           Destino = App.Path + "\Fotos\" & DBCodigo.Text & ".jpg"
        ElseIf Dir(App.Path + "\Fotos\" & DBCodigo.Text & ".gif") <> "" Then
           Destino = App.Path + "\Fotos\" & DBCodigo.Text & ".gif"
        ElseIf Dir(App.Path + "\Fotos\" & DBCodigo.Text & ".bmp") <> "" Then
           Destino = App.Path + "\Fotos\" & DBCodigo.Text & ".bmp"
   End If
        
     If Destino <> "" Then
         Image1.Picture = LoadPicture(Destino)
        Else
         Destino = App.Path + "\Fotos\Zw.bmp"
         Image1.Picture = LoadPicture(Destino)
     End If

Else
   TxtCargo.Text = ""
   TxtNombre1.Text = ""
   TxtApellido1.Text = ""
   TxtCargo.Text = "salida"
End If
Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub DBCodigo_ItemChange()
On Error GoTo TipoErrs
Dim Fecha As Date, HoraEntra As Variant
Destino = ""
DtaBusca.RecordSource = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Cargo.Cargo, Departamento.Departamento FROM Departamento INNER JOIN (Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) ON Departamento.CodDepartamento = Empleado.CodDepartamento WHERE (((Empleado.CodEmpleado)='" & DBCodigo.Columns(0).Text & "'))"
DtaBusca.Refresh
If Not DtaBusca.Recordset.EOF Then
   TxtNombre1.Text = DtaBusca.Recordset("Nombre1")
   TxtApellido1.Text = DtaBusca.Recordset("Apellido1")
   TxtCargo.Text = DtaBusca.Recordset("Cargo")
   Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo Where (((Adelanto13vo.CodEmpleado) = '" & DBCodigo.Columns(0).Text & "'))"
   Me.DtaAdelanto.Refresh
   Me.DBGAdelantos.Columns(0).Visible = False
   Me.DBGAdelantos.Columns(2).NumberFormat = "##,##0.00"
   Me.DBGAdelantos.Columns(1).NumberFormat = "dd/mm/yyyy"
   Me.DBGAdelantos.Columns(4).Button = True
   Me.DBGAdelantos.Columns(1).Button = True
   Me.DBGAdelantos.Enabled = True
   
   
    
   If Dir(App.Path + "\Fotos\" & DBCodigo.Text & ".jpg") <> "" Then
           Destino = App.Path + "\Fotos\" & DBCodigo.Text & ".jpg"
        ElseIf Dir(App.Path + "\Fotos\" & DBCodigo.Text & ".gif") <> "" Then
           Destino = App.Path + "\Fotos\" & DBCodigo.Text & ".gif"
        ElseIf Dir(App.Path + "\Fotos\" & DBCodigo.Text & ".bmp") <> "" Then
           Destino = App.Path + "\Fotos\" & DBCodigo.Text & ".bmp"
   End If
        
     If Destino <> "" Then
         Image1.Picture = LoadPicture(Destino)
        Else
         Destino = App.Path + "\Fotos\Zw.bmp"
         Image1.Picture = LoadPicture(Destino)
     End If

Else
   TxtCargo.Text = ""
   TxtNombre1.Text = ""
   TxtApellido1.Text = ""
   TxtCargo.Text = "salida"
End If
Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub DBGAdelantos_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
 Me.DBGAdelantos.Columns(0) = Me.DBCodigo.Columns(0).Text


End Sub

Private Sub DBGAdelantos_ButtonClick(ByVal ColIndex As Integer)
Dim c As Variant
 Select Case ColIndex
    Case 1
       Set c = DBGAdelantos.Columns(ColIndex)
      With Me.Fecha
      '.Left = Me.DBGAdelantos.Left + c.Left
      '.Top = DBGAdelantos.Top + DBGAdelantos.RowTop(DBGAdelantos.Row) + DBGAdelantos.RowHeight + 500
      '.Width = c. + 15
      .Value = Now
      .Visible = True
      .SetFocus
      End With
      ColIndexs = 4
     Fecha.Value = Now
 
    Case 4
      Set c = DBGAdelantos.Columns(ColIndex)
      With List1
      .Left = Me.DBGAdelantos.Left + c.Left
      .Top = DBGAdelantos.Top + DBGAdelantos.RowTop(DBGAdelantos.Row) + DBGAdelantos.RowHeight + 500
      .Width = c.Width + 15
      .Visible = True
      .SetFocus
      End With
      ColIndexs = 4
   
 End Select
End Sub

Private Sub DtaAdelanto_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Fecha_DateClick(ByVal DateClicked As Date)
Me.DBGAdelantos.Columns(1).Text = Me.Fecha.Value
Me.Fecha.Visible = False
End Sub

Private Sub Fecha_LostFocus()
 Me.Fecha.Visible = False
End Sub

Private Sub Form_Load()
'Me.DtaBusca '.DatabaseName = Ruta
'Me.DtaEmpleado '.DatabaseName = Ruta
'Me.DtaAdelanto '.DatabaseName = Ruta
Me.DtaBusca.ConnectionString = Conexion
With Me.DtaEmpleado
  .ConnectionString = Conexion
  .RecordSource = "Empleado"
  .Refresh
End With

With Me.AdoEmpleados
  .ConnectionString = Conexion
  .RecordSource = "SELECT CodEmpleado,CodEmpleado1, Nombre1 + N' ' + Nombre2 + N' ' + Apellido1 + N' ' + Apellido2 AS Nombres, Activo From Empleado Where (Activo = 1) ORDER BY CodEmpleado1 ASC"
  .Refresh
End With

Me.DBCodigo.RowSource = Me.AdoEmpleados
Me.DBCodigo.Columns(0).Visible = False

Me.DtaAdelanto.ConnectionString = Conexion
Me.Fecha.Value = Now

'Me.DBGAdelantos.DataSource = "DtaAdelanto"

Me.DtaAdelanto.RecordSource = "SELECT Adelanto13vo.CodEmpleado, Adelanto13vo.FechaAdelanto, Adelanto13vo.MontoAdelanto, Adelanto13vo.[Ref/Cheque], Adelanto13vo.TipoAdelanto From Adelanto13vo Where (((Adelanto13vo.CodEmpleado) = '1'))"
Me.DtaAdelanto.Refresh
Me.DBGAdelantos.Columns(0).Visible = False
Me.DBGAdelantos.Columns(2).NumberFormat = "##,##0.00"
Me.DBGAdelantos.Columns(4).Button = True
Me.DBGAdelantos.Columns(1).Button = True
Me.List1.AddItem ("Vacaciones")
Me.List1.AddItem ("13vo Mes")
End Sub

Private Sub MacButton1_Click()

End Sub

Private Sub List1_DblClick()
Me.DBGAdelantos.Columns(4).Text = Me.List1.Text
Me.List1.Visible = False
End Sub

Private Sub List1_LostFocus()
'DBGDetalleTiket.Columns(4).Text = Me.List1.Text
 Me.List1.Visible = False
End Sub
