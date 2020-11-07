VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmAsistenciaAgregarRegistros 
   Caption         =   "Regsitros Colectivos"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin TrueOleDBList80.TDBCombo cboDepartamento 
      Bindings        =   "FrmAsistenciaAgregarRegistros.frx":0000
      Height          =   315
      Left            =   7560
      TabIndex        =   13
      Top             =   2160
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
      _PropDict       =   $"FrmAsistenciaAgregarRegistros.frx":001F
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
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Agregar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.ComboBox CmbTipo 
      Height          =   315
      ItemData        =   "FrmAsistenciaAgregarRegistros.frx":00C9
      Left            =   1680
      List            =   "FrmAsistenciaAgregarRegistros.frx":00D3
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1680
      Width           =   2415
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6615
      _Version        =   786432
      _ExtentX        =   11668
      _ExtentY        =   2143
      _StockProps     =   79
      Caption         =   "Empleados"
      UseVisualStyle  =   -1  'True
      Begin VB.CheckBox ChkTodos 
         Caption         =   "Todos los Empleados"
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   360
         Width           =   2775
      End
      Begin TrueOleDBList80.TDBCombo cboCodigo 
         Bindings        =   "FrmAsistenciaAgregarRegistros.frx":00E8
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   360
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
         DropdownPosition=   1
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         AddItemSeparator=   ";"
         _PropDict       =   $"FrmAsistenciaAgregarRegistros.frx":0102
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lblNombre 
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   6015
      End
   End
   Begin MSMask.MaskEdBox mskPermisoHoraInicio 
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Format          =   "hh:mm AM/PM"
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker dtpFechEntrada 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   83034113
      CurrentDate     =   38570
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cancelar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   6615
      _Version        =   786432
      _ExtentX        =   11668
      _ExtentY        =   661
      _StockProps     =   93
      Appearance      =   6
   End
   Begin MSAdodcLib.Adodc adoEmpleado 
      Height          =   375
      Left            =   360
      Top             =   5400
      Width           =   4695
      _ExtentX        =   8281
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
      Caption         =   "adoEmpleado"
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
   Begin MSAdodcLib.Adodc adoAsistencia 
      Height          =   330
      Left            =   360
      Top             =   4920
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   582
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
      Connect         =   $"FrmAsistenciaAgregarRegistros.frx":01AC
      OLEDBString     =   $"FrmAsistenciaAgregarRegistros.frx":0238
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Asistencia Diaria"
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
   Begin MSAdodcLib.Adodc AdoDepartamentos 
      Height          =   375
      Left            =   240
      Top             =   4440
      Width           =   4695
      _ExtentX        =   8281
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
      Caption         =   "AdoDepartamentos"
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
   Begin MSAdodcLib.Adodc AdoTipoNomina 
      Height          =   375
      Left            =   360
      Top             =   6000
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
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
      Caption         =   "AdoTipoNomina"
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
   Begin VB.Label Label3 
      Caption         =   "Fecha y Hora"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
End
Attribute VB_Name = "FrmAsistenciaAgregarRegistros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChkTodos_Click()
 If Me.ChkTodos.Value = 1 Then
   Me.cboCodigo.Text = ""
   Me.cboCodigo.Enabled = False
  Else
     Me.cboCodigo.Enabled = True
 End If
End Sub

Private Sub Form_Load()
Me.dtpFechEntrada.Value = Format(Now, "dd/mm/yyyy")
Me.CmbTipo.Text = "Entrada"

With Me.adoEmpleado
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT CodEmpleado1, Nombre1 + ' '+ Nombre2 +' '+Apellido1+' '+Apellido2 as Nombres From Empleado Where (Activo = 1) ORDER BY CodEmpleado1"
   .Refresh
End With

With Me.adoAsistencia
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "AsistenciaEmpleado"
   .Refresh
End With

With Me.AdoDepartamentos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT CodDepartamento, Departamento From departamento"
   .Refresh
End With

With Me.AdoTipoNomina
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT CodTipoNomina, Nomina From TipoNomina"
   .Refresh
End With


End Sub

Private Sub PushButton1_Click()
Dim dFecha As Date
Dim sFechaEntrada As String
'Dim cnDB As New ADODB.Connection
'Dim rsDB As New ADODB.Recordset
Dim dCodigoEmpl As Double, SqlString As String

'cnDB.ConnectionString = "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemasNominas;Data Source=METRO"
'cnDB.Open


If Me.ChkTodos.Value = 0 Then
  SqlString = "SELECT * FROM Empleado WHERE CodEmpleado1 LIKE '" & Me.cboCodigo.Text & "' AND Activo =1"
ElseIf FrmAsistencias.OptDepartamento.Value = True Then
  SqlString = "SELECT  Empleado.* From Empleado WHERE (Activo = 1) AND (CodDepartamento = '" & FrmAsistencias.cboDepartamento.Columns(0).Text & "')"
ElseIf FrmAsistencias.OptTipoNomina.Value = True Then
  SqlString = "SELECT  Empleado.* From Empleado WHERE (Activo = 1) AND (CodTipoNomina = '" & FrmAsistencias.cboTipoNomina.Columns(0).Text & "')"
  
Else
  SqlString = "SELECT * FROM Empleado WHERE Activo =1"
End If

'     Me.adoEmpleado.CommandType = adCmdText
     Me.adoEmpleado.RecordSource = SqlString
     Me.adoEmpleado.Refresh
     
     If Not Me.adoEmpleado.Recordset.EOF Then
        Me.adoEmpleado.Recordset.MoveLast
        
        Me.ProgressBar1.Min = 0
        Me.ProgressBar1.Max = Me.adoEmpleado.Recordset.RecordCount
        Me.ProgressBar1.Value = 0
        Me.adoEmpleado.Recordset.MoveFirst
     End If

     Do While Not Me.adoEmpleado.Recordset.EOF
   
                  dFecha = Me.dtpFechEntrada.Value
                  sCodEmpl = Me.adoEmpleado.Recordset("CodEmpleado1")
                  sCodTipoNomina = Me.adoEmpleado.Recordset("CodTipoNomina")
                  
                  Me.Caption = "Procesando " & sCodEmpl
                  DoEvents
                  sFechaEntrada = Mid$(dFecha, 7, 4) & "-" & Mid$(dFecha, 4, 2) & "-" & Mid$(dFecha, 1, 2)
                  
                  
                  Me.adoAsistencia.CommandType = adCmdText
                  Me.adoAsistencia.RecordSource = "SELECT CodEmpleado, CodEmpleado1, CodTipoNomina, FechaEntrada, FechaSalida, HoraEntrada, HREntrada, HRSalida, HoraSalida, bActivo, CodTurno, HLaboradas, HExtras " & _
                                                 "FROM AsistenciaEmpleado WHERE FechaEntrada = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102) AND CodEmpleado1 ='" & sCodEmpl & "'"
                  Me.adoAsistencia.Refresh
                                   
            
                 
                    
                    
                    If Me.adoAsistencia.Recordset.EOF Then
                  
                         Me.adoAsistencia.Recordset.AddNew
                         
                         Me.adoAsistencia.Recordset.Fields("CodEmpleado") = Me.adoEmpleado.Recordset.Fields("CodEmpleado")
                         Me.adoAsistencia.Recordset.Fields("CodEmpleado1") = sCodEmpl
                        
                         Me.adoAsistencia.Recordset.Fields("CodTipoNomina") = sCodTipoNomina
                         
                         If Me.CmbTipo.Text = "Entrada" Then
                            Me.adoAsistencia.Recordset.Fields("FechaEntrada") = Me.dtpFechEntrada.Value
                            Me.adoAsistencia.Recordset.Fields("HoraEntrada") = Me.mskPermisoHoraInicio.Text
                            If IsNull(Me.adoAsistencia.Recordset.Fields("HREntrada")) Then
                              Me.adoAsistencia.Recordset.Fields("HREntrada") = Me.mskPermisoHoraInicio.Text
                            End If
                         ElseIf Me.CmbTipo.Text = "Salida" Then
                          If Not IsNull(Me.adoAsistencia.Recordset.Fields("FechaEntrada")) Then
                            Me.adoAsistencia.Recordset.Fields("FechaSalida") = Me.dtpFechEntrada.Value
                            Me.adoAsistencia.Recordset.Fields("HoraSalida") = Me.mskPermisoHoraInicio.Text
                            If IsNull(Me.adoAsistencia.Recordset.Fields("HRSalida")) Then
                              Me.adoAsistencia.Recordset.Fields("HRSalida") = Me.mskPermisoHoraInicio.Text
                            End If
                          End If
                            

                          
'                            Me.adoAsistencia.Recordset.Fields("bActivo") = 1
                         
                         End If
                         
                         
                '         If Not Me.chkSalidaManual.Value And Chequear_Hora(Me.mskPermisoHoraRegreso.Text) Then
                '            Me.adoAsistencia.Recordset.Fields("FechaSalida") = Me.dtpFechEntrada.Value
                '            Me.adoAsistencia.Recordset.Fields("HoraSalida") = Me.mskPermisoHoraRegreso.Text
                '            Me.adoAsistencia.Recordset.Fields("HRSalida") = Me.mskPermisoHoraRegreso.Text
                '            Me.adoAsistencia.Recordset.Fields("bActivo") = 0
                '         ElseIf Not Me.chkSalida.Value Then
                '            MsgBox "Debe de digitar la hora de salida correcta, verifique"
                '            Me.adoAsistencia.Recordset.CancelUpdate
                '            Exit Sub
                '         Else
                '            Me.adoAsistencia.Recordset.Fields("bActivo") = 1
                '         End If
                          
                         Me.adoAsistencia.Recordset.Fields("bActivo") = 1
                         Me.adoAsistencia.Recordset.Fields("CodTurno") = "Diurno"
                         Me.adoAsistencia.Recordset.Update
                     
                  Else
                  
                         Me.adoAsistencia.Recordset.Fields("CodEmpleado") = Me.adoEmpleado.Recordset.Fields("CodEmpleado")
                         Me.adoAsistencia.Recordset.Fields("CodEmpleado1") = sCodEmpl
                        
                         Me.adoAsistencia.Recordset.Fields("CodTipoNomina") = sCodTipoNomina
                         
                         If Me.CmbTipo.Text = "Entrada" Then
                            Me.adoAsistencia.Recordset.Fields("FechaEntrada") = Me.dtpFechEntrada.Value
                            Me.adoAsistencia.Recordset.Fields("HoraEntrada") = Me.mskPermisoHoraInicio.Text
                            If IsNull(Me.adoAsistencia.Recordset.Fields("HREntrada")) Then
                              Me.adoAsistencia.Recordset.Fields("HREntrada") = Me.mskPermisoHoraInicio.Text
                            End If
                         ElseIf Me.CmbTipo.Text = "Salida" Then
                          If Not IsNull(Me.adoAsistencia.Recordset.Fields("FechaEntrada")) Then
                            Me.adoAsistencia.Recordset.Fields("FechaSalida") = Me.dtpFechEntrada.Value
                            Me.adoAsistencia.Recordset.Fields("HoraSalida") = Me.mskPermisoHoraInicio.Text
                            If IsNull(Me.adoAsistencia.Recordset.Fields("HRSalida")) Then
                              Me.adoAsistencia.Recordset.Fields("HRSalida") = Me.mskPermisoHoraInicio.Text
                            End If
                          End If
                         End If
                         
                         
                         
                         Me.adoAsistencia.Recordset.Fields("bActivo") = 1
                         Me.adoAsistencia.Recordset.Fields("CodTurno") = "Diurno"
                         Me.adoAsistencia.Recordset.Update
                  
'                    If Me.ChkTodos.Value = 0 Then
'                     MsgBox "Se tiene registrada una asistencia de este empleado para este dia, modifique el registro"
'                    Else
'
'
'                    End If
                        
                  End If
      
      DoEvents
      Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
      Me.adoEmpleado.Recordset.MoveNext
    Loop
      

'CmdConsultar_Click


End Sub

Private Sub PushButton2_Click()
Unload Me
End Sub
