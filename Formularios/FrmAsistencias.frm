VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmAsistencias 
   Caption         =   "Registro de Asistencias"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   13995
   StartUpPosition =   3  'Windows Default
   Begin TrueOleDBList80.TDBCombo cboTipoNomina 
      Bindings        =   "FrmAsistencias.frx":0000
      Height          =   315
      Left            =   1200
      TabIndex        =   19
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
      Caption         =   "cboTipoNomina"
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
      DropdownPosition=   1
      Locked          =   0   'False
      ScrollTrack     =   0   'False
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      AddItemSeparator=   ";"
      _PropDict       =   $"FrmAsistencias.frx":001C
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
   Begin TrueOleDBList80.TDBCombo cboDepartamento 
      Bindings        =   "FrmAsistencias.frx":00C6
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   7080
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
      _PropDict       =   $"FrmAsistencias.frx":00E5
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
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   5055
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   13215
      _Version        =   786432
      _ExtentX        =   23310
      _ExtentY        =   8916
      _StockProps     =   68
      Appearance      =   10
      Color           =   4
      ItemCount       =   1
      Item(0).Caption =   "Asistencia Diaria"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TDBGrid"
      Begin TrueOleDBGrid80.TDBGrid TDBGrid 
         Bindings        =   "FrmAsistencias.frx":018F
         Height          =   4095
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   7223
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
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13335
      Begin VB.TextBox TxtCodigoEmpleado 
         Height          =   375
         Left            =   4440
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exportar Excel"
         Height          =   855
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   855
         Left            =   12000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar"
         Height          =   855
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton CmdConsultar 
         Caption         =   "Consultar"
         Height          =   855
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin XtremeSuiteControls.RadioButton OptPersona 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Personal"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptFecha 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   735
         _Version        =   786432
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Fecha"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   360
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   43286.606400463
      End
      Begin TrueOleDBList80.TDBCombo cboCodigo 
         Bindings        =   "FrmAsistencias.frx":01AB
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   960
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
         _PropDict       =   $"FrmAsistencias.frx":01C5
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
      Begin XtremeSuiteControls.DateTimePicker dtpFechaFin 
         Height          =   375
         Left            =   5280
         TabIndex        =   10
         Top             =   840
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   43286.606400463
      End
      Begin XtremeSuiteControls.RadioButton OptDepartamento 
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   480
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Departamento"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptTipoNomina 
         Height          =   375
         Left            =   3840
         TabIndex        =   18
         Top             =   480
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Tipo Nomina"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblNombre 
         Caption         =   " "
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   7095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Asistencia"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc adoEmpleado 
      Height          =   375
      Left            =   960
      Top             =   8520
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
      Left            =   5880
      Top             =   8880
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
      Connect         =   ""
      OLEDBString     =   ""
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
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   375
      Left            =   6000
      Top             =   8520
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
   Begin MSAdodcLib.Adodc AdoDepartamentos 
      Height          =   375
      Left            =   960
      Top             =   8280
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
      Left            =   6480
      Top             =   8160
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
   Begin VB.Label LblTotalHoras 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   7200
      Width           =   8175
   End
End
Attribute VB_Name = "FrmAsistencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboCodigo_ItemChange()
  Me.lblNombre.Caption = Me.cboCodigo.Columns(1).Text
  CmdConsultar_Click
End Sub

Private Sub CmdAgregar_Click()
FrmAsistenciaAgregarRegistros.dtpFechEntrada.Value = Me.dtpFecha.Value
If Me.OptFecha.Value = True Then
    FrmAsistenciaAgregarRegistros.ChkTodos.Value = 1
    FrmAsistenciaAgregarRegistros.cboCodigo.Text = ""
    FrmAsistenciaAgregarRegistros.cboCodigo.Enabled = False
    FrmAsistenciaAgregarRegistros.cboCodigo.Visible = True
   
ElseIf Me.OptDepartamento.Value = True Then
    FrmAsistenciaAgregarRegistros.ChkTodos.Value = 1
    FrmAsistenciaAgregarRegistros.cboCodigo.Text = ""
    FrmAsistenciaAgregarRegistros.cboCodigo.Enabled = False
    FrmAsistenciaAgregarRegistros.cboCodigo.Visible = False
    FrmAsistenciaAgregarRegistros.cboDepartamento.Top = 840
    FrmAsistenciaAgregarRegistros.cboDepartamento.Left = 360
    FrmAsistenciaAgregarRegistros.cboDepartamento.Visible = True
    FrmAsistenciaAgregarRegistros.cboDepartamento.Text = Me.cboDepartamento.Text
ElseIf Me.OptTipoNomina.Value = True Then
    FrmAsistenciaAgregarRegistros.ChkTodos.Value = 1
    FrmAsistenciaAgregarRegistros.cboCodigo.Text = ""
    FrmAsistenciaAgregarRegistros.cboCodigo.Enabled = False
    FrmAsistenciaAgregarRegistros.cboCodigo.Visible = False
    FrmAsistenciaAgregarRegistros.cboDepartamento.Top = 840
    FrmAsistenciaAgregarRegistros.cboDepartamento.Left = 360
    FrmAsistenciaAgregarRegistros.cboDepartamento.Visible = False
    FrmAsistenciaAgregarRegistros.cboDepartamento.Text = Me.cboDepartamento.Text
    

Else
    FrmAsistenciaAgregarRegistros.ChkTodos.Value = 0
    FrmAsistenciaAgregarRegistros.cboCodigo.Text = Me.cboCodigo.Text
    FrmAsistenciaAgregarRegistros.cboCodigo.Enabled = True
End If

FrmAsistenciaAgregarRegistros.Show 1
End Sub

Private Sub CmdConsultar_Click()

Dim sFechaEntrada As String


    res = Bitacora(Now, NombreUsuario, "Asistencia", "Se Consulto la Asistencia: ")
    
    
   Me.lblNombre.Caption = ""
      If Trim(Me.cboCodigo.Text) <> "" Then
        sCodEmpl = Me.cboCodigo.Text
      End If
      
      dFecha = Me.dtpFecha.Value
      sFechaEntrada = Mid$(dFecha, 7, 4) & "-" & Mid$(dFecha, 4, 2) & "-" & Mid$(dFecha, 1, 2)

      If Me.OptPersona.Value = True Then
      
        
        Me.adoAsistencia.CommandType = adCmdText
'        Me.adoAsistencia.RecordSource = "SELECT CodEmpleado, CodEmpleado1, FechaEntrada, FechaSalida, HoraEntrada, HoraSalida, bActivo FROM AsistenciaEmpleado WHERE [FechaEntrada] <= CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102) AND [CodEmpleado1] ='" & sCodEmpl & "'"
        Me.adoAsistencia.RecordSource = "SELECT  AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.CodEmpleado1, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraEntrada , AsistenciaEmpleado.HoraSalida, (CAST(DateDiff(n, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.HoraSalida)AS FLOAT) - HorarioEmpleado.TComida) / 60 AS HorasLab FROM  AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN HorarioEmpleado ON Empleado.CodEmpleado1 = HorarioEmpleado.CodEmpleado  " & _
                                        "WHERE (AsistenciaEmpleado.FechaEntrada BETWEEN CONVERT(DATETIME, '" & Format(Me.dtpFecha.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DtpFechaFin.Value, "yyyy-mm-dd") & "', 102)) AND (AsistenciaEmpleado.CodEmpleado1 = '" & sCodEmpl & "')"
        Me.adoAsistencia.Refresh
        
        If Not Me.adoAsistencia.Recordset.EOF Then
            Me.AdoConsulta.RecordSource = "SELECT SUM(ROUND((CAST(DATEDIFF(n, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.HoraSalida) AS FLOAT) - HorarioEmpleado.TComida) / 60,2)) AS HorasLab FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN HorarioEmpleado ON Empleado.CodEmpleado1 = HorarioEmpleado.CodEmpleado  " & _
                                         "WHERE (AsistenciaEmpleado.FechaEntrada BETWEEN CONVERT(DATETIME, '" & Format(Me.dtpFecha.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DtpFechaFin.Value, "yyyy-mm-dd") & "', 102)) AND (AsistenciaEmpleado.CodEmpleado1 = '" & sCodEmpl & "')"
            Me.AdoConsulta.Refresh
            If Not Me.AdoConsulta.Recordset.EOF Then
               Me.LblTotalHoras.Caption = "Total de " & Format(Me.AdoConsulta.Recordset("HorasLab"), "##,##0.00") & " Horas"
            End If
        End If
        
        
      ElseIf Me.OptFecha.Value = True Then
        Me.adoAsistencia.CommandType = adCmdText
'        Me.adoAsistencia.RecordSource = "SELECT CodEmpleado, CodEmpleado1, FechaEntrada, FechaSalida, HoraEntrada, HoraSalida, bActivo FROM AsistenciaEmpleado WHERE [FechaEntrada] = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102) "
        Me.adoAsistencia.RecordSource = "SELECT  AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.CodEmpleado1, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraEntrada , AsistenciaEmpleado.HoraSalida, (CAST(DateDiff(n, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.HoraSalida)AS FLOAT) - HorarioEmpleado.TComida) / 60 AS HorasLab FROM  AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN HorarioEmpleado ON Empleado.CodEmpleado1 = HorarioEmpleado.CodEmpleado  " & _
                                     "WHERE AsistenciaEmpleado.FechaEntrada = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102)"
        Me.adoAsistencia.Refresh
        
        If Not Me.adoAsistencia.Recordset.EOF Then
            Me.AdoConsulta.RecordSource = "SELECT SUM(ROUND((CAST(DATEDIFF(n, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.HoraSalida) AS FLOAT) - HorarioEmpleado.TComida) / 60, 2)) AS HorasLab FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN HorarioEmpleado ON Empleado.CodEmpleado1 = HorarioEmpleado.CodEmpleado  " & _
                                         "WHERE AsistenciaEmpleado.FechaEntrada = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102)"
            Me.AdoConsulta.Refresh
            If Not Me.AdoConsulta.Recordset.EOF Then
               Me.LblTotalHoras.Caption = "Total de " & Format(Me.AdoConsulta.Recordset("HorasLab"), "##,##0.00") & " Horas"
            End If
        End If
        
        
      ElseIf Me.OptDepartamento.Value = True Then
        
        
        Me.adoAsistencia.CommandType = adCmdText
        Me.adoAsistencia.RecordSource = "SELECT AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.CodEmpleado1, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.HoraSalida, (CAST(DATEDIFF(n, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.HoraSalida) AS FLOAT) - HorarioEmpleado.TComida) / 60 AS HorasLab, Departamento.Departamento FROM  AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN HorarioEmpleado ON Empleado.CodEmpleado1 = HorarioEmpleado.CodEmpleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento  " & _
                                        "WHERE (Empleado.CodDepartamento = '" & Me.cboDepartamento.Columns(0).Text & "') AND (AsistenciaEmpleado.FechaEntrada = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102)) ORDER BY AsistenciaEmpleado.CodEmpleado1"
        Me.adoAsistencia.Refresh
        
        If Not Me.adoAsistencia.Recordset.EOF Then
            Me.AdoConsulta.RecordSource = "SELECT SUM(ROUND((CAST(DATEDIFF(n, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.HoraSalida) AS FLOAT) - HorarioEmpleado.TComida) / 60,2)) AS HorasLab FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN HorarioEmpleado ON Empleado.CodEmpleado1 = HorarioEmpleado.CodEmpleado  " & _
                                         "WHERE AsistenciaEmpleado.FechaEntrada = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102)"
            Me.AdoConsulta.Refresh
            If Not Me.AdoConsulta.Recordset.EOF Then
               Me.LblTotalHoras.Caption = "Total de " & Format(Me.AdoConsulta.Recordset("HorasLab"), "##,##0.00") & " Horas"
            End If
        End If
        
      ElseIf Me.OptTipoNomina.Value = True Then
        
        
        Me.adoAsistencia.CommandType = adCmdText
        Me.adoAsistencia.RecordSource = "SELECT AsistenciaEmpleado.CodEmpleado, AsistenciaEmpleado.CodEmpleado1, AsistenciaEmpleado.FechaEntrada, AsistenciaEmpleado.FechaSalida, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.HoraSalida, (CAST(DATEDIFF(n, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.HoraSalida) AS FLOAT) - HorarioEmpleado.TComida) / 60 AS HorasLab, Departamento.Departamento FROM  AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN HorarioEmpleado ON Empleado.CodEmpleado1 = HorarioEmpleado.CodEmpleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento  " & _
                                        "WHERE (Empleado.CodTipoNomina = '" & Me.cboTipoNomina.Columns(0).Text & "') AND (AsistenciaEmpleado.FechaEntrada = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102)) ORDER BY AsistenciaEmpleado.CodEmpleado1"
        Me.adoAsistencia.Refresh
        
        If Not Me.adoAsistencia.Recordset.EOF Then
            Me.AdoConsulta.RecordSource = "SELECT SUM(ROUND((CAST(DATEDIFF(n, AsistenciaEmpleado.HoraEntrada, AsistenciaEmpleado.HoraSalida) AS FLOAT) - HorarioEmpleado.TComida) / 60,2)) AS HorasLab FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado INNER JOIN HorarioEmpleado ON Empleado.CodEmpleado1 = HorarioEmpleado.CodEmpleado  " & _
                                         "WHERE AsistenciaEmpleado.FechaEntrada = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102)"
            Me.AdoConsulta.Refresh
            If Not Me.AdoConsulta.Recordset.EOF Then
               Me.LblTotalHoras.Caption = "Total de " & Format(Me.AdoConsulta.Recordset("HorasLab"), "##,##0.00") & " Horas"
            End If
        End If
      End If
      
      
      Me.TDBGrid.Columns(6).NumberFormat = "##,##0.00"


End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
    Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
    'Call Formato_Excel(8, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
V = 2
H = 0
i = 1

 '///////////////////////////////////////////////////////////////////////////////////////
 '////////////////////ENCABEZADOS//////////////////////////////////////////////////////
 '///////////////////////////////////////////////////////////////////////////////////

            objExcel.ActiveSheet.Cells(1, 1) = "CodEmpleado"
            objExcel.ActiveSheet.Cells(1, 2) = "FechaEntrada"
            objExcel.ActiveSheet.Cells(1, 3) = "FechaSalida"
            objExcel.ActiveSheet.Cells(1, 4) = "HoraEntrada"
            objExcel.ActiveSheet.Cells(1, 5) = "HoraSalida"
            objExcel.ActiveSheet.Cells(1, 6) = "HorasLab"
            objExcel.ActiveSheet.Cells(1, 7) = "Departamento"

  Do While Not Me.adoAsistencia.Recordset.EOF   'esto nos sirve pa leer los datos desde
       
       
            objExcel.ActiveSheet.Cells(V, H + 1) = Me.adoAsistencia.Recordset("CodEmpleado")
            objExcel.ActiveSheet.Cells(V, H + 2) = Me.adoAsistencia.Recordset("FechaEntrada")
            objExcel.ActiveSheet.Cells(V, H + 3) = Me.adoAsistencia.Recordset("FechaSalida")
            objExcel.ActiveSheet.Cells(V, H + 4) = Me.adoAsistencia.Recordset("HoraEntrada")
            objExcel.ActiveSheet.Cells(V, H + 5) = Me.adoAsistencia.Recordset("HoraSalida")
            objExcel.ActiveSheet.Cells(V, H + 6) = Me.adoAsistencia.Recordset("HoraEntrada")
            objExcel.ActiveSheet.Cells(V, H + 7) = Me.adoAsistencia.Recordset("HoraSalida")
            objExcel.ActiveSheet.Cells(V, H + 8) = Me.adoAsistencia.Recordset("HorasLab")
            objExcel.ActiveSheet.Cells(V, H + 10) = Me.adoAsistencia.Recordset("Departamento")

            
            V = V + 1
            i = i + 1
           
           
           Me.adoAsistencia.Recordset.MoveNext

 

  Loop
  
        objExcel.ActiveSheet.Columns("A").ColumnWidth = 10
        objExcel.ActiveSheet.Columns("B").ColumnWidth = 40
        objExcel.ActiveSheet.Columns("C").ColumnWidth = 18
        objExcel.ActiveSheet.Columns("D").ColumnWidth = 10
        objExcel.ActiveSheet.Columns("E").ColumnWidth = 30
        objExcel.ActiveSheet.Columns("F").ColumnWidth = 30
        objExcel.ActiveSheet.Columns("G").ColumnWidth = 17
        objExcel.ActiveSheet.Columns("G").NumberFormat = "##,##0.00"
        objExcel.ActiveSheet.Columns("H").ColumnWidth = 10
        objExcel.ActiveSheet.Columns("I").ColumnWidth = 10
        objExcel.ActiveSheet.Columns("J").ColumnWidth = 10
        objExcel.ActiveSheet.Columns("J").NumberFormat = "##,##0.00"
        objExcel.ActiveSheet.Columns("K").ColumnWidth = 10
        objExcel.ActiveSheet.Columns("L").ColumnWidth = 17
        objExcel.ActiveSheet.Columns("L").NumberFormat = "dd/mm/yyyy"
        objExcel.ActiveSheet.Columns("M").ColumnWidth = 15
        objExcel.ActiveSheet.Columns("M").NumberFormat = "dd/mm/yyyy"
        


 
 Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
End Sub

Private Sub Form_Load()

Me.dtpFecha.Value = Format(Now, "dd/MM/yyyy")
Me.DtpFechaFin.Value = Format(Now, "dd/MM/yyyy")

 Me.TDBGrid.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.TDBGrid.OddRowStyle.BackColor = &H80000005
 Me.TDBGrid.AlternatingRowStyle = True
 
Me.BackColor = RGB(216, 228, 248)
Me.Frame1.BackColor = RGB(216, 228, 248)
Me.LblTotalHoras.BackColor = RGB(216, 228, 248)
Me.Label1.BackColor = RGB(216, 228, 248)
Me.Label2.BackColor = RGB(216, 228, 248)
Me.OptDepartamento.BackColor = RGB(216, 228, 248)
Me.OptFecha.BackColor = RGB(216, 228, 248)
Me.OptPersona.BackColor = RGB(216, 228, 248)
Me.lblNombre.BackColor = RGB(216, 228, 248)
Me.TabControl1.Color = RGB(216, 228, 248)
Me.CmdConsultar.BackColor = RGB(216, 228, 248)
Me.CmdAgregar.BackColor = RGB(216, 228, 248)
Me.CmdSalir.BackColor = RGB(216, 228, 248)
 
With Me.adoEmpleado
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT CodEmpleado1, Nombre1 + ' '+ Nombre2 +' '+Apellido1+' '+Apellido2 as Nombres From Empleado Where (Activo = 1) ORDER BY CodEmpleado1"
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

With Me.adoAsistencia
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
'   .RecordSource = "AsistenciaEmpleado"
'   .Refresh
End With

With Me.AdoConsulta
   .ConnectionString = Conexion
End With

 CmdConsultar_Click
End Sub

Private Sub OptDepartamento_Click()
If Me.OptDepartamento.Value = True Then
     Me.Label2.Visible = False
     Me.cboCodigo.Visible = False
     Me.lblNombre.Visible = False
     Me.DtpFechaFin.Visible = False
     Me.cboTipoNomina.Visible = False
     Me.cboDepartamento.Visible = True
     Me.cboDepartamento.Left = 960
     Me.cboDepartamento.Top = 960
End If
End Sub

Private Sub OptFecha_Click()
  If Me.OptFecha.Value = True Then
     Me.Label2.Visible = False
     Me.cboCodigo.Visible = False
     Me.lblNombre.Visible = False
     Me.DtpFechaFin.Visible = False
     Me.cboDepartamento.Visible = False
     Me.cboTipoNomina.Visible = False
  Else
      Me.Label2.Visible = True
      Me.cboCodigo.Visible = True
     Me.lblNombre.Visible = True
     Me.DtpFechaFin.Visible = True
  End If
End Sub

Private Sub OptPersona_Click()
  If Me.OptFecha.Value = True Then
     Me.Label2.Visible = False
     Me.cboCodigo.Visible = False
     Me.lblNombre.Visible = False
     Me.DtpFechaFin.Visible = False
     Me.cboDepartamento.Visible = False
     Me.cboTipoNomina.Visible = False
  Else
      Me.Label2.Visible = True
      Me.cboCodigo.Visible = True
     Me.lblNombre.Visible = True
     Me.DtpFechaFin.Visible = True
     Me.cboDepartamento.Visible = False
     Me.cboTipoNomina.Visible = False
  End If

End Sub

Private Sub PushButton1_Click()
Unload Me
End Sub


Private Sub OptTipoNomina_Click()
If Me.OptTipoNomina.Value = True Then
     Me.Label2.Visible = False
     Me.cboCodigo.Visible = False
     Me.lblNombre.Visible = False
     Me.DtpFechaFin.Visible = False
     Me.cboDepartamento.Visible = False
     Me.cboTipoNomina.Visible = True
     Me.cboTipoNomina.Left = 960
     Me.cboTipoNomina.Left = 960

End If
End Sub
