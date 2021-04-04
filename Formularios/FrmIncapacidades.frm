VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmIncapacidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Incapacidades"
   ClientHeight    =   5955
   ClientLeft      =   15
   ClientTop       =   405
   ClientWidth     =   6360
   HelpContextID   =   31
   Icon            =   "FrmIncapacidades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "FrmIncapacidades.frx":030A
   ScaleHeight     =   397
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   8
      Top             =   4680
      Width           =   3015
      Begin XtremeSuiteControls.PushButton CmdAnterior 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Anterior"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmIncapacidades.frx":074C
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdSiguiente 
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Siguiente"
         ForeColor       =   0
         TextAlignment   =   0
         Appearance      =   6
         Picture         =   "FrmIncapacidades.frx":0C4E
         ImageAlignment  =   1
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton CmdAgregar 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Agregar"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmIncapacidades.frx":1152
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdBorrar 
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   720
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Borrar"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmIncapacidades.frx":34B6
         ImageAlignment  =   0
      End
   End
   Begin VB.CommandButton CmdAgrega 
      Caption         =   "X"
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
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin TrueOleDBGrid70.TDBGrid DBGridTipoIncapacidad 
      Bindings        =   "FrmIncapacidades.frx":396A
      Height          =   3135
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5530
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
   Begin TrueOleDBGrid70.TDBGrid DBGridTipo 
      Bindings        =   "FrmIncapacidades.frx":398B
      Height          =   2175
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3836
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
   Begin VB.TextBox TxtEmpleado 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   4215
   End
   Begin MSAdodcLib.Adodc DtaBusca 
      Height          =   375
      Left            =   360
      Top             =   7800
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
   Begin MSAdodcLib.Adodc DtaTipoIncapacidad 
      Height          =   375
      Left            =   240
      Top             =   7200
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
      Caption         =   "DtaTipoIncapacidad"
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
   Begin MSAdodcLib.Adodc DtaIncapacidad 
      Height          =   375
      Left            =   120
      Top             =   7200
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
      Caption         =   "DtaIncapacidad"
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
      Left            =   3720
      Top             =   7440
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
   Begin MSAdodcLib.Adodc DtaEmpleado 
      Height          =   375
      Left            =   3480
      Top             =   7200
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
   Begin MSDataListLib.DataCombo DBCCodempleado 
      Bindings        =   "FrmIncapacidades.frx":39A5
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodEmpleado"
      Text            =   ""
   End
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   5400
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmIncapacidades.frx":39BF
      ImageAlignment  =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   16
      X2              =   16
      Y1              =   264
      Y2              =   96
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C00000&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   416
      X2              =   16
      Y1              =   264
      Y2              =   264
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C00000&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   416
      X2              =   416
      Y1              =   264
      Y2              =   96
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   16
      X2              =   416
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero Empleado:"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tabla de Incapacidades de los empleados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Empleado:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "FrmIncapacidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub VScroll1_Change()

End Sub

Private Sub CmdAgrega_Click()
DbCCodEmpleado.Enabled = True
 CmdSiguiente.Enabled = True
  CmdAnterior.Enabled = True
 cmdborrar.Enabled = True
 CmdAgrega.Visible = False
 Me.DBGridTipoIncapacidad.Visible = False

End Sub

Private Sub cmdAgregar_Click()
 If txtEmpleado.Text = "" Then
  MsgBox "No Puede agregar en este Registro", vbCritical, "Error:Sistema de Nominas"
  Exit Sub
 End If
  DtaTipoIncapacidad.Refresh
  DtaIncapacidad.Refresh

  DbCCodEmpleado.Enabled = False
  CmdSiguiente.Enabled = False
  CmdAnterior.Enabled = False
  cmdborrar.Enabled = False
  CmdAgrega.Visible = True
  Me.DBGridTipoIncapacidad.Visible = True
   Me.DBGridTipoIncapacidad.Columns(1).Width = 3500
End Sub

Private Sub CmdAnterior_Click()
DtaEmpleado.Recordset.MovePrevious
       If DtaEmpleado.Recordset.BOF Then
           DtaEmpleado.Recordset.MoveNext
           MsgBox "Imposible ir al registro especificado.Esta al Inicio de un conjunto de registros", vbInformation, "Sistema de Nominas"
       Else
           DbCCodEmpleado.Text = DtaEmpleado.Recordset("CodEmpleado")
       End If
End Sub

Private Sub cmdborrar_Click()
 Dim Respuesta, Rsp
If txtEmpleado.Text = "" Then
 MsgBox "No Puede Eliminar este Registro", vbCritical, "Error:Sistema de Nominas"
 Exit Sub
End If

'Elimino el registro activo en la pantalla
  DtaIncapacidad.Refresh
  Do While Not DtaIncapacidad.Recordset.EOF
    If DtaIncapacidad.Recordset("ID") = DtaConsulta.Recordset("ID") Then
        Set Rsp = DtaIncapacidad.Recordset
        Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando una Incapacidad de: " & txtEmpleado.Text)
         If Respuesta = 6 Then
          Rsp.Delete
          Rsp.MovePrevious
          SQlConsulta = "SELECT Incapacidad.FechaIncapacidad,Incapacidad.Id,Incapacidad.CodIncapacidad, TipoIncapacidad.Incapacidad FROM TipoIncapacidad INNER JOIN (Empleado INNER JOIN Incapacidad ON Empleado.CodEmpleado = Incapacidad.CodEmpleado) ON TipoIncapacidad.CodIncapacidad = Incapacidad.CodIncapacidad WHERE (((Empleado.CodEmpleado)='" & DbCCodEmpleado.Text & "'))"
          DtaConsulta.RecordSource = SQlConsulta
          DtaConsulta.Refresh
          DBGridTipoIncapacidad.Columns(2).Visible = False
          DBGridTipoIncapacidad.Columns(1).Visible = False
          Exit Sub
         End If
    End If
  DtaIncapacidad.Recordset.MoveNext
  Loop
End Sub

Private Sub cmdGrabar_Click()

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSiguiente_Click()
DtaEmpleado.Recordset.MoveNext
       If DtaEmpleado.Recordset.EOF Then
           DtaEmpleado.Recordset.MovePrevious
           MsgBox "Imposible ir al registro especificado.Esta al Final de un conjunto de registros", vbInformation, "Sistema de Nominas"
       Else
           DbCCodEmpleado.Text = DtaEmpleado.Recordset("CodEmpleado")
       End If
End Sub

Private Sub DBCCodEmpleado_Change()
 Dim SQlConsulta As String
Evaluar = True
 'Al ejecutar algun cambio en el combo actualizo el nombre del departamento

'        TxtEmpleado.Text = DtaEmpleado.Recordset("Nombre1")
        SQlConsulta = "SELECT Incapacidad.FechaIncapacidad, Incapacidad.CodIncapacidad, TipoIncapacidad.Incapacidad, Empleado.CodEmpleado,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres FROM TipoIncapacidad INNER JOIN Empleado INNER JOIN Incapacidad ON Empleado.CodEmpleado = Incapacidad.CodEmpleado ON TipoIncapacidad.CodIncapacidad = Incapacidad.CodIncapacidad WHERE (Empleado.CodEmpleado = '" & DbCCodEmpleado.Text & "')"
'        SqlConsulta = "SELECT Incapacidad.FechaIncapacidad,Incapacidad.Id,Incapacidad.CodIncapacidad, TipoIncapacidad.Incapacidad FROM TipoIncapacidad INNER JOIN (Empleado INNER JOIN Incapacidad ON Empleado.CodEmpleado = Incapacidad.CodEmpleado) ON TipoIncapacidad.CodIncapacidad = Incapacidad.CodIncapacidad WHERE (((Empleado.CodEmpleado)='" & DBCCodempleado.Text & "'))"
        DtaConsulta.RecordSource = SQlConsulta
        DtaConsulta.Refresh
        If Not DtaConsulta.Recordset.EOF Then
          Me.txtEmpleado.Text = Me.DtaConsulta.Recordset("Nombres")
          Me.DBGridTipo.Columns(1).Visible = False
          Me.DBGridTipo.Columns(2).Width = 3500
          Me.DBGridTipo.Columns(3).Visible = False
          Me.DBGridTipo.Columns(4).Visible = False
        Else
        SQlConsulta = "SELECT CodEmpleado, Nombre1 + ' ' + Nombre2 + ' ' + Apellido1 + ' ' + Apellido2 AS Nombres From Empleado WHERE     (CodEmpleado = '" & DbCCodEmpleado.Text & "')"
        DtaBusca.RecordSource = SQlConsulta
        DtaBusca.Refresh
         If Not DtaBusca.Recordset.EOF Then
          Me.txtEmpleado.Text = Me.DtaBusca.Recordset("Nombres")
          Me.DBGridTipo.Columns(1).Visible = False
          Me.DBGridTipo.Columns(2).Width = 3500
          Me.DBGridTipo.Columns(3).Visible = False
          Me.DBGridTipo.Columns(4).Visible = False
         End If
        End If
                                          


  
                                                

'       DtaEmpleado.Recordset.MoveNext
'   Loop

  
End Sub

Private Sub DBGridTipo_KeyPress(KeyAscii As Integer)

 If KeyAscii = 42 Then
   DBGridTipo.Visible = True
 Else
   Evaluar = False
  End If

End Sub

Private Sub DBGridTipoIncapacidad_DblClick()
Dim SQlConsulta As String, Codigo As Variant
DtaIncapacidad.Recordset.AddNew
DtaIncapacidad.Recordset("FechaIncapacidad") = "01/01/2000"
DtaIncapacidad.Recordset("CodIncapacidad") = DtaTipoIncapacidad.Recordset("CodIncapacidad")
DtaIncapacidad.Recordset("CodEmpleado") = DbCCodEmpleado.Text
DtaIncapacidad.Recordset.Update
        SQlConsulta = "SELECT Incapacidad.FechaIncapacidad, Incapacidad.CodIncapacidad, TipoIncapacidad.Incapacidad, Empleado.CodEmpleado,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres FROM TipoIncapacidad INNER JOIN Empleado INNER JOIN Incapacidad ON Empleado.CodEmpleado = Incapacidad.CodEmpleado ON TipoIncapacidad.CodIncapacidad = Incapacidad.CodIncapacidad WHERE (Empleado.CodEmpleado = '" & DbCCodEmpleado.Text & "')"
'        SqlConsulta = "SELECT Incapacidad.FechaIncapacidad,Incapacidad.Id,Incapacidad.CodIncapacidad, TipoIncapacidad.Incapacidad FROM TipoIncapacidad INNER JOIN (Empleado INNER JOIN Incapacidad ON Empleado.CodEmpleado = Incapacidad.CodEmpleado) ON TipoIncapacidad.CodIncapacidad = Incapacidad.CodIncapacidad WHERE (((Empleado.CodEmpleado)='" & DBCCodempleado.Text & "'))"
        DtaConsulta.RecordSource = SQlConsulta
        DtaConsulta.Refresh
        If Not DtaConsulta.Recordset.EOF Then
          Me.txtEmpleado.Text = Me.DtaConsulta.Recordset("Nombres")
          Me.DBGridTipo.Columns(1).Visible = False
          Me.DBGridTipo.Columns(2).Width = 3500
          Me.DBGridTipo.Columns(3).Visible = False
          Me.DBGridTipo.Columns(4).Visible = False
        End If
DbCCodEmpleado.Enabled = True
CmdSiguiente.Enabled = True
  CmdAnterior.Enabled = True
Me.DBGridTipoIncapacidad.Visible = False
CmdAgrega.Visible = False
cmdborrar.Enabled = True
End Sub

Private Sub DBGridTipoIncapacidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim SQlConsulta As String, Codigo As Variant
DtaIncapacidad.Recordset.AddNew
DtaIncapacidad.Recordset("FechaIncapacidad") = "01/01/2000"
DtaIncapacidad.Recordset("CodIncapacidad") = DtaTipoIncapacidad.Recordset("CodIncapacidad")
DtaIncapacidad.Recordset("CodEmpleado") = DbCCodEmpleado.Text
DtaIncapacidad.Recordset.Update
SQlConsulta = "SELECT Incapacidad.FechaIncapacidad,Incapacidad.Id,Incapacidad.CodIncapacidad, TipoIncapacidad.Incapacidad FROM TipoIncapacidad INNER JOIN (Empleado INNER JOIN Incapacidad ON Empleado.CodEmpleado = Incapacidad.CodEmpleado) ON TipoIncapacidad.CodIncapacidad = Incapacidad.CodIncapacidad WHERE (((Empleado.CodEmpleado)='" & DbCCodEmpleado.Text & "'))"
        DtaConsulta.RecordSource = SQlConsulta
        DtaConsulta.Refresh
        DBGridTipoIncapacidad.Columns(2).Visible = False
        DBGridTipoIncapacidad.Columns(1).Visible = False
DbCCodEmpleado.Enabled = True
CmdSiguiente.Enabled = True
  CmdAnterior.Enabled = True
DBGridTipo.Visible = False
CmdAgrega.Visible = False
cmdborrar.Enabled = True
End If
End Sub


Private Sub Form_Activate()
  DBGridTipoIncapacidad.Columns(0).Caption = "Fecha Incapacidad"
  DBGridTipoIncapacidad.Columns(1).Caption = "Descripcion Incapacidad"
' If Not BIncapacidad = True Then
'   CmdBorrar.Enabled = False
' End If
' If Not GIncapacidad = True Then
'   CmdAgregar.Enabled = False
' End If
End Sub

Private Sub Form_Load()

Me.BackColor = RGB(222, 227, 247)
Me.Frame1.BackColor = RGB(222, 227, 247)


With Me.DtaConsulta
   .ConnectionString = Conexion
End With

With Me.DtaBusca

   .ConnectionString = Conexion
End With

With Me.DtaEmpleado
   .ConnectionString = Conexion
   .RecordSource = "Empleado"
   .Refresh
End With

With Me.DtaIncapacidad
   .ConnectionString = Conexion
   .RecordSource = "Incapacidad"
End With

With Me.DtaTipoIncapacidad
   
   .ConnectionString = Conexion
   .RecordSource = "TipoIncapacidad"
   .Refresh
End With


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If GIncapacidad = True Then
 
End If
Exit Sub
TipoErrs:
 ControlErrores

End Sub

Private Sub TDBGrid1_Click()

End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub
