VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmListNomina 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Nóminas Cerradas"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   9360
      TabIndex        =   33
      Top             =   6120
      Width           =   135
   End
   Begin VB.CommandButton FrmReportes 
      Height          =   375
      Left            =   7560
      Picture         =   "FrmListNomina.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4920
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc DtaDetalleNominas 
      Height          =   375
      Left            =   3600
      Top             =   8520
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
   Begin MSAdodcLib.Adodc DtaMovPrestamo 
      Height          =   375
      Left            =   3960
      Top             =   8280
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
      Caption         =   "DtaMovPrestamo"
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
   Begin MSAdodcLib.Adodc DtaPrestamo 
      Height          =   375
      Left            =   720
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
      Caption         =   "DtaPrestamo"
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
      Left            =   720
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
      Height          =   6015
      Left            =   80
      ScaleHeight     =   5955
      ScaleWidth      =   9315
      TabIndex        =   0
      Top             =   80
      Width           =   9375
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   360
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin TrueOleDBGrid70.TDBGrid DbgrdetalleNominas 
         Bindings        =   "FrmListNomina.frx":16E6
         Height          =   3255
         Left            =   120
         TabIndex        =   27
         Top             =   2640
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5741
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      Begin TrueOleDBGrid70.TDBGrid DbgrNominas 
         Bindings        =   "FrmListNomina.frx":1706
         Height          =   2415
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4260
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
      Begin VB.TextBox TxtIR 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         DataField       =   "TotalMontoIR"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox TxtInss 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         DataField       =   "TotalMontoINSS"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox TxtPrestamo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         DataField       =   "TotalPrestamo"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox TxtDeducciones 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         DataField       =   "TotalDeducciones"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox TxtIncentivos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         DataField       =   "TotalIncentivos"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TxtHE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         DataField       =   "TotalHorasExtras"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TxtComisiones 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         DataField       =   "TotalComisiones"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TxtDestajo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         DataField       =   "TotalDestajo"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtOtIngre 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         DataField       =   "TotalOtrosIngresos"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TxtSalbasico 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         DataField       =   "TotalSalarioBasico"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   3255
         Left            =   7320
         TabIndex        =   1
         Top             =   2760
         Width           =   1695
         Begin VB.CommandButton CmdExportaColilla 
            DownPicture     =   "FrmListNomina.frx":171F
            Height          =   375
            Left            =   120
            Picture         =   "FrmListNomina.frx":3201
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton CmdExportar 
            DownPicture     =   "FrmListNomina.frx":4C23
            Height          =   375
            Left            =   120
            Picture         =   "FrmListNomina.frx":6705
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton CmdExportaCsv 
            DownPicture     =   "FrmListNomina.frx":8007
            Height          =   375
            Left            =   120
            Picture         =   "FrmListNomina.frx":9AE9
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CommandButton CmdSalir 
            DownPicture     =   "FrmListNomina.frx":B10B
            Height          =   375
            Left            =   120
            Picture         =   "FrmListNomina.frx":CBED
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   2400
            Width           =   1455
         End
         Begin VB.CommandButton CmdAnularNomina 
            DownPicture     =   "FrmListNomina.frx":E6CF
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            Picture         =   "FrmListNomina.frx":101B1
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   3000
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton CmdPrNomina 
            DownPicture     =   "FrmListNomina.frx":11C93
            Height          =   375
            Left            =   120
            Picture         =   "FrmListNomina.frx":13775
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton CmdPrColilla 
            DownPicture     =   "FrmListNomina.frx":15257
            Height          =   375
            Left            =   120
            Picture         =   "FrmListNomina.frx":16D39
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Salario Básico"
         Height          =   255
         Left            =   6120
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Destajo"
         Height          =   255
         Left            =   6120
         TabIndex        =   24
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Horas Extras"
         Height          =   255
         Left            =   6120
         TabIndex        =   23
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "INSS"
         Height          =   255
         Left            =   6120
         TabIndex        =   22
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "IR"
         Height          =   255
         Left            =   6120
         TabIndex        =   21
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Comisiones"
         Height          =   255
         Left            =   6120
         TabIndex        =   20
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Incentivos"
         Height          =   255
         Left            =   6120
         TabIndex        =   19
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Deducciones"
         Height          =   255
         Left            =   6120
         TabIndex        =   18
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Prestamos"
         Height          =   255
         Left            =   6120
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Otros Ingresos"
         Height          =   255
         Left            =   6120
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   8160
      Top             =   6720
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
      ChangeControlsBackColor=   0   'False
   End
   Begin MSAdodcLib.Adodc AdoBusca 
      Height          =   375
      Left            =   1200
      Top             =   9000
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
   Begin XtremeSuiteControls.ProgressBar Barra 
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   6240
      Width           =   7575
      _Version        =   786432
      _ExtentX        =   13361
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
End
Attribute VB_Name = "FrmListNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnularNomina_Click()


k% = MsgBox("Realmente desea anular esta Nómina? ", vbYesNo)
If k <> 6 Then Exit Sub

NumNomina = DtaNominas.Recordset("NumNomina")

DtaDetalleNominas.Refresh
Do While Not DtaDetalleNominas.Recordset.EOF
    If DtaDetalleNominas.Recordset("NumNomina") = NumNomina Then
        DtaDetalleNominas.Recordset.Delete
    End If
DtaDetalleNominas.Recordset.MoveNext
Loop

'DtaNominas.Recordset.Edit

'DtaNominas.Recordset.Edit
DtaNominas.Recordset("TotalSalarioBasico") = 0
DtaNominas.Recordset("TotalDestajo") = 0
DtaNominas.Recordset("TotalHorasExtras") = 0
DtaNominas.Recordset("TotalComisiones") = 0
DtaNominas.Recordset("TotalIncentivos") = 0
DtaNominas.Recordset("TotalDeducciones") = 0
DtaNominas.Recordset("TotalPrestamo") = 0
DtaNominas.Recordset("TotalMontoInss") = 0
DtaNominas.Recordset("TotalMontoIR") = 0
DtaNominas.Recordset("TotalVacaciones") = 0
DtaNominas.Recordset("TotalINSSPatronal") = 0
DtaNominas.Recordset("TotalIRPatronal") = 0
DtaNominas.Recordset("Totalmes13") = 0
DtaNominas.Recordset("Procesada") = True
DtaNominas.Recordset("anulada") = True
DtaNominas.Recordset.Update


Dim rs As New ADODB.Recordset
   
'Set dbs = OpenDatabase("c:\Sistema de Nominas\Nominas.mdb")
   
   rs.Open "UPDATE DetalleDeduccion SET DetalleDeduccion.Pagado = False WHERE DetalleDeduccion.NumNomina= " & NumNomina & "", Conexion
   rs.Open "UPDATE DetalleIncentivo SET DetalleIncentivo.Pagado = False WHERE DetalleIncentivo.NumNomina= " & NumNomina & "", Conexion
   rs.Open "UPDATE MovPrestamo SET MovPrestamo.Cancelado = False WHERE MovPrestamo.NumNomina= " & NumNomina & "", Conexion
   rs.Open "UPDATE Comisiones SET Comisiones.Pagado = False WHERE Comisiones.NumNomina= " & NumNomina & "", Conexion
   rs.Open "UPDATE Destajo SET Destajo.Pagado = False WHERE Destajo.NumNomina= " & NumNomina & "", Conexion
   rs.Open "UPDATE HorasExtras SET HorasExtras.Pagada = False WHERE HorasExtras.NumNomina= " & NumNomina & "", Conexion
   'dbs.Close


'regreso el saldo del prestamo a la normalidad

SQlPrestamo = "SELECT Prestamo.* From Prestamo"
Dtaprestamo.RecordSource = SQlPrestamo
Dtaprestamo.Refresh

DtaMovprestamo.Refresh
Do While Not DtaMovprestamo.Recordset.EOF
    If DtaMovprestamo.Recordset("NumNomina") = NumNomina Then
     Dtaprestamo.Refresh
     Do While Not Dtaprestamo.Recordset.EOF
      'MsgBox ((Str(DtaPrestamo.Recordset("NumPrestamo")) + " Movprestamo ") + Str(DtaMovPrestamo.Recordset("NumPrestamo")))
        If Dtaprestamo.Recordset("NumPrestamo") = DtaMovprestamo.Recordset("NumPrestamo") Then
            ''DtaPrestamo.Recordset.Edit
            Dtaprestamo.Recordset("Saldo") = Dtaprestamo.Recordset("Saldo") + DtaMovprestamo.Recordset("CuotaIgual")
            Dtaprestamo.Recordset.Update
        End If
    Dtaprestamo.Recordset.MoveNext
    Loop
  End If
  DtaMovprestamo.Recordset.MoveNext
  
Loop




MsgBox "La Nomina Ha sido Anulada, Los Movimientos han sido Revertidos"

End Sub

Private Sub CmdExportaColilla_Click()
Me.CommonDialog1.ShowSave
Directorio = ""
Directorio = Me.CommonDialog1.FileName + ".xls"

FechaIni = Format(DtaNominas.Recordset("FechaNominaINI"), "dddddd")
FechaFin = Format(DtaNominas.Recordset("FechaNomina"), "dddddd")
NumNomina = DtaNominas.Recordset("NumNomina")

''//////////////////INTRUCCION SQL SERVER
'SQLReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
'SQLReportes = SQLReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
'SQLReportes = SQLReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
'SQLReportes = SQLReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
'SQLReportes = SQLReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
'SQLReportes = SQLReportes & "                      Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
'SQLReportes = SQLReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
'SQLReportes = SQLReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
'SQLReportes = SQLReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
'SQLReportes = SQLReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
'SQLReportes = SQLReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
'SQLReportes = SQLReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
'SQLReportes = SQLReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion AS TotalDevengado," & vbLf
'SQLReportes = SQLReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
'SQLReportes = SQLReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
'SQLReportes = SQLReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion)" & vbLf
'SQLReportes = SQLReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
'SQLReportes = SQLReportes & "                      Empleado.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion,Empleado.CodEmpleado1" & vbLf
'SQLReportes = SQLReportes & " FROM         Nomina INNER JOIN" & vbLf
'SQLReportes = SQLReportes & "                       Grupo INNER JOIN" & vbLf
'SQLReportes = SQLReportes & "                       Cargo INNER JOIN" & vbLf
'SQLReportes = SQLReportes & "                       TipoNomina INNER JOIN" & vbLf
'SQLReportes = SQLReportes & "                       Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
'SQLReportes = SQLReportes & "                       DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
'SQLReportes = SQLReportes & "                       TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina" & vbLf
'SQLReportes = SQLReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ")" & vbLf
'SQLReportes = SQLReportes & " ORDER BY Nomina.NumNomina, DetalleNomina.CodEmpleado" & vbLf


'///////////////////////////INTRUCCION SQL SERVER
SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
SQlReportes = SQlReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
SQlReportes = SQlReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
SQlReportes = SQlReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
SQlReportes = SQlReportes & "                      Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.BonoProduccion, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
SQlReportes = SQlReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad AS TotalDevengado," & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion+ DetalleNomina.Antiguedad)" & vbLf
SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion,Empleado.CodEmpleado1,DetalleNomina.Antiguedad,DetalleNomina.AñoAntiguedad" & vbLf
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
SQlReportes = SQlReportes & " ORDER BY Nomina.NumNomina, Empleado.CodEmpleado1" & vbLf




'ArepNomina.ImgLogo.Picture = LoadPicture(RutaLogo)



MDIPrimero.DtaEmpresa.Refresh
If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
  FormatoColilla = MDIPrimero.DtaEmpresa.Recordset("FormatoColilla")
End If

Select Case FormatoColilla
  Case "Colilla Comercial"
    ArepColillasPago.AdoColillas.Source = SQlReportes
    ArepColillasPago.LblPeriodo.Caption = FechaIni & " al " & FechaFin
    ArepColillasPago.LblTitulo.Caption = Titulo
    ArepColillasPago.AdoColillas.ConnectionString = ConexionReporte
    ArepColillasPago.Show 1
     
  
  Case "Colilla Produccion"

    ArepColillas.AdoColillas.Source = SQlReportes
    ArepColillas.LblPeriodo.Caption = FechaIni & " al " & FechaFin
    ArepColillas.LblTitulo.Caption = Titulo
    ArepColillas.AdoColillas.ConnectionString = ConexionReporte
    ArepColillas.Show 1

End Select
Exportar = False

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
NumNomina = Me.DtaNominas.Recordset("NumNomina")

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
                   
                    Cadena = "CodEmpleado" & "," & "Nombres" & "," & "SalarioBasico" & "," & "Destajo" & "," & "HorasExtras" & "," & "Comisiones" & "," & "Incentivos" & "," & "VacacionesPagadas" & "," & "SeptimoDia" & "," & "IncetivoProduccion" & "," & "OtrosIngresos" & "," & "TotalDevengado" & "," & "Prestamo" & "," & "MontoINSS" & "," & "MontoIr" & "," & "Deducciones" & "," & "INSSPatronal"
                    Print #1, Cadena
                   
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
Me.CommonDialog1.ShowSave
Directorio = ""
Directorio = Me.CommonDialog1.FileName + ".xls"

Me.AdoBusca.RecordSource = "SELECT CodTipoNomina, Nomina, Periodo, UltFecha, TipoPago, Moneda, MantValor, Activa, PorcientoInss, TasaInss, PorcientoIr, TasaIr,TasaInssPatronal From TipoNomina WHERE  (CodTipoNomina = N'01')"
Me.AdoBusca.Refresh
If Not Me.AdoBusca.Recordset.EOF Then
  Moneda = Me.AdoBusca.Recordset("Moneda")
End If

Exportar = True
FechaIni = Format(DtaNominas.Recordset("FechaNominaINI"), "dddddd")
FechaFin = Format(DtaNominas.Recordset("FechaNomina"), "dddddd")
NumNomina = DtaNominas.Recordset("NumNomina")


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
'SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad AS TotalDevengado," & vbLf
'SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
'SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
'SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad)" & vbLf
'SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
'SQlReportes = SQlReportes & "                      DetalleNomina.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion,Empleado.CodEmpleado1,Empleado.NumeroInss,DetalleNomina.Antiguedad,DetalleNomina.AñoAntiguedad" & vbLf
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
'SQlReportes = SQlReportes & " ORDER BY Empleado.CodGrupo, Empleado.CodEmpleado1" & vbLf


             SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones,"
             SQlReportes = SQlReportes & "         Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones,"
             SQlReportes = SQlReportes & "         Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado,"
             SQlReportes = SQlReportes & "         Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo, Cargo.Cargo, DetalleNomina.BonoProduccion, DetalleNomina.SalarioBasico,"
             SQlReportes = SQlReportes & "         DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones,"
             SQlReportes = SQlReportes & "        DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones,"
             SQlReportes = SQlReportes & "         DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo,"
             SQlReportes = SQlReportes & "         Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE,"
             SQlReportes = SQlReportes & "         DetalleNomina.SalarioBasico, DetalleNomina.SalarioBasico, DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas"
             SQlReportes = SQlReportes & "          + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad + DetalleNomina.HorasTurno + DetalleNomina.SalarioBasico AS TotalDevengado,"
             SQlReportes = SQlReportes & "         DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir,"
             SQlReportes = SQlReportes & "         (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas"
             SQlReportes = SQlReportes & "          + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad + DetalleNomina.HorasTurno)"
             SQlReportes = SQlReportes & "         - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar, DetalleNomina.TarifaHoraria, DetalleNomina.HTrabajada,"
             SQlReportes = SQlReportes & "         DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1, Empleado.NumeroInss, DetalleNomina.AjusteINSS, Empleado.NumCedula, DetalleNomina.Antiguedad,"
             SQlReportes = SQlReportes & "         DetalleNomina.AñoAntiguedad, Empleado.SueldoPeriodo * 2 AS SalarioMensualHM, Empleado.SueldoPeriodo * 2 / 30 AS SalarioDiaHM, 14 AS D, DetalleNomina.DiasVacaciones AS DV,"
             SQlReportes = SQlReportes & "         Empleado.DiasBasico AS DD, 14 - DetalleNomina.DiasVacaciones - Empleado.DiasBasico AS DL,"
             SQlReportes = SQlReportes & "         14 - DetalleNomina.DiasVacaciones - Empleado.DiasBasico + DetalleNomina.DiasVacaciones + DetalleNomina.DiasAdicionales AS T,"
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
             SQlReportes = SQlReportes & " ORDER BY  Departamento.Departamento, Empleado.CodGrupo, Empleado.Apellido1, Empleado.Apellido2" & vbLf  'Empleado.CodEmpleado1


If Moneda = "US" Then
 ArepNominaDolares.AdoNomina.Source = SQlReportes
 ArepNominaDolares.LblTitulo.Caption = Titulo
 ArepNominaDolares.LblSubtitulo.Caption = SubTitulo
 ArepNominaDolares.ImgLogo.Picture = LoadPicture(RutaLogo)
 ArepNominaDolares.AdoNomina.ConnectionString = ConexionReporte
 ArepNominaDolares.LblFecha.Caption = Format(Now, "dddddd")
 ArepNominaDolares.LblDesde = FechaIni
 ArepNominaDolares.LblHasta = FechaFin
 ArepNominaDolares.Show 1

Else
 MDIPrimero.DtaEmpresa.Refresh
 If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
   FormatoNomina = MDIPrimero.DtaEmpresa.Recordset("FormatoNomina")
 End If
 
 Select Case FormatoNomina
  
  Case "Nomina Comercial"
    ArepNominaComercial.AdoNomina.Source = SQlReportes
    ArepNominaComercial.LblTitulo.Caption = Titulo
    ArepNominaComercial.LblSubtitulo.Caption = SubTitulo
    ArepNominaComercial.ImgLogo.Picture = LoadPicture(RutaLogo)
    ArepNominaComercial.AdoNomina.ConnectionString = ConexionReporte
    ArepNominaComercial.LblFecha.Caption = Format(Now, "dddddd")
    ArepNominaComercial.LblDesde = FechaIni
    ArepNominaComercial.LblHasta = FechaFin
    ArepNominaComercial.Show 1
   Case "Nomina Produccion"
    ArepNomina.AdoNomina.Source = SQlReportes
    ArepNomina.LblTitulo.Caption = Titulo
    ArepNomina.LblSubtitulo.Caption = SubTitulo
    ArepNomina.ImgLogo.Picture = LoadPicture(RutaLogo)
    ArepNomina.AdoNomina.ConnectionString = ConexionReporte
    ArepNomina.LblFecha.Caption = Format(Now, "dddddd")
    ArepNomina.LblDesde = FechaIni
    ArepNomina.LblHasta = FechaFin
    ArepNomina.Show 1
    
   Case "Nomina Bono Produccion"
    ArepNominaBono.AdoNomina.Source = SQlReportes
    ArepNominaBono.LblTitulo.Caption = Titulo
    ArepNominaBono.LblSubtitulo.Caption = SubTitulo
    ArepNominaBono.ImgLogo.Picture = LoadPicture(RutaLogo)
    ArepNominaBono.AdoNomina.ConnectionString = ConexionReporte
    ArepNominaBono.LblFecha.Caption = Format(Now, "dddddd")
    ArepNominaBono.LblDesde = FechaIni
    ArepNominaBono.LblHasta = FechaFin
    ArepNominaBono.Show 1
   
  End Select
End If

Exportar = False


'FrmExporta.Show
'FrmExporta.OptTransaciones.Value = True
'NumNomina = DtaNominas.Recordset("NumNomina")
'FrmExporta.LblTitulo.Caption = "Exportando nómina: " & Str(NumNomina)
'MsgBox ("Se exportará la Nómina: " & NumNomina)
End Sub

Private Sub CmdPrColilla_Click()
Dim rpt As Object
Dim fPreview As New FrmPreview
'mjmm
FechaIni = Format(DtaNominas.Recordset("FechaNominaINI"), "dddddd")
FechaFin = Format(DtaNominas.Recordset("FechaNomina"), "dddddd")
NumNomina = DtaNominas.Recordset("NumNomina")
Espacio = " "
Quien = "ListadoNominas"
FechaInicio = FechaIni
FechaFinal = FechaFin


''///////////////////////////INTRUCCION SQL SERVER
'SQLReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
'SQLReportes = SQLReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
'SQLReportes = SQLReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
'SQLReportes = SQLReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
'SQLReportes = SQLReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
'SQLReportes = SQLReportes & "                      Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
'SQLReportes = SQLReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
'SQLReportes = SQLReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
'SQLReportes = SQLReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
'SQLReportes = SQLReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
'SQLReportes = SQLReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
'SQLReportes = SQLReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
'SQLReportes = SQLReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion AS TotalDevengado," & vbLf
'SQLReportes = SQLReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
'SQLReportes = SQLReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
'SQLReportes = SQLReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion)" & vbLf
'SQLReportes = SQLReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
'SQLReportes = SQLReportes & "                      DetalleNomina.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion,Empleado.CodEmpleado1" & vbLf
'SQLReportes = SQLReportes & " FROM         Nomina INNER JOIN" & vbLf
'SQLReportes = SQLReportes & "                       Grupo INNER JOIN" & vbLf
'SQLReportes = SQLReportes & "                       Cargo INNER JOIN" & vbLf
'SQLReportes = SQLReportes & "                       TipoNomina INNER JOIN" & vbLf
'SQLReportes = SQLReportes & "                       Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
'SQLReportes = SQLReportes & "                       DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
'SQLReportes = SQLReportes & "                       TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina" & vbLf
'SQLReportes = SQLReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ") AND((dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
'SQLReportes = SQLReportes & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia)" & vbLf
'SQLReportes = SQLReportes & "                      - (dbo.DetalleNomina.Prestamo + dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.MontoIR + dbo.DetalleNomina.Deducciones) <> 0)" & vbLf
'SQLReportes = SQLReportes & " ORDER BY Nomina.NumNomina, Empleado.CodEmpleado1" & vbLf


'///////////////////////////INTRUCCION SQL SERVER
SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
SQlReportes = SQlReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
SQlReportes = SQlReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
SQlReportes = SQlReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
SQlReportes = SQlReportes & "                      Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.BonoProduccion, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
SQlReportes = SQlReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad AS TotalDevengado," & vbLf
SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad)" & vbLf
SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
SQlReportes = SQlReportes & "                      DetalleNomina.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion,Empleado.CodEmpleado1,DetalleNomina.Antiguedad,DetalleNomina.AñoAntiguedad" & vbLf
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
SQlReportes = SQlReportes & " ORDER BY Nomina.NumNomina, Empleado.Nombre1, Empleado.CodEmpleado1" & vbLf

MDIPrimero.DtaEmpresa.Refresh
If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
  FormatoColilla = MDIPrimero.DtaEmpresa.Recordset("FormatoColilla")
End If

Select Case FormatoColilla
    Case "Colilla Trides"
    ArepColillasPago4.AdoColillas.Source = SQlReportes
    ArepColillasPago4.LblPeriodo.Caption = "   Colilla de Pago # " & NumNomina & ", Corespondiente del " & DtaNominas.Recordset("FechaNominaINI") & " al " & DtaNominas.Recordset("FechaNomina")
    ArepColillasPago4.LblTitulo.Caption = Titulo
    ArepColillasPago4.AdoColillas.ConnectionString = ConexionReporte
    ArepColillasPago4.Show 1

  Case "Colilla Comercial2"
    ArepColillasPago2.AdoColillas.Source = SQlReportes
    ArepColillasPago2.LblPeriodo.Caption = FechaIni & " al " & FechaFin
    ArepColillasPago2.LblTitulo.Caption = Titulo
    ArepColillasPago2.AdoColillas.ConnectionString = ConexionReporte
'    ArepColillasPago2.Show 1
           fPreview.arv.ReportSource = ArepColillasPago2
           fPreview.Show 1

   Case "Colilla Bono Produccion"

    ArepColillasBono.AdoColillas.Source = SQlReportes
    ArepColillasBono.LblPeriodo.Caption = FechaIni & " al " & FechaFin
    ArepColillasBono.LblTitulo.Caption = Titulo
    ArepColillasBono.AdoColillas.ConnectionString = ConexionReporte
'    ArepColillasBono.Show 1
           fPreview.arv.ReportSource = ArepColillasBono
           fPreview.Show 1
    
  Case "Colilla Comercial"
    ArepColillasPago.AdoColillas.Source = SQlReportes
    ArepColillasPago.LblPeriodo.Caption = FechaIni & " al " & FechaFin
    ArepColillasPago.LblTitulo.Caption = Titulo
    ArepColillasPago.AdoColillas.ConnectionString = ConexionReporte
    ArepColillasPago.Show 1
'           fPreview.arv.ReportSource = ArepColillasPago
'           fPreview.Show 1
     
  
  Case "Colilla Produccion"

    ArepColillas.AdoColillas.Source = SQlReportes
    ArepColillas.LblPeriodo.Caption = FechaIni & " al " & FechaFin
    ArepColillas.LblTitulo.Caption = Titulo
    ArepColillas.AdoColillas.ConnectionString = ConexionReporte
'    ArepColillas.Show 1
           fPreview.arv.ReportSource = ArepColillas
           fPreview.Show 1

End Select


End Sub

Private Sub CmdprNomina_Click()
On Error GoTo TipoErrs
Dim rpt As Object
Dim fPreview As New FrmPreview
Dim rs As New ADODB.Recordset

Me.AdoBusca.RecordSource = "SELECT CodTipoNomina, Nomina, Periodo, UltFecha, TipoPago, Moneda, MantValor, Activa, PorcientoInss, TasaInss, PorcientoIr, TasaIr,TasaInssPatronal From TipoNomina WHERE  (CodTipoNomina = N'01')"
Me.AdoBusca.Refresh
If Not Me.AdoBusca.Recordset.EOF Then
  Moneda = Me.AdoBusca.Recordset("Moneda")
End If

NumNomina = DtaNominas.Recordset("NumNomina")
Quien = "ListadoNomina"

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
'SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad AS TotalDevengado," & vbLf
'SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
'SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
'SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad)" & vbLf
'SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
'SQlReportes = SQlReportes & "                      DetalleNomina.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion,Empleado.CodEmpleado1, Empleado.NumeroInss,Empleado.NumCedula,DetalleNomina.AjusteINSS,DetalleNomina.Antiguedad,DetalleNomina.AñoAntiguedad " & vbLf
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
'SQlReportes = SQlReportes & " ORDER BY Empleado.CodGrupo, Empleado.Nombre1, Empleado.CodEmpleado1" & vbLf

             SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones,"
             SQlReportes = SQlReportes & "         Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones,"
             SQlReportes = SQlReportes & "         Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado,"
             SQlReportes = SQlReportes & "         Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo, Cargo.Cargo, DetalleNomina.BonoProduccion, DetalleNomina.SalarioBasico,"
             SQlReportes = SQlReportes & "         DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones,"
             SQlReportes = SQlReportes & "        DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones,"
             SQlReportes = SQlReportes & "         DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo,"
             SQlReportes = SQlReportes & "         Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE,"
             SQlReportes = SQlReportes & "         DetalleNomina.SalarioBasico, DetalleNomina.SalarioBasico, DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas"
             SQlReportes = SQlReportes & "          + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad + DetalleNomina.HorasTurno + DetalleNomina.SalarioBasico AS TotalDevengado,"
             SQlReportes = SQlReportes & "         DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir,"
             SQlReportes = SQlReportes & "         (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas"
             SQlReportes = SQlReportes & "          + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad + DetalleNomina.HorasTurno)"
             SQlReportes = SQlReportes & "         - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar, DetalleNomina.TarifaHoraria, DetalleNomina.HTrabajada,"
             SQlReportes = SQlReportes & "         DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1, Empleado.NumeroInss, DetalleNomina.AjusteINSS, Empleado.NumCedula, DetalleNomina.Antiguedad,"
             SQlReportes = SQlReportes & "         DetalleNomina.AñoAntiguedad, Empleado.SueldoPeriodo * 2 AS SalarioMensualHM, Empleado.SueldoPeriodo * 2 / 30 AS SalarioDiaHM, 14 AS D, DetalleNomina.DiasVacaciones AS DV,"
             SQlReportes = SQlReportes & "         Empleado.DiasBasico AS DD, 14 - DetalleNomina.DiasVacaciones - Empleado.DiasBasico AS DL,"
             SQlReportes = SQlReportes & "         14 - DetalleNomina.DiasVacaciones - Empleado.DiasBasico + DetalleNomina.DiasVacaciones + DetalleNomina.DiasAdicionales AS T,"
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
             SQlReportes = SQlReportes & " ORDER BY  Departamento.Departamento, Empleado.CodGrupo, Empleado.Apellido1, Empleado.Apellido2" & vbLf  'Empleado.CodEmpleado1


If Moneda = "US" Then
 ArepNominaDolares.AdoNomina.Source = SQlReportes
 ArepNominaDolares.LblTitulo.Caption = Titulo
 ArepNominaDolares.LblSubtitulo.Caption = SubTitulo
 ArepNominaDolares.ImgLogo.Picture = LoadPicture(RutaLogo)
 ArepNominaDolares.AdoNomina.ConnectionString = ConexionReporte
 ArepNominaDolares.LblFecha.Caption = Format(Now, "dddddd")
 FechaIni = Format(DtaNominas.Recordset("FechaNominaINI"), "dddddd")
 FechaFin = Format(DtaNominas.Recordset("FechaNomina"), "dddddd")
 ArepNominaDolares.LblDesde = FechaIni
 ArepNominaDolares.LblHasta = FechaFin
' ArepNominaDolares.Show 1
           
     Set rpt = New ArepNominaDolares
     rpt.DataControl1.ConnectionString = Conexion
     rpt.DataControl1.Source = sql
     fPreview.RunReport rpt
     fPreview.Show 1

Else
 MDIPrimero.DtaEmpresa.Refresh
 If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
   FormatoNomina = MDIPrimero.DtaEmpresa.Recordset("FormatoNomina")
 End If
 
 Select Case FormatoNomina
 
 Case "Nomina Destajo"
 
  FechaInicio = Format(DtaNominas.Recordset("FechaNominaINI"), "dddddd")
  FechaFinal = Format(DtaNominas.Recordset("FechaNomina"), "dddddd")
   Set rpt = New ArepNominaDestajo
   
    rpt.LblTitulo.Caption = Titulo
    If Dir(RutaLogo, vbDirectory) <> "" Then
        rpt.ImgLogo.Picture = LoadPicture(RutaLogo)
    End If
    
    
    rpt.NumeroNomina = NumNomina
'    rpt.LblDesde.Caption = "Desde: " & Me.LblFecha1.Caption & "   Hasta: " & Me.LblFecha2.Caption
    rpt.AdoNomina.Source = SQlReportes
      
        rpt.AdoNomina.ConnectionString = ConexionReporte
        fPreview.arv.ReportSource = rpt
        fPreview.Show 1
 
 
 Case "Nomina Hanter Metal"
    
    SQlReportes = ""
SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones,   Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones,       Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado,    Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo, Cargo.Cargo, DetalleNomina.BonoProduccion, DetalleNomina.SalarioBasico,    DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones,    DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, "
SQlReportes = SQlReportes & " DetalleNomina.MontoIR,  DetalleNomina.Vacaciones,    DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo,     Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE,       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas     + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad + DetalleNomina.HorasTurno AS TotalDevengado,     DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir,    (DetalleNomina.SalarioBasico "
SQlReportes = SQlReportes & " + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas   + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad + DetalleNomina.HorasTurno)      - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar, DetalleNomina.TarifaHoraria, DetalleNomina.HTrabajada,    DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1, Empleado.NumeroInss, DetalleNomina.AjusteINSS, Empleado.NumCedula, DetalleNomina.Antiguedad,     DetalleNomina.AñoAntiguedad, Empleado.SueldoPeriodo * 2 AS SalarioMensualHM, Empleado.SueldoPeriodo * 2 / 30 AS SalarioDiaHM, 15 AS D, DetalleNomina.DiasVacaciones AS DV,     Empleado.DiasBasico AS DD, 15 - DetalleNomina.DiasVacaciones - Empleado.DiasBasico AS DL, "
SQlReportes = SQlReportes & "    15 - DetalleNomina.DiasVacaciones - Empleado.DiasBasico + DetalleNomina.DiasVacaciones + DetalleNomina.DiasAdicionales AS T,     Empleado.Nombre1 + ' ' + Empleado.Nombre2 +' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS NombreCompleto, Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30) AS VDD,      Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30) AS DiasLab, Empleado.SueldoPeriodo * 2 / 30 / 8 * 2 AS ValorHE, Empleado.SueldoPeriodo * 2 / 30 / 8 * 2 * DetalleNomina.HE AS TotalHE, (Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30))    + Empleado.SueldoPeriodo * 2 / 30 / 8 * 2 * DetalleNomina.HE AS TotalPagar, DetalleNomina.Deducciones AS OtrasDeduciones,     (Empleado.SueldoPeriodo - Empleado.DiasBasico * (Empleado.SueldoPeriodo * 2 / 30)) * 0.02 AS Inatec,  DetalleNomina.DiasAdicionales AS DA, DetalleNomina.ValorDiasAdicionales AS VDA,"
SQlReportes = SQlReportes & "        Departamento.Departamento  FROM         Nomina INNER JOIN    Grupo INNER JOIN      Cargo INNER JOIN    TipoNomina INNER JOIN   Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN    DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND       Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN   Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento"
SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ") AND((dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
SQlReportes = SQlReportes & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia)" & vbLf
SQlReportes = SQlReportes & "                      - (dbo.DetalleNomina.Prestamo + dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.MontoIR + dbo.DetalleNomina.Deducciones) <> 0)" & vbLf
SQlReportes = SQlReportes & " ORDER BY Departamento.Departamento, Empleado.CodGrupo, Empleado.Nombre1" & vbLf  'Empleado.CodEmpleado1
 
   Set rpt = New ArepNominasHM
    rpt.LblTitulo.Caption = Titulo
    If Dir(RutaLogo) <> "" Then
        rpt.ImgLogo.Picture = LoadPicture(RutaLogo)
    End If
 
    'rpt.LblDesde.Caption = "Desde: " & Me.LblFecha1.Caption & "   Hasta: " & Me.LblFecha2.Caption
    rpt.AdoNomina.Source = SQlReportes
              
         rpt.AdoNomina.ConnectionString = ConexionReporte
'        ArepNominaProduccionLegal.Show 1
           fPreview.arv.ReportSource = rpt
           fPreview.Show 1
 
 
    Case "Nomina Produccion Tamaño Legal"
'        ArepNominaProduccionLegal.AdoNomina.Source = SQlReportes
        FechaIni = Format(DtaNominas.Recordset("FechaNominaINI"), "dddddd")
        FechaFin = Format(DtaNominas.Recordset("FechaNomina"), "dddddd")
        ArepNominaProduccionLegal.LblDesde.Caption = FechaIni
        ArepNominaProduccionLegal.LblHasta.Caption = FechaFin
        ArepNominaProduccionLegal.LblFecha = Format(Now, "dddddd")
        ArepNominaProduccionLegal.LblTitulo.Caption = Titulo
        ArepNominaProduccionLegal.LblSubtitulo.Caption = SubTitulo
        If Dir(RutaLogo) <> "" Then
          ArepNominaProduccionLegal.ImgLogo.Picture = LoadPicture(RutaLogo)
        End If
'        ArepNominaProduccionLegal.AdoNomina.ConnectionString = ConexionReporte
'        ArepNominaProduccionLegal.Show 1
'           fPreview.arv.ReportSource = ArepNominaProduccionLegal
'           fPreview.Show 1
           
          Set rpt = New ArepNominaProduccionLegal
          rpt.AdoNomina.ConnectionString = ConexionReporte
          rpt.AdoNomina.Source = SQlReportes
          fPreview.RunReport rpt

          fPreview.Show 1
 
    Case "Nomina Comercial2"
    ArepNominaComercial2.AdoNomina.Source = SQlReportes
    ArepNominaComercial2.LblTitulo.Caption = Titulo
    ArepNominaComercial2.LblSubtitulo.Caption = SubTitulo
    ArepNominaComercial2.ImgLogo.Picture = LoadPicture(RutaLogo)
    ArepNominaComercial2.AdoNomina.ConnectionString = ConexionReporte
    ArepNominaComercial2.LblFecha.Caption = Format(Now, "dddddd")
    FechaIni = Format(DtaNominas.Recordset("FechaNominaINI"), "dddddd")
    FechaFin = Format(DtaNominas.Recordset("FechaNomina"), "dddddd")
    ArepNominaComercial2.LblDesde = FechaIni
    ArepNominaComercial2.LblHasta = FechaFin
'    ArepNominaComercial2.Show 1
'           fPreview.arv.ReportSource = ArepNominaComercial2
'           fPreview.Show 1
     Set rpt = New ArepNominaComercial2
     rpt.AdoNomina.ConnectionString = Conexion
     rpt.AdoNomina.Source = sql
     fPreview.RunReport rpt
     fPreview.Show 1
  
  Case "Nomina Comercial"
    ArepNominaComercial.AdoNomina.Source = SQlReportes
    ArepNominaComercial.LblTitulo.Caption = Titulo
    ArepNominaComercial.LblSubtitulo.Caption = SubTitulo
    ArepNominaComercial.ImgLogo.Picture = LoadPicture(RutaLogo)
    ArepNominaComercial.AdoNomina.ConnectionString = ConexionReporte
    ArepNominaComercial.LblFecha.Caption = Format(Now, "dddddd")
    FechaIni = Format(DtaNominas.Recordset("FechaNominaINI"), "dddddd")
    FechaFin = Format(DtaNominas.Recordset("FechaNomina"), "dddddd")
    ArepNominaComercial.LblDesde = FechaIni
    ArepNominaComercial.LblHasta = FechaFin
'    ArepNominaComercial.Show 1
'           fPreview.arv.ReportSource = ArepNominaComercial
'           fPreview.Show 1
'     Set rpt = New ArepNominaComercial
'     rpt.AdoNomina.ConnectionString = Conexion
'     rpt.AdoNomina.Source = sql
'     fPreview.RunReport rpt
'     fPreview.Show 1

 ArepNominaComercial.Show 1
    
   Case "Nomina Produccion"
    ArepNomina.AdoNomina.Source = SQlReportes
    ArepNomina.LblTitulo.Caption = Titulo
    ArepNomina.LblSubtitulo.Caption = SubTitulo
    ArepNomina.ImgLogo.Picture = LoadPicture(RutaLogo)
    ArepNomina.AdoNomina.ConnectionString = ConexionReporte
    ArepNomina.LblFecha.Caption = Format(Now, "dddddd")
    FechaIni = Format(DtaNominas.Recordset("FechaNominaINI"), "dddddd")
    FechaFin = Format(DtaNominas.Recordset("FechaNomina"), "dddddd")
    ArepNomina.LblDesde = FechaIni
    ArepNomina.LblHasta = FechaFin
    ArepNomina.NumeroNomina = NumNomina
    ArepNomina.Show 1
'           Set rpt = New ArepNomina
'           fPreview.arv.ReportSource = ArepNomina
'           fPreview.RunReport rpt
'           fPreview.Show 1
'
    Case "Nomina Bono Produccion"
    ArepNominaBono.AdoNomina.Source = SQlReportes
    ArepNominaBono.LblTitulo.Caption = Titulo
    ArepNominaBono.LblSubtitulo.Caption = SubTitulo
    ArepNominaBono.ImgLogo.Picture = LoadPicture(RutaLogo)
    ArepNominaBono.AdoNomina.ConnectionString = ConexionReporte
    ArepNominaBono.LblFecha.Caption = Format(Now, "dddddd")
    FechaIni = Format(DtaNominas.Recordset("FechaNominaINI"), "dddddd")
    FechaFin = Format(DtaNominas.Recordset("FechaNomina"), "dddddd")
    ArepNominaBono.LblDesde = FechaIni
    ArepNominaBono.LblHasta = FechaFin
    ArepNominaBono.Show 1
'           Set rpt = New ArepNominaBono
'           fPreview.r
''           fPreview.arv.ReportSource = ArepNominaBono
'           fPreview.RunReport rpt
'           fPreview.Show 1
   
  End Select
End If



'With CrtNomina
'  .ReportFileName = "C:\Documents and Settings\Juan Gabriel\Mis documentos\Reporte.rpt"
'  .SQLQuery = SQLReportes
'  .Connect = "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemaNominas"
'  .RetrieveDataFiles
'  .WindowShowPrintBtn = True
'  .Action = 1
'End With

Exit Sub
TipoErrs:
 MsgBox Err.Description

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim SalarioBasico As Double, Devengado As Double
Dim INSSPATRONAL As Double, INATEC As Double, VACACIONES As Double, AGUINALDO As Double
NumNominas = DtaNominas.Recordset("NumNomina")

AdoBusca.RecordSource = "SELECT  *, SalarioBasico + Destajo + HorasExtras + Comisiones + OtrosIngresos + Incentivos + SeptimoDia + IncetivoProduccion +  BonoProduccion AS TotalDevengado From DetalleNomina WHERE (DetalleNomina.NumNomina = '" & NumNominas & "' )"
'Me.AdoBusca.RecordSource = "SELECT   * FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina WHERE (Nomina.NumNomina = '" & NumNominas & "' )"
Me.AdoBusca.Refresh
Do While Not Me.AdoBusca.Recordset.EOF
  SalarioBasico = Me.AdoBusca.Recordset("SalarioBasico")
  Devengado = Me.AdoBusca.Recordset("TotalDevengado")
  INSSPATRONAL = Devengado * 0.16
  INATEC = SalarioBasico * 0.02
  VACACIONES = SalarioBasico / 12
  AGUINALDO = SalarioBasico / 12
  
  Me.AdoBusca.Recordset("INSSPatronal") = INSSPATRONAL
  Me.AdoBusca.Recordset("INATEC") = INATEC
  Me.AdoBusca.Recordset("Vacaciones") = VACACIONES
  Me.AdoBusca.Recordset("Mes13") = AGUINALDO
 Me.AdoBusca.Recordset.Update

 Me.AdoBusca.Recordset.MoveNext
Loop

MsgBox "pROCESO TERMINADO"


End Sub

Private Sub DbgrNominas_Click()
Dim SqlNominas As String
Dim SqlDetalleNominas As String
Dim NumNominas As Long


'SQlNominas = "SELECT Nomina.NumNomina, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal, Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, Nomina.Anulada From Nomina WHERE (((Nomina.Anulada)=False)"
'DtaNominas.RecordSource = SQlNominas
'DtaNominas.Refresh

NumNominas = DtaNominas.Recordset("NumNomina")


SqlDetalleNominas = "SELECT DetalleNomina.NumNomina, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2,DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos,DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones,DetalleNomina.INSSPatronal , DetalleNomina.Mes13 FROM Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado Where (DetalleNomina.NumNomina = " & NumNominas & ") ORDER BY Empleado.CodEmpleado1"
DtaDetalleNominas.RecordSource = SqlDetalleNominas
DtaDetalleNominas.Refresh


End Sub

Private Sub Form_Load()
Me.DbgrdetalleNominas.EvenRowStyle.BackColor = &HC0FFFF
 Me.DbgrdetalleNominas.OddRowStyle.BackColor = &HFFFFFF
 Me.DbgrdetalleNominas.AlternatingRowStyle = True
 
Me.DbgrNominas.EvenRowStyle.BackColor = &HC0FFFF
 Me.DbgrNominas.OddRowStyle.BackColor = &HFFFFFF
 Me.DbgrNominas.AlternatingRowStyle = True
 
 With Me.AdoBusca
   .ConnectionString = Conexion
 End With


With Me.DtaDetalleNominas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT DetalleNomina.NumNomina, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2,DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos,DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones,DetalleNomina.INSSPatronal , DetalleNomina.Mes13 FROM Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ORDER BY Empleado.CodEmpleado1"
   .Refresh
End With

With Me.DtaMovprestamo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "MovPrestamo"
   .Refresh
End With

With Me.DtaNominas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Nomina"
   .Refresh
End With

With Me.Dtaprestamo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Prestamo"
   .Refresh
End With

  
End Sub

Private Sub FrmReportes_Click()
On Error GoTo TipoErrs
Quien = "Listado"
CodTipoNomina = Me.DtaNominas.Recordset("CodTipoNomina")
NumNomina = DtaNominas.Recordset("NumNomina")
FrmNominaActiva.Show 1
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub TxtComisiones_Change()
TxtComisiones.Text = Format(TxtComisiones.Text, "###,##0.00")
End Sub

Private Sub TxtDeducciones_Change()
TxtDeducciones.Text = Format(TxtDeducciones.Text, "###,##0.00")
End Sub

Private Sub TxtDestajo_Change()
TxtDestajo.Text = Format(TxtDestajo.Text, "###,##0.00")
End Sub

Private Sub TxtHE_Change()
TxtHE.Text = Format(TxtHE.Text, "###,##0.00")
End Sub

Private Sub TxtIncentivos_Change()
TxtIncentivos.Text = Format(TxtIncentivos.Text, "###,##0.00")
End Sub

Private Sub TxtInss_Change()
TxtInss.Text = Format(TxtInss.Text, "###,##0.00")
End Sub

Private Sub TxtIR_Change()
TxtIR.Text = Format(TxtIR.Text, "###,##0.00")
End Sub

Private Sub TxtOtIngre_Change()
TxtOtIngre.Text = Format(TxtOtIngre.Text, "###,##0.00")
End Sub

Private Sub TxtPrestamo_Change()
TxtPrestamo.Text = Format(TxtPrestamo.Text, "###,##0.00")
End Sub

Private Sub TxtSalbasico_Change()
txtSalBasico.Text = Format(txtSalBasico.Text, "###,##0.00")
End Sub
