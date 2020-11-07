VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmSalarioHistorial 
   Caption         =   "Historial Salarial"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14640
   LinkTopic       =   "Form2"
   ScaleHeight     =   7380
   ScaleWidth      =   14640
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoSalarioVacaciones 
      Height          =   375
      Left            =   3720
      Top             =   8280
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
      Caption         =   "AdoSalarioVacaciones"
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
   Begin MSAdodcLib.Adodc AdoSalarios 
      Height          =   375
      Left            =   840
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
      Caption         =   "AdoSalarios"
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
      Left            =   720
      Top             =   8520
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
   Begin MSDataListLib.DataCombo DBCNominas 
      Bindings        =   "FrmSalarioHistorial.frx":0000
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      ListField       =   "Nomina"
      Text            =   "Listado de Nominas"
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12632319
      TabCaption(0)   =   "Vacaciones"
      TabPicture(0)   =   "FrmSalarioHistorial.frx":001C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Image1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TxtFINIVaca"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TxtFFinVaca"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DbgrVacaciones"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TxtDiasDescuento"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TxtNumNomVaca"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Trecavo Mes"
      TabPicture(1)   =   "FrmSalarioHistorial.frx":0038
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtNumNom13"
      Tab(1).Control(1)=   "Dbgr13Mes"
      Tab(1).Control(2)=   "TxtFFIN13"
      Tab(1).Control(3)=   "TxtFINI13"
      Tab(1).Control(4)=   "Image2"
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(6)=   "Label5"
      Tab(1).Control(7)=   "Label2"
      Tab(1).ControlCount=   8
      Begin VB.TextBox TxtNumNomVaca 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox TxtNumNom13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   -74520
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox TxtDiasDescuento 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Text            =   "0"
         Top             =   3000
         Visible         =   0   'False
         Width           =   375
      End
      Begin TrueOleDBGrid70.TDBGrid Dbgr13Mes 
         Bindings        =   "FrmSalarioHistorial.frx":0054
         Height          =   5415
         Left            =   -72960
         TabIndex        =   2
         Top             =   480
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   9551
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
      Begin TrueOleDBGrid70.TDBGrid DbgrVacaciones 
         Bindings        =   "FrmSalarioHistorial.frx":006E
         Height          =   5295
         Left            =   2160
         TabIndex        =   3
         Top             =   600
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   9340
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
      Begin MSComCtl2.DTPicker TxtFFIN13 
         Height          =   315
         Left            =   -74760
         TabIndex        =   4
         Top             =   2280
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   77529089
         CurrentDate     =   38305
      End
      Begin MSComCtl2.DTPicker TxtFINI13 
         Height          =   315
         Left            =   -74760
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   77529089
         CurrentDate     =   38305
      End
      Begin MSComCtl2.DTPicker TxtFFinVaca 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   77529089
         CurrentDate     =   38305
      End
      Begin MSComCtl2.DTPicker TxtFINIVaca 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   77529089
         CurrentDate     =   38305
      End
      Begin VB.Image Image2 
         Height          =   2040
         Left            =   -74880
         Picture         =   "FrmSalarioHistorial.frx":0091
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   1845
      End
      Begin VB.Image Image1 
         Height          =   1920
         Left            =   120
         Picture         =   "FrmSalarioHistorial.frx":30D3
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   1845
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Fecha Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Fecha Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Fecha Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   1920
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Fecha Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   12
         Top             =   2760
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Días de Descuento"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.PushButton CmdPrnNomina 
      Height          =   375
      Left            =   11520
      TabIndex        =   19
      Top             =   6840
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Imprmir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmSalarioHistorial.frx":BB8B
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   13080
      TabIndex        =   20
      Top             =   6840
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmSalarioHistorial.frx":DE77
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton ButtonExcelAguinaldo 
      Height          =   375
      Left            =   9960
      TabIndex        =   21
      Top             =   6840
      Visible         =   0   'False
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Excel"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmSalarioHistorial.frx":E37B
      ImageAlignment  =   0
   End
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   375
      Left            =   6960
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
   Begin XtremeSuiteControls.PushButton ButtonImprimirAguinaldo 
      Height          =   375
      Left            =   11520
      TabIndex        =   22
      Top             =   6840
      Visible         =   0   'False
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Imprmir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmSalarioHistorial.frx":10667
      ImageAlignment  =   0
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Listado de Nominas"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FrmSalarioHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NumeroNomina As Double
Public TipoNomina As String

Private Sub ButtonExcelAguinaldo_Click()
Dim sql As String


    Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
'    Call Formato_Excel(8, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
V = 6
H = 0
i = 1

'///////////////////////////////////////////////////////////////////////////////////////
'////////////////////ENCABEZADOS//////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////
           objExcel.ActiveSheet.Cells(1, 4) = Titulo
           objExcel.ActiveSheet.Cells(2, 4) = "REPORTE  DE AGUINALDO"
'           objExcel.ActiveSheet.Cells(3, 4) = "Impreso desde: " & Me.TxtFecha1.Value & " Hasta: " & Me.TxtFecha2.Value
           objExcel.ActiveSheet.Cells(4, 4) = Format(Now, "Long Date")
           objExcel.ActiveSheet.Columns("D").HorizontalAlignment = 3
            
            
            objExcel.ActiveSheet.Cells(5, 1) = "Codigo Empleado"
            objExcel.ActiveSheet.Cells(5, 2) = "Nombre Empleado"
            objExcel.ActiveSheet.Cells(5, 3) = "Fecha Corte"
            objExcel.ActiveSheet.Cells(5, 4) = "Junio"
            objExcel.ActiveSheet.Cells(5, 5) = "Julio"
            objExcel.ActiveSheet.Cells(5, 6) = "Agosto"
            objExcel.ActiveSheet.Cells(5, 7) = "Septiembre"
            objExcel.ActiveSheet.Cells(5, 8) = "Octubre"
            objExcel.ActiveSheet.Cells(5, 9) = "Noviembre"
            

NumeroNomina = Me.TxtNumNom13.Text
sql = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres," & vbLf
sql = sql & "HistorialSalarioMes.Fechaini , HistorialSalarioMes.Fechafin, HistorialSalarioMes.Enero, HistorialSalarioMes.Febrero, HistorialSalarioMes.Marzo, " & vbLf
sql = sql & "HistorialSalarioMes.Abril , HistorialSalarioMes.Mayo, HistorialSalarioMes.Junio, HistorialSalarioMes.Julio, HistorialSalarioMes.Agosto, " & vbLf
sql = sql & "HistorialSalarioMes.Septiembre , HistorialSalarioMes.Octubre, HistorialSalarioMes.Noviembre, HistorialSalarioMes.Diciembre, " & vbLf
sql = sql & "HistorialSalarioMes.NumNomina " & vbLf
sql = sql & "FROM HistorialSalarioMes INNER JOIN" & vbLf
sql = sql & "Empleado ON HistorialSalarioMes.CodEmpleado = Empleado.CodEmpleado" & vbLf
sql = sql & "Where (HistorialSalarioMes.NumNomina = " & NumeroNomina & ") AND (HistorialSalarioMes.Tipo = 'Aguinaldo') ORDER BY Empleado.CodEmpleado1"
            
            
            
       Me.AdoConsulta.RecordSource = sql
       Me.AdoConsulta.Refresh
       Do While Not Me.AdoConsulta.Recordset.EOF
         
                objExcel.ActiveSheet.Cells(V, H + 1) = Me.AdoConsulta.Recordset("CodEmpleado1")
                objExcel.ActiveSheet.Cells(V, H + 2) = Me.AdoConsulta.Recordset("Nombres")
                objExcel.ActiveSheet.Cells(V, H + 3) = Format(Me.AdoConsulta.Recordset("FechaFin"), "dd/mm/yyyy")
                objExcel.ActiveSheet.Cells(V, H + 4) = Format(Me.AdoConsulta.Recordset("Junio"), "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 5) = Format(Me.AdoConsulta.Recordset("Julio"), "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 6) = Format(Me.AdoConsulta.Recordset("Agosto"), "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 7) = Format(Me.AdoConsulta.Recordset("Septiembre"), "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 8) = Format(Me.AdoConsulta.Recordset("Octubre"), "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 9) = Format(Me.AdoConsulta.Recordset("Noviembre"), "##,##0.00")
      
         V = V + 1
         Me.AdoConsulta.Recordset.MoveNext
       Loop
            
            
   objExcel.ActiveSheet.Columns("A").ColumnWidth = 13.3
   objExcel.ActiveSheet.Columns("B").ColumnWidth = 50

End Sub

Private Sub ButtonImprimirAguinaldo_Click()
Dim rpt As New ArepHistorialSalarial13vo
Dim fPreview As New FrmPreview



NumeroNomina = Me.TxtNumNom13.Text
sql = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres," & vbLf
sql = sql & "HistorialSalarioMes.Fechaini , HistorialSalarioMes.Fechafin, HistorialSalarioMes.Enero, HistorialSalarioMes.Febrero, HistorialSalarioMes.Marzo, " & vbLf
sql = sql & "HistorialSalarioMes.Abril , HistorialSalarioMes.Mayo, HistorialSalarioMes.Junio, HistorialSalarioMes.Julio, HistorialSalarioMes.Agosto, " & vbLf
sql = sql & "HistorialSalarioMes.Septiembre , HistorialSalarioMes.Octubre, HistorialSalarioMes.Noviembre, HistorialSalarioMes.Diciembre, " & vbLf
sql = sql & "HistorialSalarioMes.NumNomina " & vbLf
sql = sql & "FROM HistorialSalarioMes INNER JOIN" & vbLf
sql = sql & "Empleado ON HistorialSalarioMes.CodEmpleado = Empleado.CodEmpleado" & vbLf
sql = sql & "Where (HistorialSalarioMes.NumNomina = " & NumeroNomina & ") AND (HistorialSalarioMes.Tipo = 'Aguinaldo')  ORDER BY Nombres"

      'ArepHistorial.DataControl1.Source = SQlReportes
'      rpt.lbltiponom.Text = TipoNomina

             
             rpt.AdoHistorial.ConnectionString = Conexion
             rpt.AdoHistorial.Source = sql
             fPreview.RunReport rpt
        
        
             fPreview.Show 1
End Sub

Private Sub CmdPrnNomina_Click()
Dim rpt As New ArepHistorialSalarial13vo, Tipo As String
Dim fPreview As New FrmPreview



NumeroNomina = Me.TxtNumNom13.Text
sql = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres," & vbLf
sql = sql & "HistorialSalarioMes.Fechaini , HistorialSalarioMes.Fechafin, HistorialSalarioMes.Enero, HistorialSalarioMes.Febrero, HistorialSalarioMes.Marzo, " & vbLf
sql = sql & "HistorialSalarioMes.Abril , HistorialSalarioMes.Mayo, HistorialSalarioMes.Junio, HistorialSalarioMes.Julio, HistorialSalarioMes.Agosto, " & vbLf
sql = sql & "HistorialSalarioMes.Septiembre , HistorialSalarioMes.Octubre, HistorialSalarioMes.Noviembre, HistorialSalarioMes.Diciembre, " & vbLf
sql = sql & "HistorialSalarioMes.NumNomina " & vbLf
sql = sql & "FROM HistorialSalarioMes INNER JOIN" & vbLf
sql = sql & "Empleado ON HistorialSalarioMes.CodEmpleado = Empleado.CodEmpleado" & vbLf
sql = sql & "Where (HistorialSalarioMes.NumNomina = " & NumeroNomina & ") AND (HistorialSalarioMes.Tipo = '" & TipoNomina & "')  ORDER BY Nombres"

      'ArepHistorial.DataControl1.Source = SQlReportes
'      rpt.lbltiponom.Text = TipoNomina

            
          Quien = "Vacaciones"
             
             rpt.AdoHistorial.ConnectionString = Conexion
             rpt.AdoHistorial.Source = sql
             fPreview.RunReport rpt
        
        
             fPreview.Show 1

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
 Me.Dbgr13Mes.EvenRowStyle.BackColor = RGB(175, 189, 133)
 Me.Dbgr13Mes.OddRowStyle.BackColor = &H80000005
 Me.Dbgr13Mes.AlternatingRowStyle = True
 

   
 Me.DbgrVacaciones.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DbgrVacaciones.OddRowStyle.BackColor = &H80000005
 Me.DbgrVacaciones.AlternatingRowStyle = True

With Me.AdoTipoNomina
 '.DatabaseName = Ruta
 .ConnectionString = Conexion
End With

With Me.AdoSalarioVacaciones
 '.DatabaseName = Ruta
 .ConnectionString = Conexion
End With

With Me.AdoSalarios
 .ConnectionString = Conexion
' .RecordSource = ""
' .Refresh
End With

With Me.AdoConsulta
 .ConnectionString = Conexion
' .RecordSource = ""
' .Refresh
End With

End Sub

Private Sub PushButton1_Click()

End Sub
