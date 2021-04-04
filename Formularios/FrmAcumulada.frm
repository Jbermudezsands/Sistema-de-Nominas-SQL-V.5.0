VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form FrmNominaAcumulada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FrmRegistro de Nomina Acumulada"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   14400
   Begin MSAdodcLib.Adodc AdoDetalleNominaAcumulada 
      Height          =   375
      Left            =   5160
      Top             =   8640
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
      Caption         =   "AdoDetalleNominaAcumulada"
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
   Begin VB.TextBox TxtAno 
      Height          =   285
      Left            =   360
      TabIndex        =   14
      Top             =   5280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc AdoEmpleados 
      Height          =   375
      Left            =   600
      Top             =   7560
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   14415
      TabIndex        =   2
      Top             =   0
      Width           =   14415
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRO DE NOMINAS ACUMULADAS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   4560
         TabIndex        =   3
         Top             =   360
         Width           =   4440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   14400
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image2 
         Height          =   1020
         Left            =   840
         Picture         =   "FrmAcumulada.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   12615
      Begin TrueOleDBList80.TDBCombo DBCNominas 
         Bindings        =   "FrmAcumulada.frx":4027
         Height          =   315
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
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
         ComboStyle      =   2
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
         _PropDict       =   $"FrmAcumulada.frx":4043
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
      Begin MSComCtl2.DTPicker DtaFechaINI 
         Height          =   300
         Left            =   3960
         TabIndex        =   6
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   54591489
         CurrentDate     =   40804
      End
      Begin VB.TextBox TxtNumNomina 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin XtremeSuiteControls.PushButton CmdGenerar 
         Height          =   495
         Left            =   9240
         TabIndex        =   1
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Generar"
         ForeColor       =   0
         TextAlignment   =   0
         Appearance      =   6
         Picture         =   "FrmAcumulada.frx":40ED
         ImageAlignment  =   1
         TextImageRelation=   4
      End
      Begin MSDataListLib.DataCombo DBComboPeriodo 
         Bindings        =   "FrmAcumulada.frx":702F
         Height          =   315
         Left            =   960
         TabIndex        =   10
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Periodo"
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DtFechaFin 
         Height          =   300
         Left            =   7080
         TabIndex        =   11
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Format          =   54591489
         CurrentDate     =   40804
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   495
         Left            =   10800
         TabIndex        =   17
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Agregar Empleado"
         ForeColor       =   0
         TextAlignment   =   0
         Appearance      =   6
         Picture         =   "FrmAcumulada.frx":7049
         ImageAlignment  =   1
         TextImageRelation=   4
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin"
         Height          =   255
         Left            =   6240
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Nomina"
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio"
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero Nomina"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin TrueOleDBGrid70.TDBGrid DbgrdetalleNominas 
      Bindings        =   "FrmAcumulada.frx":941B
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   6800
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
   Begin MSAdodcLib.Adodc AdoNomina 
      Height          =   375
      Left            =   600
      Top             =   8040
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
   Begin MSAdodcLib.Adodc DtaTipoNomina 
      Height          =   375
      Left            =   4440
      Top             =   8040
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
   Begin MSAdodcLib.Adodc DtaPeriodos 
      Height          =   375
      Left            =   4560
      Top             =   8160
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
      Caption         =   "DtaPeriodos"
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
      Left            =   6600
      Top             =   7920
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
   Begin XtremeSuiteControls.ProgressBar Barra 
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   6480
      Width           =   14175
      _Version        =   786432
      _ExtentX        =   25003
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin MSAdodcLib.Adodc DtaDetalleNomina 
      Height          =   375
      Left            =   600
      Top             =   8280
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
   Begin MSAdodcLib.Adodc AdoTotales 
      Height          =   375
      Left            =   10320
      Top             =   7920
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
      Caption         =   "AdoTotales"
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
Attribute VB_Name = "FrmNominaAcumulada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdUltimo_Click()

End Sub

Private Sub CmdGenerar_Click()
Dim sql As String, Año As Double, CodTipoNomina As String
Dim CodEmpleado As String, Periodo As Double, Mes As Double
Dim i As Double

If Me.DBCNominas.Text = "" Then
 Exit Sub
End If

If Me.DBComboPeriodo.Text = "" Then
 Exit Sub
End If

Año = Me.txtAno.Text
CodTipoNomina = Me.DBCNominas.Columns(0).Text
Periodo = Me.DBComboPeriodo.Text

Me.AdoConsulta.RecordSource = "SELECT  * From Fecha_Planilla WHERE (año = " & Año & ") AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Periodo & ")"
Me.AdoConsulta.Refresh
If Not Me.AdoConsulta.Recordset.EOF Then
  Mes = Me.AdoConsulta.Recordset("mes")
End If

Me.adoNomina.RecordSource = "SELECT  * From NominaAcumulada"
Me.adoNomina.Refresh

Me.DtaDetalleNomina.RecordSource = "SELECT  * From DetalleNominaAcumulada"
Me.DtaDetalleNomina.Refresh


'*******************************************************************************************************
'///////////////////////////////AGREGO LA NOMINA CERO //////////////////////////////////////////////////
'*******************************************************************************************************
adoNomina.Recordset.AddNew

adoNomina.Recordset("NumNomina") = Me.TxtNumNomina.Text
adoNomina.Recordset("CodTipoNomina") = CodTipoNomina
adoNomina.Recordset("FechaNomina") = Format(CDate(Me.DTFechaFin.Value), "DD/MM/YYYY")
adoNomina.Recordset("FechaNominaINI") = Format(CDate(Me.DtaFechaINI.Value), "DD/MM/YYYY")
adoNomina.Recordset("Activa") = 0
adoNomina.Recordset("TotalSalarioBasico") = 0
adoNomina.Recordset("TotalDestajo") = 0
adoNomina.Recordset("TotalHorasExtras") = 0
adoNomina.Recordset("TotalComisiones") = 0
adoNomina.Recordset("TotalIncentivos") = 0
adoNomina.Recordset("TotalDeducciones") = 0
adoNomina.Recordset("TotalPrestamo") = 0
adoNomina.Recordset("TotalMontoInss") = 0
adoNomina.Recordset("TotalMontoIR") = 0
adoNomina.Recordset("TotalVacaciones") = 0
adoNomina.Recordset("TotalINSSPatronal") = 0
adoNomina.Recordset("TotalIRPatronal") = 0
adoNomina.Recordset("Anulada") = 0
adoNomina.Recordset("Cerrada") = 0
adoNomina.Recordset("Procesada") = 0
adoNomina.Recordset("Mes") = Mes
adoNomina.Recordset("Ano") = Año
adoNomina.Recordset("Periodo") = Periodo
adoNomina.Recordset.Update

sql = "SELECT  * From Empleado Where (Activo = 1)"
Me.AdoEmpleados.RecordSource = sql
Me.AdoEmpleados.Refresh

 If Not Me.AdoEmpleados.Recordset.EOF Then
    Me.AdoEmpleados.Recordset.MoveLast
    Me.Barra.Visible = True
    Me.Barra.Min = 0
    Me.Barra.Max = Me.AdoEmpleados.Recordset.RecordCount
    Me.Barra.Value = 0
    Me.AdoEmpleados.Recordset.MoveFirst
 End If
 
 i = 0
Do While Not Me.AdoEmpleados.Recordset.EOF

        DoEvents

        CodEmpleado = Me.AdoEmpleados.Recordset("CodEmpleado")

        DtaDetalleNomina.Recordset.AddNew
        DtaDetalleNomina.Recordset("NumNomina") = Me.TxtNumNomina.Text
        DtaDetalleNomina.Recordset("SeptimoDia") = 0
        DtaDetalleNomina.Recordset("BonoProduccion") = 0
        DtaDetalleNomina.Recordset("TarifaHoraria") = 0
        DtaDetalleNomina.Recordset("IncetivoProduccion") = 0
        DtaDetalleNomina.Recordset("HTrabajada") = 0
        DtaDetalleNomina.Recordset("CodEmpleado") = CodEmpleado
        DtaDetalleNomina.Recordset("produjo") = "N"
        DtaDetalleNomina.Recordset("SalarioBasico") = 0
        DtaDetalleNomina.Recordset("destajo") = 0
        DtaDetalleNomina.Recordset("HE") = 0
        DtaDetalleNomina.Recordset("HorasExtras") = 0
        DtaDetalleNomina.Recordset("Comisiones") = 0
        DtaDetalleNomina.Recordset("incentivos") = 0
        DtaDetalleNomina.Recordset("OtrosIngresos") = 0
        DtaDetalleNomina.Recordset("DescripOtrIngre") = "Ninguno"
        DtaDetalleNomina.Recordset("Deducciones") = 0
        DtaDetalleNomina.Recordset("Prestamo") = 0
        DtaDetalleNomina.Recordset("MontoInss") = 0
        DtaDetalleNomina.Recordset("MontoIR") = 0
        DtaDetalleNomina.Recordset("Vacaciones") = 0
        DtaDetalleNomina.Recordset("Mes13") = 0
        DtaDetalleNomina.Recordset("INSSPatronal") = 0
        DtaDetalleNomina.Recordset("IRPatronal") = 0
        DtaDetalleNomina.Recordset("INATEC") = 0
        DtaDetalleNomina.Recordset("DD") = 0
        DtaDetalleNomina.Recordset("Adelantos") = 0
        DtaDetalleNomina.Recordset("DiasDescuento") = 0
        DtaDetalleNomina.Recordset.Update
  Me.AdoEmpleados.Recordset.MoveNext
  i = i + 1
  Me.Barra.Value = i
  DoEvents
Loop

MsgBox "Proceso Terminado!!!", vbExclamation, "Zeus Nominas"

    Me.AdoDetalleNominaAcumulada.RecordSource = "SELECT  Empleado.CodEmpleado1, Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2 AS Nombres, DetalleNominaAcumulada.SalarioBasico, DetalleNominaAcumulada.Destajo, DetalleNominaAcumulada.SeptimoDia, DetalleNominaAcumulada.HorasExtras, DetalleNominaAcumulada.Comisiones, DetalleNominaAcumulada.Incentivos, DetalleNominaAcumulada.IncetivoProduccion, DetalleNominaAcumulada.BonoProduccion,DetalleNominaAcumulada.OtrosIngresos, DetalleNominaAcumulada.TIngresos AS TotalIngresos, DetalleNominaAcumulada.Deducciones, DetalleNominaAcumulada.Prestamo, DetalleNominaAcumulada.MontoINSS, DetalleNominaAcumulada.MontoIR, DetalleNominaAcumulada.INSSPatronal, DetalleNominaAcumulada.IRPatronal, DetalleNominaAcumulada.INATEC, DetalleNominaAcumulada.TGastos AS TotalDeducciones  " & _
                                                "FROM  Empleado INNER JOIN DetalleNominaAcumulada ON Empleado.CodEmpleado = DetalleNominaAcumulada.CodEmpleado Where (DetalleNominaAcumulada.NumNomina = 0) ORDER BY Empleado.CodEmpleado1"
    Me.AdoDetalleNominaAcumulada.Refresh

    Me.DbgrdetalleNominas.Columns(0).Width = 1000
    Me.DbgrdetalleNominas.Columns(0).Locked = True
    Me.DbgrdetalleNominas.Columns(1).Width = 4000
    Me.DbgrdetalleNominas.Columns(0).Locked = True
    Me.DbgrdetalleNominas.Columns(2).Width = 1200
    Me.DbgrdetalleNominas.Columns(2).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(3).Width = 1200
    Me.DbgrdetalleNominas.Columns(3).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(4).Width = 1200
    Me.DbgrdetalleNominas.Columns(4).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(5).Width = 1200
    Me.DbgrdetalleNominas.Columns(5).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(6).Width = 1200
    Me.DbgrdetalleNominas.Columns(6).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(7).Width = 1200
    Me.DbgrdetalleNominas.Columns(7).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(8).Width = 1200
    Me.DbgrdetalleNominas.Columns(8).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(9).Width = 1200
    Me.DbgrdetalleNominas.Columns(9).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(10).Width = 1200
    Me.DbgrdetalleNominas.Columns(10).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(11).Width = 1200
    Me.DbgrdetalleNominas.Columns(11).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(11).Locked = True
    Me.DbgrdetalleNominas.Columns(12).Width = 1200
    Me.DbgrdetalleNominas.Columns(12).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(13).Width = 1200
    Me.DbgrdetalleNominas.Columns(13).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(14).Width = 1200
    Me.DbgrdetalleNominas.Columns(14).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(15).Width = 1200
    Me.DbgrdetalleNominas.Columns(15).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(16).Width = 1200
    Me.DbgrdetalleNominas.Columns(16).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(17).Width = 1200
    Me.DbgrdetalleNominas.Columns(17).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(18).Width = 1200
    Me.DbgrdetalleNominas.Columns(18).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(19).Width = 1200
    Me.DbgrdetalleNominas.Columns(19).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(19).Locked = True
    
      Me.CmdGenerar.Enabled = False
      Me.DBComboPeriodo.Enabled = False
      Me.DBCNominas.Enabled = False

End Sub



Private Sub DBCNominas_Change()
Dim Año As Double, Periodo As Double, Fecha1 As Date, Fecha2 As Date


'///////////////////////////////BUSCO EL TIPO DE NOMINAS /////////////////////////////////////
 CodTipoNomina = Me.DBCNominas.Columns(0).Text

'//////////////////////////////BUSCO EL PERIDO ACTIVO ////////////////////////////////////////
   Me.AdoConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
   Me.AdoConsulta.Refresh
   If Not AdoConsulta.Recordset.EOF Then
      Año = Me.AdoConsulta.Recordset("año")
      Me.txtAno.Text = Me.AdoConsulta.Recordset("año")
      Periodo = Me.AdoConsulta.Recordset("Periodo")
      Fecha1 = Me.AdoConsulta.Recordset("Inicio")
      Fecha2 = Me.AdoConsulta.Recordset("Final")
   End If
   
   
Me.DtaPeriodos.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (Calculada = 0)AND (año = " & Año & ")"
Me.DtaPeriodos.Refresh
End Sub

Private Sub DBCNominas_Click()
Dim Año As Double, Periodo As Double, Fecha1 As Date, Fecha2 As Date


'///////////////////////////////BUSCO EL TIPO DE NOMINAS /////////////////////////////////////
 CodTipoNomina = Me.DBCNominas.Columns(0).Text

'//////////////////////////////BUSCO EL PERIDO ACTIVO ////////////////////////////////////////
   Me.AdoConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
   Me.AdoConsulta.Refresh
   If Not AdoConsulta.Recordset.EOF Then
      Año = Me.AdoConsulta.Recordset("año")
      Me.txtAno.Text = Me.AdoConsulta.Recordset("año")
      Periodo = Me.AdoConsulta.Recordset("Periodo")
      Fecha1 = Me.AdoConsulta.Recordset("Inicio")
      Fecha2 = Me.AdoConsulta.Recordset("Final")
   End If
   
   
Me.DtaPeriodos.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (Calculada = 0)AND (año = " & Año & ")"
Me.DtaPeriodos.Refresh
End Sub

Private Sub DBCNominas_DblClick()
Dim Año As Double, Periodo As Double, Fecha1 As Date, Fecha2 As Date


'///////////////////////////////BUSCO EL TIPO DE NOMINAS /////////////////////////////////////
 CodTipoNomina = Me.DBCNominas.Columns(0).Text

'//////////////////////////////BUSCO EL PERIDO ACTIVO ////////////////////////////////////////
   Me.AdoConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
   Me.AdoConsulta.Refresh
   If Not AdoConsulta.Recordset.EOF Then
      Año = Me.AdoConsulta.Recordset("año")
      Me.txtAno.Text = Me.AdoConsulta.Recordset("año")
      Periodo = Me.AdoConsulta.Recordset("Periodo")
      Fecha1 = Me.AdoConsulta.Recordset("Inicio")
      Fecha2 = Me.AdoConsulta.Recordset("Final")
   End If
   
   
Me.DtaPeriodos.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (Calculada = 0)AND (año = " & Año & ")"
Me.DtaPeriodos.Refresh
End Sub

Private Sub DBComboPeriodo_Change()
Dim Periodo As Integer, Ano As Integer
CodTipoNomina = Me.DBCNominas.Columns(0).Text
Periodo = val(Me.DBComboPeriodo.Text)


Ano = val(Me.txtAno.Text)
Me.AdoConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Periodo & ")AND (año = " & Ano & ")"
'InputBox "", "", Me.DtaConsulta.RecordSource
Me.AdoConsulta.Refresh

If Not AdoConsulta.Recordset.EOF Then
   Me.DTFechaFin.Value = Me.AdoConsulta.Recordset("Final")
   Me.DtaFechaINI.Value = Me.AdoConsulta.Recordset("Inicio")
'   LblFechaLarga.Caption = "Pago: " + Format(MtxtFecha.Value, "Long Date")
End If
End Sub

Private Sub DbgrdetalleNominas_AfterColUpdate(ByVal ColIndex As Integer)
Dim TotalIngreso As Double, i As Double, TotalGastos As Double


TotalIngresos = CDbl(Me.DbgrdetalleNominas.Columns(2).Text) + CDbl(Me.DbgrdetalleNominas.Columns(3).Text) + CDbl(Me.DbgrdetalleNominas.Columns(4).Text) + CDbl(Me.DbgrdetalleNominas.Columns(5).Text) + CDbl(Me.DbgrdetalleNominas.Columns(6).Text) + CDbl(Me.DbgrdetalleNominas.Columns(7).Text) + CDbl(Me.DbgrdetalleNominas.Columns(8).Text) + CDbl(Me.DbgrdetalleNominas.Columns(9).Text) + CDbl(Me.DbgrdetalleNominas.Columns(10).Text)
TotalGastos = CDbl(Me.DbgrdetalleNominas.Columns(12).Text) + CDbl(Me.DbgrdetalleNominas.Columns(13).Text) + CDbl(Me.DbgrdetalleNominas.Columns(14).Text) + CDbl(Me.DbgrdetalleNominas.Columns(15).Text) + CDbl(Me.DbgrdetalleNominas.Columns(16).Text) + CDbl(Me.DbgrdetalleNominas.Columns(17).Text) + CDbl(Me.DbgrdetalleNominas.Columns(18).Text)
Me.DbgrdetalleNominas.Columns(11).Text = TotalIngresos
Me.DbgrdetalleNominas.Columns(19).Text = TotalGastos



End Sub

Private Sub DbgrdetalleNominas_AfterUpdate()
'/////////////////////////SUMO LOS TOTALES /////////////////////////
Me.AdoTotales.RecordSource = "SELECT  SUM(SalarioBasico) AS SalarioBasico, SUM(Destajo) AS Destajo, SUM(HorasExtras) AS HorasExtras, SUM(Comisiones) AS Comisiones, SUM(OtrosIngresos) AS OtrosIngresos, SUM(Incentivos) AS Incentivos, SUM(SeptimoDia) AS SeptimoDia, SUM(IncetivoProduccion) AS IncentivoProduccion, SUM(BonoProduccion) AS BonoProduccion, SUM(TIngresos) AS TIngresos, SUM(TGastos) AS TGastos, SUM(Deducciones) AS Deducciones, SUM(Prestamo) AS Prestamo, SUM(MontoINSS) AS MontoINSS, SUM(MontoIR) AS MontoIR, SUM(Adelantos) AS Adelantos From DetalleNomina Where (NumNomina = 0)"
Me.AdoTotales.Refresh

Me.adoNomina.RecordSource = "SELECT  * From Nomina Where (NumNomina = 0)"
Me.adoNomina.Refresh
If Not Me.adoNomina.Recordset.EOF Then
        
    adoNomina.Recordset("TotalSalarioBasico") = Me.AdoTotales.Recordset("SalarioBasico")
    adoNomina.Recordset("TotalDestajo") = Me.AdoTotales.Recordset("Destajo")
    adoNomina.Recordset("TotalHorasExtras") = Me.AdoTotales.Recordset("HorasExtras")
    adoNomina.Recordset("TotalComisiones") = Me.AdoTotales.Recordset("Comisiones")
    adoNomina.Recordset("TotalIncentivos") = Me.AdoTotales.Recordset("OtrosIngresos")
    adoNomina.Recordset("TotalDeducciones") = Me.AdoTotales.Recordset("Deducciones")
    adoNomina.Recordset("TotalPrestamo") = Me.AdoTotales.Recordset("Prestamo")
    adoNomina.Recordset("TotalMontoInss") = Me.AdoTotales.Recordset("MontoINSS")
    adoNomina.Recordset("TotalMontoIR") = Me.AdoTotales.Recordset("MontoIR")
'    AdoNomina.Recordset("TotalVacaciones") = Me.AdoTotales.Recordset("Prestamos")
    adoNomina.Recordset("TotalINSSPatronal") = 0
    adoNomina.Recordset("TotalIRPatronal") = 0
    adoNomina.Recordset.Update

End If


End Sub

Private Sub Form_Load()
 Dim Año As Double, CodTipoNomina As String
Dim Periodo As Double, Mes As Double
Me.BackColor = RGB(222, 227, 247)
Me.Frame1.BackColor = RGB(222, 227, 247)

 Me.DbgrdetalleNominas.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DbgrdetalleNominas.OddRowStyle.BackColor = &H80000005
 Me.DbgrdetalleNominas.AlternatingRowStyle = True

    With Me.AdoEmpleados
       .ConnectionString = Conexion
    End With
    
    With Me.AdoTotales
       .ConnectionString = Conexion
    End With
    
    With Me.AdoConsulta
       .ConnectionString = Conexion
    End With
    
    With Me.DtaPeriodos
       .ConnectionString = Conexion
    End With
    
    With Me.adoNomina
       .ConnectionString = Conexion
    End With
    
    With Me.DtaDetalleNomina
       .ConnectionString = Conexion
    End With
    
    With Me.AdoDetalleNominaAcumulada
       .ConnectionString = Conexion
    End With
    
    With Me.DtaTipoNomina
    .ConnectionString = Conexion
    .RecordSource = "SELECT CodTipoNomina, Nomina, Periodo FROM   TipoNomina"
    .Refresh
    End With
    
'    Me.AdoDetalleNominaAcumulada.RecordSource = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres,  DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.SeptimoDia, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.Incentivos, DetalleNomina.IncetivoProduccion, DetalleNomina.BonoProduccion, DetalleNomina.OtrosIngresos, TIngresos AS TotalIngresos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.INSSPatronal, " & _
'                                                "DetalleNomina.IRPatronal, DetalleNomina.INATEC,TGastos AS TotalDeducciones FROM   DetalleNomina INNER JOIN  Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado Where (DetalleNomina.NumNomina = 0) ORDER BY Empleado.CodEmpleado1"
    
    Me.AdoDetalleNominaAcumulada.RecordSource = "SELECT  Empleado.CodEmpleado1, Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2 AS Nombres, DetalleNominaAcumulada.SalarioBasico, DetalleNominaAcumulada.Destajo, DetalleNominaAcumulada.SeptimoDia, DetalleNominaAcumulada.HorasExtras, DetalleNominaAcumulada.Comisiones, DetalleNominaAcumulada.Incentivos, DetalleNominaAcumulada.IncetivoProduccion, DetalleNominaAcumulada.BonoProduccion,DetalleNominaAcumulada.OtrosIngresos, DetalleNominaAcumulada.TIngresos AS TotalIngresos, DetalleNominaAcumulada.Deducciones, DetalleNominaAcumulada.Prestamo, DetalleNominaAcumulada.MontoINSS, DetalleNominaAcumulada.MontoIR, DetalleNominaAcumulada.INSSPatronal, DetalleNominaAcumulada.IRPatronal, DetalleNominaAcumulada.INATEC, DetalleNominaAcumulada.TGastos AS TotalDeducciones  " & _
                                                "FROM  Empleado INNER JOIN DetalleNominaAcumulada ON Empleado.CodEmpleado = DetalleNominaAcumulada.CodEmpleado Where (DetalleNominaAcumulada.NumNomina = 0) ORDER BY Empleado.CodEmpleado1"
    Me.AdoDetalleNominaAcumulada.Refresh
    
    If Not Me.AdoDetalleNominaAcumulada.Recordset.EOF Then
          Me.CmdGenerar.Enabled = False
          Me.DBComboPeriodo.Enabled = False
          Me.DBCNominas.Enabled = False
          
          '//////////////////////////////////////////BUSCO LA INFORMACION BASICA DE LA NOMINA ///////////////////////////
'        Me.AdoConsulta.RecordSource = "SELECT Nomina.NumNomina, Nomina.CodTipoNomina, TipoNomina.Nomina, Nomina.FechaNominaINI, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano, Nomina.Periodo FROM  Nomina INNER JOIN TipoNomina ON Nomina.CodTipoNomina = TipoNomina.CodTipoNomina Where (Nomina.NumNomina = 0)"
        Me.AdoConsulta.RecordSource = "SELECT TipoNomina.Nomina, NominaAcumulada.NumNomina, NominaAcumulada.CodTipoNomina, NominaAcumulada.FechaNominaINI, NominaAcumulada.FechaNomina, NominaAcumulada.Mes , NominaAcumulada.Ano, NominaAcumulada.Periodo FROM  TipoNomina INNER JOIN NominaAcumulada ON TipoNomina.CodTipoNomina = NominaAcumulada.CodTipoNomina Where (NominaAcumulada.NumNomina = 0)"
        Me.AdoConsulta.Refresh
        If Not Me.AdoConsulta.Recordset.EOF Then
          Mes = Me.AdoConsulta.Recordset("Mes")
          Año = Me.AdoConsulta.Recordset("Ano")
          Me.txtAno.Text = Me.AdoConsulta.Recordset("Ano")
          Periodo = Me.AdoConsulta.Recordset("Periodo")
          Me.DtaFechaINI.Value = Me.AdoConsulta.Recordset("FechaNominaINI")
          Me.DTFechaFin.Value = Me.AdoConsulta.Recordset("FechaNomina")
          Me.DBCNominas.Text = Me.AdoConsulta.Recordset("Nomina")
          Me.DBComboPeriodo.Text = Periodo
        End If
      
    Else
      Me.CmdGenerar.Enabled = True
      Me.DBComboPeriodo.Enabled = True
      Me.DBCNominas.Enabled = True
    End If
    
    Me.DbgrdetalleNominas.Columns(0).Width = 1000
    Me.DbgrdetalleNominas.Columns(0).Locked = True
    Me.DbgrdetalleNominas.Columns(1).Width = 4000
    Me.DbgrdetalleNominas.Columns(0).Locked = True
    Me.DbgrdetalleNominas.Columns(2).Width = 1200
    Me.DbgrdetalleNominas.Columns(2).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(3).Width = 1200
    Me.DbgrdetalleNominas.Columns(3).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(4).Width = 1200
    Me.DbgrdetalleNominas.Columns(4).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(5).Width = 1200
    Me.DbgrdetalleNominas.Columns(5).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(6).Width = 1200
    Me.DbgrdetalleNominas.Columns(6).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(7).Width = 1200
    Me.DbgrdetalleNominas.Columns(7).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(8).Width = 1200
    Me.DbgrdetalleNominas.Columns(8).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(9).Width = 1200
    Me.DbgrdetalleNominas.Columns(9).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(10).Width = 1200
    Me.DbgrdetalleNominas.Columns(10).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(11).Width = 1200
    Me.DbgrdetalleNominas.Columns(11).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(11).Locked = True
    Me.DbgrdetalleNominas.Columns(12).Width = 1200
    Me.DbgrdetalleNominas.Columns(12).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(13).Width = 1200
    Me.DbgrdetalleNominas.Columns(13).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(14).Width = 1200
    Me.DbgrdetalleNominas.Columns(14).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(15).Width = 1200
    Me.DbgrdetalleNominas.Columns(15).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(16).Width = 1200
    Me.DbgrdetalleNominas.Columns(16).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(17).Width = 1200
    Me.DbgrdetalleNominas.Columns(17).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(18).Width = 1200
    Me.DbgrdetalleNominas.Columns(18).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(19).Width = 1200
    Me.DbgrdetalleNominas.Columns(19).NumberFormat = "##,##0.00"
    Me.DbgrdetalleNominas.Columns(19).Locked = True
End Sub

Private Sub PushButton1_Click()
 FrmConsultas.Show 1
End Sub
