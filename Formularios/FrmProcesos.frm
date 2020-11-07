VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form FrmProcesos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesos"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   584
   Begin MSAdodcLib.Adodc AdoBusca 
      Height          =   450
      Left            =   4080
      Top             =   7560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   794
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   8775
      TabIndex        =   29
      Top             =   0
      Width           =   8775
      Begin VB.Image Image2 
         Height          =   1020
         Left            =   120
         Picture         =   "FrmProcesos.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1290
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   8760
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTROS DE PROCESOS"
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
         Left            =   1920
         TabIndex        =   30
         Top             =   360
         Width           =   3840
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   22
      Top             =   5160
      Width           =   3135
      Begin VB.CommandButton CmdSiguiente 
         DownPicture     =   "FrmProcesos.frx":098A
         Height          =   375
         Left            =   1560
         MouseIcon       =   "FrmProcesos.frx":246C
         MousePointer    =   99  'Custom
         Picture         =   "FrmProcesos.frx":28AE
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdAnterior 
         DownPicture     =   "FrmProcesos.frx":4390
         Height          =   375
         Left            =   120
         MouseIcon       =   "FrmProcesos.frx":5E72
         MousePointer    =   99  'Custom
         Picture         =   "FrmProcesos.frx":62B4
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdPirmero 
         DownPicture     =   "FrmProcesos.frx":7D96
         Height          =   375
         Left            =   120
         MouseIcon       =   "FrmProcesos.frx":9878
         MousePointer    =   99  'Custom
         Picture         =   "FrmProcesos.frx":9CBA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton CmdUltimo 
         DownPicture     =   "FrmProcesos.frx":B79C
         Height          =   375
         Left            =   1560
         MouseIcon       =   "FrmProcesos.frx":D27E
         MousePointer    =   99  'Custom
         Picture         =   "FrmProcesos.frx":D6C0
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdGrabar 
      DownPicture     =   "FrmProcesos.frx":F1A2
      Height          =   375
      Left            =   3480
      MouseIcon       =   "FrmProcesos.frx":10C84
      MousePointer    =   99  'Custom
      Picture         =   "FrmProcesos.frx":110C6
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalir 
      DownPicture     =   "FrmProcesos.frx":12BA8
      Height          =   375
      Left            =   5400
      MouseIcon       =   "FrmProcesos.frx":1468A
      MousePointer    =   99  'Custom
      Picture         =   "FrmProcesos.frx":14ACC
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton CmdBorrar 
      DownPicture     =   "FrmProcesos.frx":165AE
      Height          =   375
      Left            =   3480
      MouseIcon       =   "FrmProcesos.frx":18090
      MousePointer    =   99  'Custom
      Picture         =   "FrmProcesos.frx":184D2
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   8415
      Begin TrueOleDBList80.TDBCombo DBCodigoReferencia 
         Bindings        =   "FrmProcesos.frx":19FB4
         Height          =   315
         Left            =   1320
         TabIndex        =   31
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
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
         ListField       =   "CodReferencia1"
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
         _PropDict       =   $"FrmProcesos.frx":19FD0
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=88,.bold=0,.fontsize=825,.italic=0"
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
      Begin VB.TextBox TxtDescripcion 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   840
         Width           =   8055
      End
      Begin VB.TextBox TxtReferencia 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1095
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
         Left            =   7920
         Picture         =   "FrmProcesos.frx":1A07A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblCodRef 
         AutoSize        =   -1  'True
         Caption         =   "Codigo de Ref"
         Height          =   195
         Left            =   1440
         TabIndex        =   18
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label lblRef 
         AutoSize        =   -1  'True
         Caption         =   "Referencia"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   780
      End
   End
   Begin VB.TextBox TxtBReferencia 
      Height          =   285
      Left            =   3960
      TabIndex        =   12
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      DownPicture     =   "FrmProcesos.frx":1A1C8
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Picture         =   "FrmProcesos.frx":1BCAA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox TxtBProceso 
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Top             =   4560
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc DtaProcesos 
      Height          =   375
      Left            =   480
      Top             =   7440
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "DtaProcesos"
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
      Left            =   480
      Top             =   7800
      Width           =   3255
      _ExtentX        =   5741
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
   Begin MSAdodcLib.Adodc DtaReferencia 
      Height          =   375
      Left            =   480
      Top             =   7080
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "DtaReferencia"
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
   Begin VB.Frame fraProc 
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   8415
      Begin VB.CommandButton Command1 
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
         Left            =   1680
         Picture         =   "FrmProcesos.frx":1D78C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin MSDataListLib.DataCombo DBCodigoProcesos 
         Bindings        =   "FrmProcesos.frx":1D8DA
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "CodProceso"
         Text            =   ""
      End
      Begin VB.TextBox TxtDescripcionProceso 
         Height          =   615
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   840
         Width           =   6495
      End
      Begin VB.TextBox TxtPrecio 
         Alignment       =   1  'Right Justify
         DataField       =   "Precio"
         DataSource      =   "dtcProc"
         Height          =   285
         Left            =   5400
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtUnidades 
         Alignment       =   2  'Center
         DataField       =   "Unid"
         DataSource      =   "dtcProc"
         Height          =   285
         Left            =   7080
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblProc 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblprec 
         AutoSize        =   -1  'True
         Caption         =   "Precio por unds"
         Height          =   195
         Left            =   5400
         TabIndex        =   5
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label lblUnd 
         AutoSize        =   -1  'True
         Caption         =   "Unidades"
         Height          =   195
         Left            =   7200
         TabIndex        =   4
         Top             =   600
         Width           =   675
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   840
         TabIndex        =   3
         Top             =   960
         Width           =   840
      End
   End
   Begin TrueOleDBList80.TDBCombo TDBReferencia2 
      Bindings        =   "FrmProcesos.frx":1D8F4
      Height          =   315
      Left            =   3120
      TabIndex        =   32
      Top             =   4560
      Width           =   5535
      _ExtentX        =   9763
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
      ListField       =   "CodReferencia1"
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
      _PropDict       =   $"FrmProcesos.frx":1D910
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=88,.bold=0,.fontsize=825,.italic=0"
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
      Caption         =   "Referencia"
      Height          =   255
      Left            =   2160
      TabIndex        =   28
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Proceso"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   4560
      Width           =   735
   End
End
Attribute VB_Name = "FrmProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAnterior_Click()
 Me.DtaProcesos.Recordset.MovePrevious
 If Not Me.DtaProcesos.Recordset.BOF Then
   Me.DBCodigoProcesos.Text = Me.DtaProcesos.Recordset("CodProceso")
 Else
   MsgBox "Este es el Primer Registro", vbExclamation, "sistema de Nominas"
   Me.DtaProcesos.Recordset.MoveNext
 End If
End Sub

Private Sub cmdborrar_Click()
On Error GoTo TipoErrs
Dim Respuesta, Rsp
If IsNull(Me.DtaProcesos.Recordset("CodProceso")) = True Then
        MsgBox "No Existe Registro"
        Exit Sub
End If
If Me.DBCodigoProcesos.Text = "" Then
 MsgBox "No se Puede Eliminar este Registro", vbInformation, "Sistema de Nominas"
 Exit Sub
End If
'Elimino el registro activo en la pantalla

  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando Proceso: " & Me.TxtDescripcionProceso.Text)
   If Respuesta = 6 Then
     Me.DtaConsulta.RecordSource = "SELECT Ref, CodReferencia, Descrip From Referencia WHERE     (CodReferencia = '" & Me.DBCodigoReferencia.Text & "')"
     Me.DtaConsulta.Refresh
     If Not Me.DtaConsulta.Recordset.EOF Then
       Me.DtaConsulta.Recordset.Delete
     End If
     If IsNull(Me.DtaProcesos.Recordset("CodProceso")) = True Then
        MsgBox "No Existe Registro"
        Exit Sub
     Else
      Me.DBCodigoProcesos.Text = ""
      Me.DtaProcesos.Refresh
     End If
 End If
Exit Sub
TipoErrs:
   ControlErrores
   MsgBox Err.Description
 Unload Me
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo TipoErrs
 Me.DtaConsulta.RecordSource = "SELECT Ref, CodReferencia, Descrip From Referencia WHERE     (CodReferencia = " & Me.DBCodigoReferencia.Columns(2).Text & ")"
 Me.DtaConsulta.Refresh
 If Me.DtaConsulta.Recordset.EOF Then
   MsgBox "No Existe el Codigo de Referencia", vbCritical, "Sistema de Nominas"
   Exit Sub
 End If
 
 If Not IsNumeric(Me.TxtPrecio.Text) Then
  MsgBox "El precio Digitado no es Numerico", vbCritical, "Sistema de Nominas"
  Exit Sub
 End If
 
  If Not IsNumeric(Me.TxtUnidades.Text) Then
  MsgBox "Las Unidades Digitadas no es Numerico", vbCritical, "Sistema de Nominas"
  Exit Sub
 End If
 
 If Me.DBCodigoReferencia.Text = "" Then
  MsgBox "Se necesita seleccionar una referencia", vbCritical, "Sistema de Nominas"
  Exit Sub
 End If
 
 
 Me.DtaConsulta.RecordSource = "SELECT Ref, CodReferencia, CodProceso, Descrip, Precio, Unid From Procesos WHERE     (CodReferencia = " & Me.DBCodigoReferencia.Columns(2).Text & ") AND (CodProceso = '" & Me.DBCodigoProcesos.Text & "') "
 ' Me.DtaConsulta.RecordSource = "SELECT Ref, CodReferencia, CodProceso, Descrip, Precio, Unid From Procesos Where (CodProceso= '" & Me.DBCodigoProcesos.Text & "')"
 Me.DtaConsulta.Refresh
 If Me.DtaConsulta.Recordset.EOF Then
   Me.DtaConsulta.Recordset.AddNew
   Me.DtaConsulta.Recordset("Ref") = Me.TxtReferencia
   Me.DtaConsulta.Recordset("CodProceso") = Me.DBCodigoProcesos.Text
   Me.DtaConsulta.Recordset("Descrip") = Me.TxtDescripcionProceso.Text
   Me.DtaConsulta.Recordset("Precio") = Me.TxtPrecio.Text
   Me.DtaConsulta.Recordset("Unid") = Me.TxtUnidades.Text
   Me.DtaConsulta.Recordset("CodReferencia") = Me.DBCodigoReferencia.Columns(2).Text
   Me.DtaConsulta.Recordset.Update
  Else
   Me.DtaConsulta.Recordset("Ref") = Me.TxtReferencia
   Me.DtaConsulta.Recordset("Descrip") = Me.TxtDescripcionProceso.Text
   Me.DtaConsulta.Recordset("Precio") = Me.TxtPrecio.Text
   Me.DtaConsulta.Recordset("Unid") = Me.TxtUnidades.Text
'   Me.DtaConsulta.Recordset("CodReferencia1") = Me.DBCodigoReferencia.Text
   Me.DtaConsulta.Recordset.Update
  End If
  
  Me.DtaProcesos.Refresh
  Me.DBCodigoProcesos.Text = ""
Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub CmdPirmero_Click()
 Me.DtaProcesos.Recordset.MoveFirst
 If Not Me.DtaProcesos.Recordset.BOF Then
   Me.DBCodigoProcesos.Text = Me.DtaProcesos.Recordset("CodProceso")
 Else
   MsgBox "Este es el Primer Registro", vbExclamation, "sistema de Nominas"
   Me.DtaProcesos.Recordset.MoveNext
 End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DataCombo1_Change()


End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub CmdSiguiente_Click()
 Me.DtaProcesos.Recordset.MoveNext
 If Not Me.DtaProcesos.Recordset.EOF Then
   Me.DBCodigoProcesos.Text = Me.DtaProcesos.Recordset("CodProceso")
 Else
   MsgBox "Este es el Primer Registro", vbExclamation, "sistema de Nominas"
   Me.DtaProcesos.Recordset.MovePrevious
 End If
End Sub

Private Sub CmdUltimo_Click()
 Me.DtaProcesos.Recordset.MoveLast
 If Not Me.DtaProcesos.Recordset.EOF Then
   Me.DBCodigoProcesos.Text = Me.DtaProcesos.Recordset("CodProceso")
 Else
   MsgBox "Este es el Primer Registro", vbExclamation, "sistema de Nominas"
   Me.DtaProcesos.Recordset.MovePrevious
 End If
End Sub

Private Sub Command2_Click()
 Me.DtaConsulta.RecordSource = "SELECT Ref, CodReferencia, CodProceso, Descrip, Precio, Unid From Procesos WHERE (CodProceso = '" & Me.TxtBProceso & "') AND (CodReferencia = " & Me.TDBReferencia2.Columns(2).Text & ")"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then

    Me.DBCodigoReferencia.Text = Me.TDBReferencia2.Columns(1).Text
    
    Me.DtaConsulta.RecordSource = "SELECT Ref, CodReferencia, Descrip From Referencia WHERE     (CodReferencia = '" & Me.TDBReferencia2.Columns(2).Text & "')"
    Me.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
      Me.TxtReferencia.Text = Me.DtaConsulta.Recordset("Ref")
      Me.TxtDescripcion.Text = Me.DtaConsulta.Recordset("Descrip")
    Else
      Me.TxtReferencia.Text = ""
      Me.TxtDescripcion.Text = ""
    End If
   
   
'   Me.DBCodigoReferencia.Text = Me.TxtBReferencia
   Me.DBCodigoProcesos.Text = Me.TxtBProceso
   
   
   
 Else
   MsgBox "No Existe este Proceso", vbExclamation, "Sistema de Nominas"

   Me.DBCodigoReferencia.Text = ""
      Me.DBCodigoProcesos.Text = ""
 End If
End Sub

Private Sub DBCodigoProcesos_Change()
Dim CodigoReferencia As String
 Me.DtaConsulta.RecordSource = "SELECT Ref, CodReferencia, CodProceso, Descrip, Precio, Unid From Procesos WHERE (CodProceso = '" & Me.DBCodigoProcesos & "') AND (CodReferencia = " & Me.DBCodigoReferencia.Columns(2).Text & ")"
 Me.DtaConsulta.Refresh
 If Not Me.DtaConsulta.Recordset.EOF Then
   Me.TxtDescripcionProceso.Text = Me.DtaConsulta.Recordset("Descrip")
   Me.TxtPrecio.Text = Format(Me.DtaConsulta.Recordset("Precio"), "##,##0.00")
   Me.TxtUnidades.Text = Me.DtaConsulta.Recordset("Unid")
'   Me.DBCodigoReferencia.Text = Me.DtaConsulta.Recordset("CodReferencia")
  Else
   Me.TxtDescripcionProceso.Text = ""
   Me.TxtPrecio.Text = ""
   Me.TxtUnidades.Text = ""
   Me.DBCodigoReferencia.Text = ""
  End If
End Sub

Private Sub DBCodigoReferencia_Change()
 Me.DtaConsulta.RecordSource = "SELECT Ref, CodReferencia, Descrip From Referencia WHERE  (CodReferencia1 = '" & Me.DBCodigoReferencia.Text & "') and (Activo=1)"
 Me.DtaConsulta.Refresh
 If Not Me.DtaConsulta.Recordset.EOF Then
    Me.TxtReferencia.Text = Me.DtaConsulta.Recordset("Ref")
    Me.TxtDescripcion.Text = Me.DtaConsulta.Recordset("Descrip")
 Else
    Me.TxtReferencia.Text = ""
    Me.TxtDescripcion.Text = ""
 End If
End Sub

Private Sub DBCodigoReferencia_ItemChange()
 Me.DtaConsulta.RecordSource = "SELECT Ref, CodReferencia, Descrip From Referencia WHERE     (CodReferencia1 = '" & Me.DBCodigoReferencia.Text & "') and (Activo=1)"
 Me.DtaConsulta.Refresh
 If Not Me.DtaConsulta.Recordset.EOF Then
    Me.TxtReferencia.Text = Me.DtaConsulta.Recordset("Ref")
    Me.TxtDescripcion.Text = Me.DtaConsulta.Recordset("Descrip")
 Else
    Me.TxtReferencia.Text = ""
    Me.TxtDescripcion.Text = ""
 End If
End Sub

Private Sub Form_Activate()
 With Me.DtaReferencia
    .ConnectionString = Conexion
    .RecordSource = "SELECT Ref, CodReferencia1, CodReferencia, Descrip, Activo From Referencia Where (Activo = 1) ORDER BY CodReferencia1"
    .Refresh
 End With
End Sub

Private Sub Form_Load()
 With Me.DtaReferencia
    .ConnectionString = Conexion
    .RecordSource = "SELECT Ref, CodReferencia1, CodReferencia, Descrip, Activo From Referencia Where (Activo = 1) ORDER BY CodReferencia1"
    .Refresh
 End With
 
  With Me.DtaConsulta
    .ConnectionString = Conexion
 
 End With
 
 With Me.AdoBusca
    .ConnectionString = Conexion
 End With
 
  With Me.DtaProcesos
    .ConnectionString = Conexion
    .RecordSource = "Procesos"
    .Refresh
 End With
 



End Sub

Private Sub TxtPrecio_LostFocus()
 Me.TxtPrecio.Text = Format(Me.TxtPrecio.Text, "##,##0.00")
End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub
