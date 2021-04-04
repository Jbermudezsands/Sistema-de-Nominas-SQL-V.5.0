VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmHistorial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial Salarial"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   537
   ScaleMode       =   0  'User
   ScaleWidth      =   1219.355
   Begin VB.CommandButton CmdConstanciaActivos 
      Caption         =   "Constancia"
      Height          =   1095
      Left            =   360
      Picture         =   "FrmHistorial.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   5280
      Top             =   8640
      Visible         =   0   'False
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
   Begin VB.CommandButton CmdCopiar 
      Caption         =   "Copiar Perfil"
      Height          =   975
      Left            =   12360
      Picture         =   "FrmHistorial.frx":4809
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc AdoNomina 
      Height          =   375
      Left            =   360
      Top             =   8880
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
   Begin VB.CommandButton CmdDesactivar 
      Caption         =   "Desactivar"
      Height          =   1095
      Left            =   12360
      Picture         =   "FrmHistorial.frx":4FE4
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CmdActivar 
      Caption         =   "Activar"
      Height          =   1095
      Left            =   12360
      Picture         =   "FrmHistorial.frx":6B26
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CmdDuplicado 
      Caption         =   "DUPLICADOS"
      DownPicture     =   "FrmHistorial.frx":8668
      Height          =   375
      Left            =   9240
      Picture         =   "FrmHistorial.frx":A14A
      TabIndex        =   33
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TRASLADAR"
      DownPicture     =   "FrmHistorial.frx":BC2C
      Height          =   375
      Left            =   10920
      Picture         =   "FrmHistorial.frx":D70E
      TabIndex        =   32
      Top             =   7560
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   255
      Left            =   0
      OleObjectBlob   =   "FrmHistorial.frx":F1F0
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc AdoBusca 
      Height          =   375
      Left            =   5640
      Top             =   9120
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
   Begin VB.TextBox TxtCodEmpleado 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   9360
      TabIndex        =   27
      Top             =   1080
      Width           =   975
   End
   Begin MSAdodcLib.Adodc AdoHistorial 
      Height          =   375
      Left            =   8280
      Top             =   9240
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
      Caption         =   "AdoHistorial"
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
      Height          =   2175
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   10215
      Begin VB.TextBox TxtSexo 
         Height          =   285
         Left            =   6840
         TabIndex        =   38
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   5520
         ScaleHeight     =   1035
         ScaleWidth      =   1035
         TabIndex        =   30
         Top             =   720
         Width           =   1095
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   1020
            Left            =   0
            Picture         =   "FrmHistorial.frx":F26E
            Top             =   0
            Width           =   1020
         End
      End
      Begin TrueOleDBList80.TDBCombo TDBCombo1 
         Bindings        =   "FrmHistorial.frx":122B0
         Height          =   315
         Left            =   1680
         TabIndex        =   29
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
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
         _PropDict       =   $"FrmHistorial.frx":122CB
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmHistorial.frx":12375
         TabIndex        =   22
         Top             =   1680
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmHistorial.frx":123F3
         TabIndex        =   21
         Top             =   1320
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmHistorial.frx":1246F
         TabIndex        =   20
         Top             =   960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmHistorial.frx":124E9
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmHistorial.frx":12561
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtNombre2 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   5
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox TxtApellido1 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox TxtApellido2 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox TxtNombre1 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   2
         Top             =   600
         Width           =   3615
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
         Left            =   9720
         Picture         =   "FrmHistorial.frx":125CB
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   435
         Left            =   6720
         TabIndex        =   31
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   767
         _Version        =   196610
         Font3D          =   2
         MarqueeStyle    =   4
         ForeColor       =   192
         MarqueeDelay    =   5
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   4
         AutoRepeat      =   -1  'True
      End
   End
   Begin MSAdodcLib.Adodc DtaPagos 
      Height          =   375
      Left            =   840
      Top             =   9480
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
      Caption         =   "DtaPagos"
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
   Begin MSAdodcLib.Adodc DtaEmpleados 
      Height          =   375
      Left            =   0
      Top             =   8400
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "DtaEmpleados"
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
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "ACEPTAR"
      DownPicture     =   "FrmHistorial.frx":12719
      Height          =   375
      Left            =   12600
      Picture         =   "FrmHistorial.frx":141FB
      TabIndex        =   15
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   8040
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   3975
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmHistorial.frx":15CDD
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmHistorial.frx":15D45
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtCargo 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox TxtDepartamento 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1335
      Left            =   8160
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   3975
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmHistorial.frx":15DAB
         TabIndex        =   26
         Top             =   720
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmHistorial.frx":15E0F
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox CmdAnnoFin 
         Height          =   315
         ItemData        =   "FrmHistorial.frx":15E79
         Left            =   2760
         List            =   "FrmHistorial.frx":15EC2
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox CmdMesFin 
         Height          =   315
         ItemData        =   "FrmHistorial.frx":15F50
         Left            =   720
         List            =   "FrmHistorial.frx":15F78
         TabIndex        =   13
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox CmdAnnoIni 
         Height          =   315
         ItemData        =   "FrmHistorial.frx":15FE1
         Left            =   2760
         List            =   "FrmHistorial.frx":1602A
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox CmbMesIni 
         Height          =   315
         ItemData        =   "FrmHistorial.frx":160B8
         Left            =   720
         List            =   "FrmHistorial.frx":160E0
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4815
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   13935
      Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
         Bindings        =   "FrmHistorial.frx":16149
         Height          =   4455
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   7858
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "MES"
         Columns(0).DataField=   "MES"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "AÑO"
         Columns(1).DataField=   "AÑO"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Salario Basico"
         Columns(2).DataField=   "SalarioBasico"
         Columns(2).NumberFormat=   "##,##0.00"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Produccion"
         Columns(3).DataField=   "Destajo"
         Columns(3).NumberFormat=   "##,##0.00"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Incentivos"
         Columns(4).DataField=   "Incentivos"
         Columns(4).NumberFormat=   "##,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Numero HE"
         Columns(5).DataField=   "HE"
         Columns(5).NumberFormat=   "##,##0.00"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Horas Extras"
         Columns(6).DataField=   "HorasExtras"
         Columns(6).NumberFormat=   "##,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Puntualidad"
         Columns(7).DataField=   "Comisiones"
         Columns(7).NumberFormat=   "##,##0.00"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Otros Ingresos"
         Columns(8).DataField=   "OtrosIngresos"
         Columns(8).NumberFormat=   "##,##0.00"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Total Ingresos"
         Columns(9).DataField=   "TotalIngresos"
         Columns(9).NumberFormat=   "##,##0.00"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Deducciones"
         Columns(10).DataField=   "Deducciones"
         Columns(10).NumberFormat=   "##,##0.00"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Prestamo"
         Columns(11).DataField=   "Prestamo"
         Columns(11).NumberFormat=   "##,##0.00"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "MontoInss"
         Columns(12).DataField=   "MontoInss"
         Columns(12).NumberFormat=   "##,##0.00"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "MontoIR"
         Columns(13).DataField=   "MontoIR"
         Columns(13).NumberFormat=   "##,##0.00"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Total Egresos"
         Columns(14).DataField=   "TotalEgresos"
         Columns(14).NumberFormat=   "##,##0.00"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "Neto Pagar"
         Columns(15).DataField=   "NetoPagar"
         Columns(15).NumberFormat=   "##,##0.00"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "INSS PATRONAL"
         Columns(16).DataField=   "INSSPATRONAL"
         Columns(16).NumberFormat=   "##,##0.00"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "IR PATRONAL"
         Columns(17).DataField=   "IRPATRONAL"
         Columns(17).NumberFormat=   "##,##0.00"
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   0
         Columns(18)._MaxComboItems=   5
         Columns(18).Caption=   "INATEC"
         Columns(18).DataField=   "INATEC"
         Columns(18).NumberFormat=   "##,##0.00"
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(19)._VlistStyle=   0
         Columns(19)._MaxComboItems=   5
         Columns(19).DataField=   ""
         Columns(19).NumberFormat=   "##,##0.00"
         Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(20)._VlistStyle=   0
         Columns(20)._MaxComboItems=   5
         Columns(20).Caption=   "TARIFA"
         Columns(20).DataField=   "TARIFA"
         Columns(20).NumberFormat=   "##,##0.00"
         Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   21
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).Caption=   "PERIODOS E INGRESOS"
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=21"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(10).Visible=0"
         Splits(0)._ColumnProps(45)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(46)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(47)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(49)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(50)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(51)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(52)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(53)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(54)=   "Column(12).Visible=0"
         Splits(0)._ColumnProps(55)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(56)=   "Column(13).Width=2725"
         Splits(0)._ColumnProps(57)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(58)=   "Column(13)._WidthInPix=2646"
         Splits(0)._ColumnProps(59)=   "Column(13).Visible=0"
         Splits(0)._ColumnProps(60)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(61)=   "Column(14).Width=2725"
         Splits(0)._ColumnProps(62)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(14)._WidthInPix=2646"
         Splits(0)._ColumnProps(64)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(65)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(66)=   "Column(15).Width=2725"
         Splits(0)._ColumnProps(67)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(68)=   "Column(15)._WidthInPix=2646"
         Splits(0)._ColumnProps(69)=   "Column(15).Visible=0"
         Splits(0)._ColumnProps(70)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(71)=   "Column(16).Width=2725"
         Splits(0)._ColumnProps(72)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(73)=   "Column(16)._WidthInPix=2646"
         Splits(0)._ColumnProps(74)=   "Column(16).Visible=0"
         Splits(0)._ColumnProps(75)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(76)=   "Column(17).Width=2725"
         Splits(0)._ColumnProps(77)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(78)=   "Column(17)._WidthInPix=2646"
         Splits(0)._ColumnProps(79)=   "Column(17).Visible=0"
         Splits(0)._ColumnProps(80)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(81)=   "Column(18).Width=2725"
         Splits(0)._ColumnProps(82)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(83)=   "Column(18)._WidthInPix=2646"
         Splits(0)._ColumnProps(84)=   "Column(18).Visible=0"
         Splits(0)._ColumnProps(85)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(86)=   "Column(19).Width=2725"
         Splits(0)._ColumnProps(87)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(88)=   "Column(19)._WidthInPix=2646"
         Splits(0)._ColumnProps(89)=   "Column(19).Visible=0"
         Splits(0)._ColumnProps(90)=   "Column(19).Order=20"
         Splits(0)._ColumnProps(91)=   "Column(20).Width=2725"
         Splits(0)._ColumnProps(92)=   "Column(20).DividerColor=0"
         Splits(0)._ColumnProps(93)=   "Column(20)._WidthInPix=2646"
         Splits(0)._ColumnProps(94)=   "Column(20).Visible=0"
         Splits(0)._ColumnProps(95)=   "Column(20).Order=21"
         Splits(1)._UserFlags=   0
         Splits(1).RecordSelectorWidth=   688
         Splits(1)._SavedRecordSelectors=   -1  'True
         Splits(1).Caption=   "DEDUCCIONES Y NETO PAGAR"
         Splits(1).DividerColor=   14215660
         Splits(1).SpringMode=   0   'False
         Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(1)._ColumnProps(0)=   "Columns.Count=21"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(1)._ColumnProps(4)=   "Column(0).Visible=0"
         Splits(1)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(1)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(1)._ColumnProps(9)=   "Column(1).Visible=0"
         Splits(1)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(1)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(1)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(1)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(1)._ColumnProps(14)=   "Column(2).Visible=0"
         Splits(1)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(1)._ColumnProps(16)=   "Column(3).Width=2725"
         Splits(1)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(1)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
         Splits(1)._ColumnProps(19)=   "Column(3).Visible=0"
         Splits(1)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(1)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(1)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(1)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(1)._ColumnProps(24)=   "Column(4).Visible=0"
         Splits(1)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(1)._ColumnProps(26)=   "Column(5).Width=2725"
         Splits(1)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(1)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
         Splits(1)._ColumnProps(29)=   "Column(5).Visible=0"
         Splits(1)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(1)._ColumnProps(31)=   "Column(6).Width=2725"
         Splits(1)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(1)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
         Splits(1)._ColumnProps(34)=   "Column(6).Visible=0"
         Splits(1)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(1)._ColumnProps(36)=   "Column(7).Width=2725"
         Splits(1)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(1)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
         Splits(1)._ColumnProps(39)=   "Column(7).Visible=0"
         Splits(1)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(1)._ColumnProps(41)=   "Column(8).Width=2725"
         Splits(1)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(1)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
         Splits(1)._ColumnProps(44)=   "Column(8).Visible=0"
         Splits(1)._ColumnProps(45)=   "Column(8).Order=9"
         Splits(1)._ColumnProps(46)=   "Column(9).Width=2725"
         Splits(1)._ColumnProps(47)=   "Column(9).DividerColor=0"
         Splits(1)._ColumnProps(48)=   "Column(9)._WidthInPix=2646"
         Splits(1)._ColumnProps(49)=   "Column(9).Visible=0"
         Splits(1)._ColumnProps(50)=   "Column(9).Order=10"
         Splits(1)._ColumnProps(51)=   "Column(10).Width=2725"
         Splits(1)._ColumnProps(52)=   "Column(10).DividerColor=0"
         Splits(1)._ColumnProps(53)=   "Column(10)._WidthInPix=2646"
         Splits(1)._ColumnProps(54)=   "Column(10).Order=11"
         Splits(1)._ColumnProps(55)=   "Column(11).Width=2725"
         Splits(1)._ColumnProps(56)=   "Column(11).DividerColor=0"
         Splits(1)._ColumnProps(57)=   "Column(11)._WidthInPix=2646"
         Splits(1)._ColumnProps(58)=   "Column(11).Order=12"
         Splits(1)._ColumnProps(59)=   "Column(12).Width=2725"
         Splits(1)._ColumnProps(60)=   "Column(12).DividerColor=0"
         Splits(1)._ColumnProps(61)=   "Column(12)._WidthInPix=2646"
         Splits(1)._ColumnProps(62)=   "Column(12).Order=13"
         Splits(1)._ColumnProps(63)=   "Column(13).Width=2725"
         Splits(1)._ColumnProps(64)=   "Column(13).DividerColor=0"
         Splits(1)._ColumnProps(65)=   "Column(13)._WidthInPix=2646"
         Splits(1)._ColumnProps(66)=   "Column(13).Order=14"
         Splits(1)._ColumnProps(67)=   "Column(14).Width=2725"
         Splits(1)._ColumnProps(68)=   "Column(14).DividerColor=0"
         Splits(1)._ColumnProps(69)=   "Column(14)._WidthInPix=2646"
         Splits(1)._ColumnProps(70)=   "Column(14).Order=15"
         Splits(1)._ColumnProps(71)=   "Column(15).Width=2725"
         Splits(1)._ColumnProps(72)=   "Column(15).DividerColor=0"
         Splits(1)._ColumnProps(73)=   "Column(15)._WidthInPix=2646"
         Splits(1)._ColumnProps(74)=   "Column(15).Order=16"
         Splits(1)._ColumnProps(75)=   "Column(16).Width=2725"
         Splits(1)._ColumnProps(76)=   "Column(16).DividerColor=0"
         Splits(1)._ColumnProps(77)=   "Column(16)._WidthInPix=2646"
         Splits(1)._ColumnProps(78)=   "Column(16).Visible=0"
         Splits(1)._ColumnProps(79)=   "Column(16).Order=17"
         Splits(1)._ColumnProps(80)=   "Column(17).Width=2725"
         Splits(1)._ColumnProps(81)=   "Column(17).DividerColor=0"
         Splits(1)._ColumnProps(82)=   "Column(17)._WidthInPix=2646"
         Splits(1)._ColumnProps(83)=   "Column(17).Visible=0"
         Splits(1)._ColumnProps(84)=   "Column(17).Order=18"
         Splits(1)._ColumnProps(85)=   "Column(18).Width=2725"
         Splits(1)._ColumnProps(86)=   "Column(18).DividerColor=0"
         Splits(1)._ColumnProps(87)=   "Column(18)._WidthInPix=2646"
         Splits(1)._ColumnProps(88)=   "Column(18).Visible=0"
         Splits(1)._ColumnProps(89)=   "Column(18).Order=19"
         Splits(1)._ColumnProps(90)=   "Column(19).Width=2725"
         Splits(1)._ColumnProps(91)=   "Column(19).DividerColor=0"
         Splits(1)._ColumnProps(92)=   "Column(19)._WidthInPix=2646"
         Splits(1)._ColumnProps(93)=   "Column(19).Visible=0"
         Splits(1)._ColumnProps(94)=   "Column(19).Order=20"
         Splits(1)._ColumnProps(95)=   "Column(20).Width=2725"
         Splits(1)._ColumnProps(96)=   "Column(20).DividerColor=0"
         Splits(1)._ColumnProps(97)=   "Column(20)._WidthInPix=2646"
         Splits(1)._ColumnProps(98)=   "Column(20).Visible=0"
         Splits(1)._ColumnProps(99)=   "Column(20).Order=21"
         Splits(2)._UserFlags=   0
         Splits(2).RecordSelectorWidth=   688
         Splits(2)._SavedRecordSelectors=   -1  'True
         Splits(2).Caption=   "DATOS GENERALES"
         Splits(2).DividerColor=   14215660
         Splits(2).SpringMode=   0   'False
         Splits(2)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(2)._ColumnProps(0)=   "Columns.Count=21"
         Splits(2)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(2)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(2)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(2)._ColumnProps(4)=   "Column(0).Visible=0"
         Splits(2)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(2)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(2)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(2)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(2)._ColumnProps(9)=   "Column(1).Visible=0"
         Splits(2)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(2)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(2)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(2)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(2)._ColumnProps(14)=   "Column(2).Visible=0"
         Splits(2)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(2)._ColumnProps(16)=   "Column(3).Width=2725"
         Splits(2)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(2)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
         Splits(2)._ColumnProps(19)=   "Column(3).Visible=0"
         Splits(2)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(2)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(2)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(2)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(2)._ColumnProps(24)=   "Column(4).Visible=0"
         Splits(2)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(2)._ColumnProps(26)=   "Column(5).Width=2725"
         Splits(2)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(2)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
         Splits(2)._ColumnProps(29)=   "Column(5).Visible=0"
         Splits(2)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(2)._ColumnProps(31)=   "Column(6).Width=2725"
         Splits(2)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(2)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
         Splits(2)._ColumnProps(34)=   "Column(6).Visible=0"
         Splits(2)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(2)._ColumnProps(36)=   "Column(7).Width=2725"
         Splits(2)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(2)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
         Splits(2)._ColumnProps(39)=   "Column(7).Visible=0"
         Splits(2)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(2)._ColumnProps(41)=   "Column(8).Width=2725"
         Splits(2)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(2)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
         Splits(2)._ColumnProps(44)=   "Column(8).Visible=0"
         Splits(2)._ColumnProps(45)=   "Column(8).Order=9"
         Splits(2)._ColumnProps(46)=   "Column(9).Width=2725"
         Splits(2)._ColumnProps(47)=   "Column(9).DividerColor=0"
         Splits(2)._ColumnProps(48)=   "Column(9)._WidthInPix=2646"
         Splits(2)._ColumnProps(49)=   "Column(9).Visible=0"
         Splits(2)._ColumnProps(50)=   "Column(9).Order=10"
         Splits(2)._ColumnProps(51)=   "Column(10).Width=2725"
         Splits(2)._ColumnProps(52)=   "Column(10).DividerColor=0"
         Splits(2)._ColumnProps(53)=   "Column(10)._WidthInPix=2646"
         Splits(2)._ColumnProps(54)=   "Column(10).Visible=0"
         Splits(2)._ColumnProps(55)=   "Column(10).Order=11"
         Splits(2)._ColumnProps(56)=   "Column(11).Width=2725"
         Splits(2)._ColumnProps(57)=   "Column(11).DividerColor=0"
         Splits(2)._ColumnProps(58)=   "Column(11)._WidthInPix=2646"
         Splits(2)._ColumnProps(59)=   "Column(11).Visible=0"
         Splits(2)._ColumnProps(60)=   "Column(11).Order=12"
         Splits(2)._ColumnProps(61)=   "Column(12).Width=2725"
         Splits(2)._ColumnProps(62)=   "Column(12).DividerColor=0"
         Splits(2)._ColumnProps(63)=   "Column(12)._WidthInPix=2646"
         Splits(2)._ColumnProps(64)=   "Column(12).Visible=0"
         Splits(2)._ColumnProps(65)=   "Column(12).Order=13"
         Splits(2)._ColumnProps(66)=   "Column(13).Width=2725"
         Splits(2)._ColumnProps(67)=   "Column(13).DividerColor=0"
         Splits(2)._ColumnProps(68)=   "Column(13)._WidthInPix=2646"
         Splits(2)._ColumnProps(69)=   "Column(13).Visible=0"
         Splits(2)._ColumnProps(70)=   "Column(13).Order=14"
         Splits(2)._ColumnProps(71)=   "Column(14).Width=2725"
         Splits(2)._ColumnProps(72)=   "Column(14).DividerColor=0"
         Splits(2)._ColumnProps(73)=   "Column(14)._WidthInPix=2646"
         Splits(2)._ColumnProps(74)=   "Column(14).Visible=0"
         Splits(2)._ColumnProps(75)=   "Column(14).Order=15"
         Splits(2)._ColumnProps(76)=   "Column(15).Width=2725"
         Splits(2)._ColumnProps(77)=   "Column(15).DividerColor=0"
         Splits(2)._ColumnProps(78)=   "Column(15)._WidthInPix=2646"
         Splits(2)._ColumnProps(79)=   "Column(15).Visible=0"
         Splits(2)._ColumnProps(80)=   "Column(15).Order=16"
         Splits(2)._ColumnProps(81)=   "Column(16).Width=2725"
         Splits(2)._ColumnProps(82)=   "Column(16).DividerColor=0"
         Splits(2)._ColumnProps(83)=   "Column(16)._WidthInPix=2646"
         Splits(2)._ColumnProps(84)=   "Column(16).Order=17"
         Splits(2)._ColumnProps(85)=   "Column(17).Width=2725"
         Splits(2)._ColumnProps(86)=   "Column(17).DividerColor=0"
         Splits(2)._ColumnProps(87)=   "Column(17)._WidthInPix=2646"
         Splits(2)._ColumnProps(88)=   "Column(17).Order=18"
         Splits(2)._ColumnProps(89)=   "Column(18).Width=2725"
         Splits(2)._ColumnProps(90)=   "Column(18).DividerColor=0"
         Splits(2)._ColumnProps(91)=   "Column(18)._WidthInPix=2646"
         Splits(2)._ColumnProps(92)=   "Column(18).Order=19"
         Splits(2)._ColumnProps(93)=   "Column(19).Width=2725"
         Splits(2)._ColumnProps(94)=   "Column(19).DividerColor=0"
         Splits(2)._ColumnProps(95)=   "Column(19)._WidthInPix=2646"
         Splits(2)._ColumnProps(96)=   "Column(19).Visible=0"
         Splits(2)._ColumnProps(97)=   "Column(19).Order=20"
         Splits(2)._ColumnProps(98)=   "Column(20).Width=2725"
         Splits(2)._ColumnProps(99)=   "Column(20).DividerColor=0"
         Splits(2)._ColumnProps(100)=   "Column(20)._WidthInPix=2646"
         Splits(2)._ColumnProps(101)=   "Column(20).Order=21"
         Splits.Count    =   3
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   3
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
         _StyleDefs(18)  =   "Splits(0).Style:id=223,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=232,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=224,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=225,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=226,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=228,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=227,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=229,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=230,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=231,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=233,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=234,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=238,.parent=223"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=235,.parent=224"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=236,.parent=225"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=237,.parent=227"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=242,.parent=223"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=239,.parent=224"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=240,.parent=225"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=241,.parent=227"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=246,.parent=223"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=243,.parent=224"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=244,.parent=225"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=245,.parent=227"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=250,.parent=223"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=247,.parent=224"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=248,.parent=225"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=249,.parent=227"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=254,.parent=223"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=251,.parent=224"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=252,.parent=225"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=253,.parent=227"
         _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=258,.parent=223"
         _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=255,.parent=224"
         _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=256,.parent=225"
         _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=257,.parent=227"
         _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=262,.parent=223"
         _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=259,.parent=224"
         _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=260,.parent=225"
         _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=261,.parent=227"
         _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=266,.parent=223"
         _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=263,.parent=224"
         _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=264,.parent=225"
         _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=265,.parent=227"
         _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=270,.parent=223"
         _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=267,.parent=224"
         _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=268,.parent=225"
         _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=269,.parent=227"
         _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=274,.parent=223"
         _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=271,.parent=224"
         _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=272,.parent=225"
         _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=273,.parent=227"
         _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=278,.parent=223"
         _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=275,.parent=224"
         _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=276,.parent=225"
         _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=277,.parent=227"
         _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=282,.parent=223"
         _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=279,.parent=224"
         _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=280,.parent=225"
         _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=281,.parent=227"
         _StyleDefs(78)  =   "Splits(0).Columns(12).Style:id=286,.parent=223"
         _StyleDefs(79)  =   "Splits(0).Columns(12).HeadingStyle:id=283,.parent=224"
         _StyleDefs(80)  =   "Splits(0).Columns(12).FooterStyle:id=284,.parent=225"
         _StyleDefs(81)  =   "Splits(0).Columns(12).EditorStyle:id=285,.parent=227"
         _StyleDefs(82)  =   "Splits(0).Columns(13).Style:id=290,.parent=223"
         _StyleDefs(83)  =   "Splits(0).Columns(13).HeadingStyle:id=287,.parent=224"
         _StyleDefs(84)  =   "Splits(0).Columns(13).FooterStyle:id=288,.parent=225"
         _StyleDefs(85)  =   "Splits(0).Columns(13).EditorStyle:id=289,.parent=227"
         _StyleDefs(86)  =   "Splits(0).Columns(14).Style:id=294,.parent=223"
         _StyleDefs(87)  =   "Splits(0).Columns(14).HeadingStyle:id=291,.parent=224"
         _StyleDefs(88)  =   "Splits(0).Columns(14).FooterStyle:id=292,.parent=225"
         _StyleDefs(89)  =   "Splits(0).Columns(14).EditorStyle:id=293,.parent=227"
         _StyleDefs(90)  =   "Splits(0).Columns(15).Style:id=298,.parent=223"
         _StyleDefs(91)  =   "Splits(0).Columns(15).HeadingStyle:id=295,.parent=224"
         _StyleDefs(92)  =   "Splits(0).Columns(15).FooterStyle:id=296,.parent=225"
         _StyleDefs(93)  =   "Splits(0).Columns(15).EditorStyle:id=297,.parent=227"
         _StyleDefs(94)  =   "Splits(0).Columns(16).Style:id=302,.parent=223"
         _StyleDefs(95)  =   "Splits(0).Columns(16).HeadingStyle:id=299,.parent=224"
         _StyleDefs(96)  =   "Splits(0).Columns(16).FooterStyle:id=300,.parent=225"
         _StyleDefs(97)  =   "Splits(0).Columns(16).EditorStyle:id=301,.parent=227"
         _StyleDefs(98)  =   "Splits(0).Columns(17).Style:id=306,.parent=223"
         _StyleDefs(99)  =   "Splits(0).Columns(17).HeadingStyle:id=303,.parent=224"
         _StyleDefs(100) =   "Splits(0).Columns(17).FooterStyle:id=304,.parent=225"
         _StyleDefs(101) =   "Splits(0).Columns(17).EditorStyle:id=305,.parent=227"
         _StyleDefs(102) =   "Splits(0).Columns(18).Style:id=310,.parent=223"
         _StyleDefs(103) =   "Splits(0).Columns(18).HeadingStyle:id=307,.parent=224"
         _StyleDefs(104) =   "Splits(0).Columns(18).FooterStyle:id=308,.parent=225"
         _StyleDefs(105) =   "Splits(0).Columns(18).EditorStyle:id=309,.parent=227"
         _StyleDefs(106) =   "Splits(0).Columns(19).Style:id=314,.parent=223"
         _StyleDefs(107) =   "Splits(0).Columns(19).HeadingStyle:id=311,.parent=224"
         _StyleDefs(108) =   "Splits(0).Columns(19).FooterStyle:id=312,.parent=225"
         _StyleDefs(109) =   "Splits(0).Columns(19).EditorStyle:id=313,.parent=227"
         _StyleDefs(110) =   "Splits(0).Columns(20).Style:id=318,.parent=223"
         _StyleDefs(111) =   "Splits(0).Columns(20).HeadingStyle:id=315,.parent=224"
         _StyleDefs(112) =   "Splits(0).Columns(20).FooterStyle:id=316,.parent=225"
         _StyleDefs(113) =   "Splits(0).Columns(20).EditorStyle:id=317,.parent=227"
         _StyleDefs(114) =   "Splits(1).Style:id=123,.parent=1"
         _StyleDefs(115) =   "Splits(1).CaptionStyle:id=132,.parent=4"
         _StyleDefs(116) =   "Splits(1).HeadingStyle:id=124,.parent=2"
         _StyleDefs(117) =   "Splits(1).FooterStyle:id=125,.parent=3"
         _StyleDefs(118) =   "Splits(1).InactiveStyle:id=126,.parent=5"
         _StyleDefs(119) =   "Splits(1).SelectedStyle:id=128,.parent=6"
         _StyleDefs(120) =   "Splits(1).EditorStyle:id=127,.parent=7"
         _StyleDefs(121) =   "Splits(1).HighlightRowStyle:id=129,.parent=8"
         _StyleDefs(122) =   "Splits(1).EvenRowStyle:id=130,.parent=9"
         _StyleDefs(123) =   "Splits(1).OddRowStyle:id=131,.parent=10"
         _StyleDefs(124) =   "Splits(1).RecordSelectorStyle:id=133,.parent=11"
         _StyleDefs(125) =   "Splits(1).FilterBarStyle:id=134,.parent=12"
         _StyleDefs(126) =   "Splits(1).Columns(0).Style:id=138,.parent=123"
         _StyleDefs(127) =   "Splits(1).Columns(0).HeadingStyle:id=135,.parent=124"
         _StyleDefs(128) =   "Splits(1).Columns(0).FooterStyle:id=136,.parent=125"
         _StyleDefs(129) =   "Splits(1).Columns(0).EditorStyle:id=137,.parent=127"
         _StyleDefs(130) =   "Splits(1).Columns(1).Style:id=142,.parent=123"
         _StyleDefs(131) =   "Splits(1).Columns(1).HeadingStyle:id=139,.parent=124"
         _StyleDefs(132) =   "Splits(1).Columns(1).FooterStyle:id=140,.parent=125"
         _StyleDefs(133) =   "Splits(1).Columns(1).EditorStyle:id=141,.parent=127"
         _StyleDefs(134) =   "Splits(1).Columns(2).Style:id=146,.parent=123"
         _StyleDefs(135) =   "Splits(1).Columns(2).HeadingStyle:id=143,.parent=124"
         _StyleDefs(136) =   "Splits(1).Columns(2).FooterStyle:id=144,.parent=125"
         _StyleDefs(137) =   "Splits(1).Columns(2).EditorStyle:id=145,.parent=127"
         _StyleDefs(138) =   "Splits(1).Columns(3).Style:id=150,.parent=123"
         _StyleDefs(139) =   "Splits(1).Columns(3).HeadingStyle:id=147,.parent=124"
         _StyleDefs(140) =   "Splits(1).Columns(3).FooterStyle:id=148,.parent=125"
         _StyleDefs(141) =   "Splits(1).Columns(3).EditorStyle:id=149,.parent=127"
         _StyleDefs(142) =   "Splits(1).Columns(4).Style:id=154,.parent=123"
         _StyleDefs(143) =   "Splits(1).Columns(4).HeadingStyle:id=151,.parent=124"
         _StyleDefs(144) =   "Splits(1).Columns(4).FooterStyle:id=152,.parent=125"
         _StyleDefs(145) =   "Splits(1).Columns(4).EditorStyle:id=153,.parent=127"
         _StyleDefs(146) =   "Splits(1).Columns(5).Style:id=158,.parent=123"
         _StyleDefs(147) =   "Splits(1).Columns(5).HeadingStyle:id=155,.parent=124"
         _StyleDefs(148) =   "Splits(1).Columns(5).FooterStyle:id=156,.parent=125"
         _StyleDefs(149) =   "Splits(1).Columns(5).EditorStyle:id=157,.parent=127"
         _StyleDefs(150) =   "Splits(1).Columns(6).Style:id=162,.parent=123"
         _StyleDefs(151) =   "Splits(1).Columns(6).HeadingStyle:id=159,.parent=124"
         _StyleDefs(152) =   "Splits(1).Columns(6).FooterStyle:id=160,.parent=125"
         _StyleDefs(153) =   "Splits(1).Columns(6).EditorStyle:id=161,.parent=127"
         _StyleDefs(154) =   "Splits(1).Columns(7).Style:id=166,.parent=123"
         _StyleDefs(155) =   "Splits(1).Columns(7).HeadingStyle:id=163,.parent=124"
         _StyleDefs(156) =   "Splits(1).Columns(7).FooterStyle:id=164,.parent=125"
         _StyleDefs(157) =   "Splits(1).Columns(7).EditorStyle:id=165,.parent=127"
         _StyleDefs(158) =   "Splits(1).Columns(8).Style:id=170,.parent=123"
         _StyleDefs(159) =   "Splits(1).Columns(8).HeadingStyle:id=167,.parent=124"
         _StyleDefs(160) =   "Splits(1).Columns(8).FooterStyle:id=168,.parent=125"
         _StyleDefs(161) =   "Splits(1).Columns(8).EditorStyle:id=169,.parent=127"
         _StyleDefs(162) =   "Splits(1).Columns(9).Style:id=174,.parent=123"
         _StyleDefs(163) =   "Splits(1).Columns(9).HeadingStyle:id=171,.parent=124"
         _StyleDefs(164) =   "Splits(1).Columns(9).FooterStyle:id=172,.parent=125"
         _StyleDefs(165) =   "Splits(1).Columns(9).EditorStyle:id=173,.parent=127"
         _StyleDefs(166) =   "Splits(1).Columns(10).Style:id=178,.parent=123"
         _StyleDefs(167) =   "Splits(1).Columns(10).HeadingStyle:id=175,.parent=124"
         _StyleDefs(168) =   "Splits(1).Columns(10).FooterStyle:id=176,.parent=125"
         _StyleDefs(169) =   "Splits(1).Columns(10).EditorStyle:id=177,.parent=127"
         _StyleDefs(170) =   "Splits(1).Columns(11).Style:id=182,.parent=123"
         _StyleDefs(171) =   "Splits(1).Columns(11).HeadingStyle:id=179,.parent=124"
         _StyleDefs(172) =   "Splits(1).Columns(11).FooterStyle:id=180,.parent=125"
         _StyleDefs(173) =   "Splits(1).Columns(11).EditorStyle:id=181,.parent=127"
         _StyleDefs(174) =   "Splits(1).Columns(12).Style:id=186,.parent=123"
         _StyleDefs(175) =   "Splits(1).Columns(12).HeadingStyle:id=183,.parent=124"
         _StyleDefs(176) =   "Splits(1).Columns(12).FooterStyle:id=184,.parent=125"
         _StyleDefs(177) =   "Splits(1).Columns(12).EditorStyle:id=185,.parent=127"
         _StyleDefs(178) =   "Splits(1).Columns(13).Style:id=190,.parent=123"
         _StyleDefs(179) =   "Splits(1).Columns(13).HeadingStyle:id=187,.parent=124"
         _StyleDefs(180) =   "Splits(1).Columns(13).FooterStyle:id=188,.parent=125"
         _StyleDefs(181) =   "Splits(1).Columns(13).EditorStyle:id=189,.parent=127"
         _StyleDefs(182) =   "Splits(1).Columns(14).Style:id=194,.parent=123"
         _StyleDefs(183) =   "Splits(1).Columns(14).HeadingStyle:id=191,.parent=124"
         _StyleDefs(184) =   "Splits(1).Columns(14).FooterStyle:id=192,.parent=125"
         _StyleDefs(185) =   "Splits(1).Columns(14).EditorStyle:id=193,.parent=127"
         _StyleDefs(186) =   "Splits(1).Columns(15).Style:id=198,.parent=123"
         _StyleDefs(187) =   "Splits(1).Columns(15).HeadingStyle:id=195,.parent=124"
         _StyleDefs(188) =   "Splits(1).Columns(15).FooterStyle:id=196,.parent=125"
         _StyleDefs(189) =   "Splits(1).Columns(15).EditorStyle:id=197,.parent=127"
         _StyleDefs(190) =   "Splits(1).Columns(16).Style:id=202,.parent=123"
         _StyleDefs(191) =   "Splits(1).Columns(16).HeadingStyle:id=199,.parent=124"
         _StyleDefs(192) =   "Splits(1).Columns(16).FooterStyle:id=200,.parent=125"
         _StyleDefs(193) =   "Splits(1).Columns(16).EditorStyle:id=201,.parent=127"
         _StyleDefs(194) =   "Splits(1).Columns(17).Style:id=206,.parent=123"
         _StyleDefs(195) =   "Splits(1).Columns(17).HeadingStyle:id=203,.parent=124"
         _StyleDefs(196) =   "Splits(1).Columns(17).FooterStyle:id=204,.parent=125"
         _StyleDefs(197) =   "Splits(1).Columns(17).EditorStyle:id=205,.parent=127"
         _StyleDefs(198) =   "Splits(1).Columns(18).Style:id=210,.parent=123"
         _StyleDefs(199) =   "Splits(1).Columns(18).HeadingStyle:id=207,.parent=124"
         _StyleDefs(200) =   "Splits(1).Columns(18).FooterStyle:id=208,.parent=125"
         _StyleDefs(201) =   "Splits(1).Columns(18).EditorStyle:id=209,.parent=127"
         _StyleDefs(202) =   "Splits(1).Columns(19).Style:id=214,.parent=123"
         _StyleDefs(203) =   "Splits(1).Columns(19).HeadingStyle:id=211,.parent=124"
         _StyleDefs(204) =   "Splits(1).Columns(19).FooterStyle:id=212,.parent=125"
         _StyleDefs(205) =   "Splits(1).Columns(19).EditorStyle:id=213,.parent=127"
         _StyleDefs(206) =   "Splits(1).Columns(20).Style:id=218,.parent=123"
         _StyleDefs(207) =   "Splits(1).Columns(20).HeadingStyle:id=215,.parent=124"
         _StyleDefs(208) =   "Splits(1).Columns(20).FooterStyle:id=216,.parent=125"
         _StyleDefs(209) =   "Splits(1).Columns(20).EditorStyle:id=217,.parent=127"
         _StyleDefs(210) =   "Splits(2).Style:id=13,.parent=1"
         _StyleDefs(211) =   "Splits(2).CaptionStyle:id=22,.parent=4"
         _StyleDefs(212) =   "Splits(2).HeadingStyle:id=14,.parent=2"
         _StyleDefs(213) =   "Splits(2).FooterStyle:id=15,.parent=3"
         _StyleDefs(214) =   "Splits(2).InactiveStyle:id=16,.parent=5"
         _StyleDefs(215) =   "Splits(2).SelectedStyle:id=18,.parent=6"
         _StyleDefs(216) =   "Splits(2).EditorStyle:id=17,.parent=7"
         _StyleDefs(217) =   "Splits(2).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(218) =   "Splits(2).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(219) =   "Splits(2).OddRowStyle:id=21,.parent=10"
         _StyleDefs(220) =   "Splits(2).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(221) =   "Splits(2).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(222) =   "Splits(2).Columns(0).Style:id=122,.parent=13"
         _StyleDefs(223) =   "Splits(2).Columns(0).HeadingStyle:id=119,.parent=14"
         _StyleDefs(224) =   "Splits(2).Columns(0).FooterStyle:id=120,.parent=15"
         _StyleDefs(225) =   "Splits(2).Columns(0).EditorStyle:id=121,.parent=17"
         _StyleDefs(226) =   "Splits(2).Columns(1).Style:id=118,.parent=13"
         _StyleDefs(227) =   "Splits(2).Columns(1).HeadingStyle:id=115,.parent=14"
         _StyleDefs(228) =   "Splits(2).Columns(1).FooterStyle:id=116,.parent=15"
         _StyleDefs(229) =   "Splits(2).Columns(1).EditorStyle:id=117,.parent=17"
         _StyleDefs(230) =   "Splits(2).Columns(2).Style:id=114,.parent=13"
         _StyleDefs(231) =   "Splits(2).Columns(2).HeadingStyle:id=111,.parent=14"
         _StyleDefs(232) =   "Splits(2).Columns(2).FooterStyle:id=112,.parent=15"
         _StyleDefs(233) =   "Splits(2).Columns(2).EditorStyle:id=113,.parent=17"
         _StyleDefs(234) =   "Splits(2).Columns(3).Style:id=110,.parent=13"
         _StyleDefs(235) =   "Splits(2).Columns(3).HeadingStyle:id=107,.parent=14"
         _StyleDefs(236) =   "Splits(2).Columns(3).FooterStyle:id=108,.parent=15"
         _StyleDefs(237) =   "Splits(2).Columns(3).EditorStyle:id=109,.parent=17"
         _StyleDefs(238) =   "Splits(2).Columns(4).Style:id=106,.parent=13"
         _StyleDefs(239) =   "Splits(2).Columns(4).HeadingStyle:id=103,.parent=14"
         _StyleDefs(240) =   "Splits(2).Columns(4).FooterStyle:id=104,.parent=15"
         _StyleDefs(241) =   "Splits(2).Columns(4).EditorStyle:id=105,.parent=17"
         _StyleDefs(242) =   "Splits(2).Columns(5).Style:id=102,.parent=13"
         _StyleDefs(243) =   "Splits(2).Columns(5).HeadingStyle:id=99,.parent=14"
         _StyleDefs(244) =   "Splits(2).Columns(5).FooterStyle:id=100,.parent=15"
         _StyleDefs(245) =   "Splits(2).Columns(5).EditorStyle:id=101,.parent=17"
         _StyleDefs(246) =   "Splits(2).Columns(6).Style:id=98,.parent=13"
         _StyleDefs(247) =   "Splits(2).Columns(6).HeadingStyle:id=95,.parent=14"
         _StyleDefs(248) =   "Splits(2).Columns(6).FooterStyle:id=96,.parent=15"
         _StyleDefs(249) =   "Splits(2).Columns(6).EditorStyle:id=97,.parent=17"
         _StyleDefs(250) =   "Splits(2).Columns(7).Style:id=94,.parent=13"
         _StyleDefs(251) =   "Splits(2).Columns(7).HeadingStyle:id=91,.parent=14"
         _StyleDefs(252) =   "Splits(2).Columns(7).FooterStyle:id=92,.parent=15"
         _StyleDefs(253) =   "Splits(2).Columns(7).EditorStyle:id=93,.parent=17"
         _StyleDefs(254) =   "Splits(2).Columns(8).Style:id=90,.parent=13"
         _StyleDefs(255) =   "Splits(2).Columns(8).HeadingStyle:id=87,.parent=14"
         _StyleDefs(256) =   "Splits(2).Columns(8).FooterStyle:id=88,.parent=15"
         _StyleDefs(257) =   "Splits(2).Columns(8).EditorStyle:id=89,.parent=17"
         _StyleDefs(258) =   "Splits(2).Columns(9).Style:id=86,.parent=13"
         _StyleDefs(259) =   "Splits(2).Columns(9).HeadingStyle:id=83,.parent=14"
         _StyleDefs(260) =   "Splits(2).Columns(9).FooterStyle:id=84,.parent=15"
         _StyleDefs(261) =   "Splits(2).Columns(9).EditorStyle:id=85,.parent=17"
         _StyleDefs(262) =   "Splits(2).Columns(10).Style:id=82,.parent=13"
         _StyleDefs(263) =   "Splits(2).Columns(10).HeadingStyle:id=79,.parent=14"
         _StyleDefs(264) =   "Splits(2).Columns(10).FooterStyle:id=80,.parent=15"
         _StyleDefs(265) =   "Splits(2).Columns(10).EditorStyle:id=81,.parent=17"
         _StyleDefs(266) =   "Splits(2).Columns(11).Style:id=78,.parent=13"
         _StyleDefs(267) =   "Splits(2).Columns(11).HeadingStyle:id=75,.parent=14"
         _StyleDefs(268) =   "Splits(2).Columns(11).FooterStyle:id=76,.parent=15"
         _StyleDefs(269) =   "Splits(2).Columns(11).EditorStyle:id=77,.parent=17"
         _StyleDefs(270) =   "Splits(2).Columns(12).Style:id=74,.parent=13"
         _StyleDefs(271) =   "Splits(2).Columns(12).HeadingStyle:id=71,.parent=14"
         _StyleDefs(272) =   "Splits(2).Columns(12).FooterStyle:id=72,.parent=15"
         _StyleDefs(273) =   "Splits(2).Columns(12).EditorStyle:id=73,.parent=17"
         _StyleDefs(274) =   "Splits(2).Columns(13).Style:id=70,.parent=13"
         _StyleDefs(275) =   "Splits(2).Columns(13).HeadingStyle:id=67,.parent=14"
         _StyleDefs(276) =   "Splits(2).Columns(13).FooterStyle:id=68,.parent=15"
         _StyleDefs(277) =   "Splits(2).Columns(13).EditorStyle:id=69,.parent=17"
         _StyleDefs(278) =   "Splits(2).Columns(14).Style:id=66,.parent=13"
         _StyleDefs(279) =   "Splits(2).Columns(14).HeadingStyle:id=63,.parent=14"
         _StyleDefs(280) =   "Splits(2).Columns(14).FooterStyle:id=64,.parent=15"
         _StyleDefs(281) =   "Splits(2).Columns(14).EditorStyle:id=65,.parent=17"
         _StyleDefs(282) =   "Splits(2).Columns(15).Style:id=62,.parent=13"
         _StyleDefs(283) =   "Splits(2).Columns(15).HeadingStyle:id=59,.parent=14"
         _StyleDefs(284) =   "Splits(2).Columns(15).FooterStyle:id=60,.parent=15"
         _StyleDefs(285) =   "Splits(2).Columns(15).EditorStyle:id=61,.parent=17"
         _StyleDefs(286) =   "Splits(2).Columns(16).Style:id=58,.parent=13"
         _StyleDefs(287) =   "Splits(2).Columns(16).HeadingStyle:id=55,.parent=14"
         _StyleDefs(288) =   "Splits(2).Columns(16).FooterStyle:id=56,.parent=15"
         _StyleDefs(289) =   "Splits(2).Columns(16).EditorStyle:id=57,.parent=17"
         _StyleDefs(290) =   "Splits(2).Columns(17).Style:id=54,.parent=13"
         _StyleDefs(291) =   "Splits(2).Columns(17).HeadingStyle:id=51,.parent=14"
         _StyleDefs(292) =   "Splits(2).Columns(17).FooterStyle:id=52,.parent=15"
         _StyleDefs(293) =   "Splits(2).Columns(17).EditorStyle:id=53,.parent=17"
         _StyleDefs(294) =   "Splits(2).Columns(18).Style:id=50,.parent=13"
         _StyleDefs(295) =   "Splits(2).Columns(18).HeadingStyle:id=47,.parent=14"
         _StyleDefs(296) =   "Splits(2).Columns(18).FooterStyle:id=48,.parent=15"
         _StyleDefs(297) =   "Splits(2).Columns(18).EditorStyle:id=49,.parent=17"
         _StyleDefs(298) =   "Splits(2).Columns(19).Style:id=46,.parent=13"
         _StyleDefs(299) =   "Splits(2).Columns(19).HeadingStyle:id=43,.parent=14"
         _StyleDefs(300) =   "Splits(2).Columns(19).FooterStyle:id=44,.parent=15"
         _StyleDefs(301) =   "Splits(2).Columns(19).EditorStyle:id=45,.parent=17"
         _StyleDefs(302) =   "Splits(2).Columns(20).Style:id=28,.parent=13"
         _StyleDefs(303) =   "Splits(2).Columns(20).HeadingStyle:id=25,.parent=14"
         _StyleDefs(304) =   "Splits(2).Columns(20).FooterStyle:id=26,.parent=15"
         _StyleDefs(305) =   "Splits(2).Columns(20).EditorStyle:id=27,.parent=17"
         _StyleDefs(306) =   "Named:id=33:Normal"
         _StyleDefs(307) =   ":id=33,.parent=0"
         _StyleDefs(308) =   "Named:id=34:Heading"
         _StyleDefs(309) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(310) =   ":id=34,.wraptext=-1"
         _StyleDefs(311) =   "Named:id=35:Footing"
         _StyleDefs(312) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(313) =   "Named:id=36:Selected"
         _StyleDefs(314) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(315) =   "Named:id=37:Caption"
         _StyleDefs(316) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(317) =   "Named:id=38:HighlightRow"
         _StyleDefs(318) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(319) =   "Named:id=39:EvenRow"
         _StyleDefs(320) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(321) =   "Named:id=40:OddRow"
         _StyleDefs(322) =   ":id=40,.parent=33"
         _StyleDefs(323) =   "Named:id=41:RecordSelector"
         _StyleDefs(324) =   ":id=41,.parent=34"
         _StyleDefs(325) =   "Named:id=42:FilterBar"
         _StyleDefs(326) =   ":id=42,.parent=33"
      End
   End
   Begin Threed.SSCommand CmdAcercade 
      Height          =   435
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   767
      _Version        =   196610
      Font3D          =   2
      MarqueeStyle    =   4
      ForeColor       =   8388608
      MarqueeDelay    =   5
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Historial Salarial x Empleado"
      ButtonStyle     =   4
      AutoRepeat      =   -1  'True
   End
   Begin MSAdodcLib.Adodc DtaEmpleadosNuevos 
      Height          =   375
      Left            =   7320
      Top             =   8760
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Caption         =   "DtaEmpleadosNuevos"
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
   Begin MSAdodcLib.Adodc AdoHistoricos 
      Height          =   375
      Left            =   8160
      Top             =   8280
      Visible         =   0   'False
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
      Caption         =   "AdoHistoricos"
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
   Begin MSAdodcLib.Adodc DtaHorarioEmpleado 
      Height          =   375
      Left            =   4920
      Top             =   8400
      Visible         =   0   'False
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
      Caption         =   "DtaHorarioEmpleado"
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
   Begin MSAdodcLib.Adodc DtaTurnos 
      Height          =   375
      Left            =   1200
      Top             =   8040
      Visible         =   0   'False
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
      Caption         =   "DtaTurnos"
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
Attribute VB_Name = "FrmHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public EmpleadoActivo As Boolean, Sexo As String, Cargo As String, Cedula As String, Basico As Double, FechaIngreso As Date
Private Sub CmdAceptar_Click()
Unload Me
End Sub

Private Sub CmdActivar_Click()
Dim Respuesta As String, CodEmpleados As Double
 Respuesta = MsgBox("Esta Seguro de Activar el Empleado", vbYesNo, "Empleado " & Me.TxtCodEmpleado.Text)
 
 If Respuesta = 6 Then
  CodEmpleados = Me.TxtCodEmpleado.Text
  Me.AdoBusca.RecordSource = "SELECT * From Empleado Where (CodEmpleado = " & CodEmpleados & ")"
  Me.AdoBusca.Refresh
  If Not Me.AdoBusca.Recordset.EOF Then
   If Me.AdoBusca.Recordset("Activo") = False Then
     Me.AdoBusca.Recordset("Activo") = 1
     Me.AdoBusca.Recordset.Update
   End If
  End If
 End If
 
 With Me.DtaEmpleados
   .ConnectionString = Conexion
   .RecordSource = "SELECT Empleado.CodEmpleado1, Empleado.CodEmpleado,Empleado.Nombre1 +' '+ Empleado.Nombre2+' ' + Empleado.Apellido1+' ' + Empleado.Apellido2 AS Nombres, Empleado.Activo FROM Departamento INNER JOIN  Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento ORDER BY Empleado.CodEmpleado1 "
   .Refresh
 End With
 MsgBox "Empleado Activado", vbExclamation
End Sub

Private Sub CmdBuscarEmpleado_Click()
'Quien = "Historial"
'FrmBuscaEmpleado.Show 1

QueProducto = "CodigoProductoHistorico"
FrmConsulta.Show 1
Me.TDBCombo1.Text = FrmConsulta.CodigoEmpleado1
TDBCombo1_ItemChange

End Sub

Private Sub DtaEmpleados_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub DBCodigoEmpleado_KeyPress(KeyAscii As Integer)

End Sub

Private Sub CmdConstanciaActivos_Click()
 If EmpleadoActivo = True Then
   ArepConstanciaActivos.LblLineaUno.Caption = ""
   
 End If
End Sub

Private Sub CmdCopiar_Click()
Dim CodEmpleado As Double, CodEmpleado1 As String, CodEmpleadoAnterior As Double
Dim FechaNacimiento As Date, Id As Double, Contador As Double, Contador2 As Double
Dim Mes As Integer, CodigoNuevo As String, Numero As Double


 
  Respuesta = MsgBox("Desea reutilizar Codigos?", vbYesNo, "Zeus Facturacion")
 
 
 If Respuesta = 6 Then
    CodEmpleado = Me.TxtCodEmpleado.Text
    CodEmpleado1 = Me.TDBCombo1.Text
 Else
    Mes = Month(Now)
    CodigoNuevo = "S1" & Mid(Year(Now), Len(Year(Now)) - 1, Len(Year(Now)) - 2) & Format(Mes, "0#")
    Me.DtaConsulta.RecordSource = "SELECT CodEmpleado1 From Empleado WHERE (CodEmpleado1 LIKE '%" & CodigoNuevo & "%') ORDER BY CodEmpleado1 DESC"
    Me.DtaConsulta.Refresh
    If Not Me.DtaConsulta.Recordset.EOF Then
      Numero = Mid(Me.DtaConsulta.Recordset("CodEmpleado1"), 7, 4)
      CodigoNuevo = CodigoNuevo & Format(Numero + 1, "000#")
    Else
      CodigoNuevo = CodigoNuevo & "0001"
    End If
    
      CodEmpleado = Me.TxtCodEmpleado.Text
      CodEmpleado1 = CodigoNuevo
      
 End If
 
 Me.DtaEmpleados.ConnectionString = Conexion
 Me.DtaEmpleados.RecordSource = "SELECT  Empleado.* From Empleado"
 Me.DtaEmpleados.Refresh
 


 '////////////////////////////////////AGREGO EL REGISTRO DEL EMPLEADO //////////////////////////////////////
 

 Me.AdoBusca.RecordSource = "SELECT * From Empleado Where (CodEmpleado = " & Me.TxtCodEmpleado.Text & ")"
 Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
 
        DtaEmpleados.Recordset.AddNew

            DtaEmpleados.Recordset("CodEmpleado1") = CodEmpleado1
            DtaEmpleados.Recordset("Nombre1") = Me.AdoBusca.Recordset("Nombre1")
            DtaEmpleados.Recordset("Nombre2") = Me.AdoBusca.Recordset("Nombre2")
            DtaEmpleados.Recordset("Apellido1") = Me.AdoBusca.Recordset("Apellido1")
            DtaEmpleados.Recordset("Apellido2") = Me.AdoBusca.Recordset("Apellido2")
            DtaEmpleados.Recordset("Direccion") = Me.AdoBusca.Recordset("Direccion")
            DtaEmpleados.Recordset("Nacionalidad") = Me.AdoBusca.Recordset("Nacionalidad")
            DtaEmpleados.Recordset("CodigoPostal") = Me.AdoBusca.Recordset("CodigoPostal")
            DtaEmpleados.Recordset("numcedula") = Me.AdoBusca.Recordset("numcedula")
            DtaEmpleados.Recordset("sexo") = Me.AdoBusca.Recordset("sexo")
            DtaEmpleados.Recordset("NumeroInss") = Me.AdoBusca.Recordset("NumeroInss")
            DtaEmpleados.Recordset("numeroruc") = Me.AdoBusca.Recordset("numeroruc")
            DtaEmpleados.Recordset("CodDepartamento") = Me.AdoBusca.Recordset("CodDepartamento")
            DtaEmpleados.Recordset("CodCargo") = Me.AdoBusca.Recordset("CodCargo")
            DtaEmpleados.Recordset("Codgrupo") = Me.AdoBusca.Recordset("Codgrupo")
            DtaEmpleados.Recordset("Sindicalista") = Me.AdoBusca.Recordset("Sindicalista")
            DtaEmpleados.Recordset("CodTipoNomina") = Me.AdoBusca.Recordset("CodTipoNomina")
            DtaEmpleados.Recordset("numhijos") = Me.AdoBusca.Recordset("numhijos")
            DtaEmpleados.Recordset("PorcientoIncentivo") = 0
            DtaEmpleados.Recordset("SueldoPeriodo") = Me.AdoBusca.Recordset("SueldoPeriodo")
            DtaEmpleados.Recordset("TarifaHoraria") = Me.AdoBusca.Recordset("TarifaHoraria")
            DtaEmpleados.Recordset("PorcentajeComision") = Me.AdoBusca.Recordset("PorcentajeComision")
            DtaEmpleados.Recordset("OtrosIngresos") = Me.AdoBusca.Recordset("OtrosIngresos")
            DtaEmpleados.Recordset("salariominimo") = Me.AdoBusca.Recordset("salariominimo")
            DtaEmpleados.Recordset("ExentoInss") = Me.AdoBusca.Recordset("ExentoInss")
            DtaEmpleados.Recordset("ExentoIr") = Me.AdoBusca.Recordset("ExentoIr")
            DtaEmpleados.Recordset("PagoInssPatronal") = Me.AdoBusca.Recordset("PagoInssPatronal")
            If Not IsNull(Me.AdoBusca.Recordset("ViaticoxDia")) Then
             DtaEmpleados.Recordset("ViaticoxDia") = Me.AdoBusca.Recordset("ViaticoxDia")
            End If
            If Not IsNull(Me.AdoBusca.Recordset("CuentaBanco")) Then
             DtaEmpleados.Recordset("CuentaBanco") = Me.AdoBusca.Recordset("CuentaBanco")
            End If
            
      DtaEmpleados.Recordset.Update
      
      
            '///////////////////////////////////////////////////////////////////////////////////////////////
      '//////////////////////////BUSCO LA FECHA DE NACIMIENTO DE EMPLEADO DADO DE BAJA /////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////////////
      CodEmpleadoAnterior = CodEmpleado
      Me.DtaConsulta.RecordSource = "SELECT  * From Historico Where (CodEmpleado = " & CodEmpleado & ")"
      Me.DtaConsulta.Refresh
      If Not Me.DtaConsulta.Recordset.EOF Then
        FechaNacimiento = Me.DtaConsulta.Recordset("FechaNacimiento")
      End If

      
      '///////////////////////////////////////////////////////////////////////////////////////////////
      '//////////////////////////BUSCO EL CODIGO INTERNO PARA EL EMPLEADO QUE ACABO DE GRABAR /////////////////////////
      '////////////////////////////////////////////////////////////////////////////////////////////////
      Me.DtaConsulta.RecordSource = "SELECT * From Empleado WHERE (CodEmpleado1 = '" & CodEmpleado1 & "') AND (Activo = 1) ORDER BY CodEmpleado"
      Me.DtaConsulta.Refresh
      If Not Me.DtaConsulta.Recordset.EOF Then
        CodEmpleado = Me.DtaConsulta.Recordset("CodEmpleado")
      End If
      
      
      '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      '/////////////////////////////GRABO LA FECHA DE INGRESO Y VACACIONES ///////////////////////////////////////////
      '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      
       Me.DtaConsulta.RecordSource = "SELECT  * From Historico"
       Me.DtaConsulta.Refresh
       If Not Me.DtaConsulta.Recordset.EOF Then
         Me.DtaConsulta.Recordset.MoveLast
         Id = Me.DtaConsulta.Recordset("id") + 1
       Else
        Id = 1
       End If
      
       Me.AdoHistoricos.RecordSource = "SELECT  * From Historico Where (CodEmpleado = " & CodEmpleado & ")"
       Me.AdoHistoricos.Refresh
       If Me.AdoHistoricos.Recordset.EOF Then
         Me.AdoHistoricos.Recordset.AddNew
          Me.AdoHistoricos.Recordset("id") = Id
          Me.AdoHistoricos.Recordset("Codempleado") = CodEmpleado
          Me.AdoHistoricos.Recordset("FechaNacimiento") = FechaNacimiento
          Me.AdoHistoricos.Recordset("FechaContrato") = Format(Now, "dd/mm/yyyy")
          Me.AdoHistoricos.Recordset("FechaContratoVac") = Format(Now, "dd/mm/yyyy")
         Me.AdoHistoricos.Recordset.Update
       Else
          Me.AdoHistoricos.Recordset("FechaNacimiento") = FechaNacimiento
          Me.AdoHistoricos.Recordset("FechaContrato") = Format(Now, "dd/mm/yyyy")
          Me.AdoHistoricos.Recordset("FechaContratoVac") = Format(Now, "dd/mm/yyyy")
         Me.AdoHistoricos.Recordset.Update
       End If
       
       
       Me.DtaTurnos.ConnectionString = Conexion
       Me.DtaTurnos.RecordSource = "SELECT CodEmpleado, LEntrada, LSalida, MEntrada, MSalida, MCEntrada, MCSalida, JEntrada, JSalida, VEntrada, VSalida, TComida, TurnoLunes,TurnoMartes , TurnoMiercoles, TurnoJueves, TurnoViernes, TurnoSabado, TurnoDomingo, SEntrada, SSalida, DEntrada, DSalida From dbo.HorarioEmpleado "
       Me.DtaTurnos.Refresh
       
       
       
'     ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////////////////////////////////GRABO EL HORARIO DE EMPLEADOS /////////////////////////////////////////////////////////////
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////
       Me.DtaHorarioEmpleado.ConnectionString = Conexion
       Me.DtaHorarioEmpleado.RecordSource = "SELECT CodEmpleado, LEntrada, LSalida, MEntrada, MSalida, MCEntrada, MCSalida, JEntrada, JSalida, VEntrada, VSalida, TComida, TurnoLunes,TurnoMartes , TurnoMiercoles, TurnoJueves, TurnoViernes, TurnoSabado, TurnoDomingo, SEntrada, SSalida, DEntrada, DSalida From dbo.HorarioEmpleado WHERE(CodEmpleado ='" & Me.TDBCombo1.Text & "')"
       Me.DtaHorarioEmpleado.Refresh
       If Me.DtaHorarioEmpleado.Recordset.EOF Then
         Me.DtaTurnos.Refresh
         If Not Me.DtaTurnos.Recordset.EOF Then
'           CodTurno = Me.DtaTurnos.Recordset("CodTurno")
           Me.DtaHorarioEmpleado.Recordset.AddNew
           Me.DtaHorarioEmpleado.Recordset("CodEmpleado") = CodEmpleado1
           Me.DtaHorarioEmpleado.Recordset("LEntrada") = Me.DtaTurnos.Recordset("LEntrada")
           Me.DtaHorarioEmpleado.Recordset("LSalida") = Me.DtaTurnos.Recordset("LSalida")
           Me.DtaHorarioEmpleado.Recordset("MEntrada") = Me.DtaTurnos.Recordset("MEntrada")
           Me.DtaHorarioEmpleado.Recordset("MSalida") = Me.DtaTurnos.Recordset("MSalida")
           Me.DtaHorarioEmpleado.Recordset("MCEntrada") = Me.DtaTurnos.Recordset("MCEntrada")
           Me.DtaHorarioEmpleado.Recordset("MCSalida") = Me.DtaTurnos.Recordset("MCSalida")
           Me.DtaHorarioEmpleado.Recordset("JEntrada") = Me.DtaTurnos.Recordset("JEntrada")
           Me.DtaHorarioEmpleado.Recordset("JSalida") = Me.DtaTurnos.Recordset("JSalida")
           Me.DtaHorarioEmpleado.Recordset("VEntrada") = Me.DtaTurnos.Recordset("VEntrada")
           Me.DtaHorarioEmpleado.Recordset("VSalida") = Me.DtaTurnos.Recordset("VSalida")
           Me.DtaHorarioEmpleado.Recordset("TComida") = Me.DtaTurnos.Recordset("TComida")
           Me.DtaHorarioEmpleado.Recordset("TurnoLunes") = Me.DtaTurnos.Recordset("TurnoLunes")
           Me.DtaHorarioEmpleado.Recordset("TurnoMartes") = Me.DtaTurnos.Recordset("TurnoMartes")
           Me.DtaHorarioEmpleado.Recordset("TurnoMiercoles") = Me.DtaTurnos.Recordset("TurnoMiercoles")
           Me.DtaHorarioEmpleado.Recordset("TurnoJueves") = Me.DtaTurnos.Recordset("TurnoJueves")
           Me.DtaHorarioEmpleado.Recordset("TurnoViernes") = Me.DtaTurnos.Recordset("TurnoViernes")
           Me.DtaHorarioEmpleado.Recordset("TurnoSabado") = Me.DtaTurnos.Recordset("TurnoSabado")
           Me.DtaHorarioEmpleado.Recordset("TurnoDomingo") = Me.DtaTurnos.Recordset("TurnoDomingo")
           Me.DtaHorarioEmpleado.Recordset("SEntrada") = Me.DtaTurnos.Recordset("SEntrada")
           Me.DtaHorarioEmpleado.Recordset("SSalida") = Me.DtaTurnos.Recordset("SEntrada")
           Me.DtaHorarioEmpleado.Recordset("DEntrada") = Me.DtaTurnos.Recordset("SEntrada")
           Me.DtaHorarioEmpleado.Recordset("DSalida") = Me.DtaTurnos.Recordset("SEntrada")
    
         Me.DtaHorarioEmpleado.Recordset.Update
         End If
         
    End If
    
End If
      
End Sub

Private Sub CmdDesactivar_Click()
Dim Respuesta As String, CodEmpleados As Double
Dim NumNomina As Double
Dim rs As New ADODB.Recordset

 Respuesta = MsgBox("Esta Seguro de Desactivar el Empleado", vbYesNo, "Empleado " & Me.TxtCodEmpleado.Text)
 
 If Respuesta = 6 Then
  CodEmpleados = Me.TxtCodEmpleado.Text
  Me.AdoBusca.RecordSource = "SELECT * From Empleado Where (CodEmpleado = " & CodEmpleados & ")"
  Me.AdoBusca.Refresh
  If Not Me.AdoBusca.Recordset.EOF Then
   If Me.AdoBusca.Recordset("Activo") = True Then
     Me.AdoBusca.Recordset("Activo") = 0
     Me.AdoBusca.Recordset.Update
   End If
  End If
 End If
 
 '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 '////////////SI EL EMPLEADO ESTA EN PLANILLA ACTIVA LO ELIMINO////////////////////////////////////////////////////////////
 '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  Me.AdoNomina.RecordSource = "SELECT Nomina.Activa, Nomina.CodTipoNomina, DetalleNomina.* FROM Nomina INNER JOIN DetalleNomina ON Nomina.NumNomina = DetalleNomina.NumNomina WHERE (Nomina.Activa = 1) AND (DetalleNomina.CodEmpleado = " & CodEmpleados & " )"
  Me.AdoNomina.Refresh
  If Not Me.AdoNomina.Recordset.EOF Then
        NumNomina = Me.AdoNomina.Recordset("NumNomina")
        rs.Open "DELETE FROM DetalleNomina WHERE (NumNomina = " & NumNomina & ") AND (CodEmpleado = " & CodEmpleados & ")", Conexion
  End If
 
  With Me.DtaEmpleados
   .ConnectionString = Conexion
   .RecordSource = "SELECT Empleado.CodEmpleado1, Empleado.CodEmpleado,Empleado.Nombre1 +' '+ Empleado.Nombre2+' ' + Empleado.Apellido1+' ' + Empleado.Apellido2 AS Nombres, Empleado.Activo FROM Departamento INNER JOIN  Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento ORDER BY Empleado.CodEmpleado1 "
   .Refresh
 End With

 MsgBox "Empleado Desactivado", vbExclamation
End Sub

Private Sub CmdDuplicado_Click()
Dim sql As String
On Error GoTo TipoErrs

 sql = "SELECT COUNT(CodEmpleado) AS Registro, CodEmpleado1, MAX(Nombre1 + ' ' + Nombre2 + ' ' + Apellido1 + ' ' + Apellido2) AS Nombres From Empleado Where (Activo = 1) GROUP BY CodEmpleado1 Having (Count(CodEmpleado) > 1)"
 
 ArepEmpleadoDuplicados.DataControl1.ConnectionString = ConexionReporte
 ArepEmpleadoDuplicados.LblTitulo.Caption = Titulo
 ArepEmpleadoDuplicados.LblSubtitulo.Caption = "REPORTE EMPLEADOS DUPLICADOS"
 ArepEmpleadoDuplicados.ImgLogo.Picture = LoadPicture(RutaLogo)
 ArepEmpleadoDuplicados.DataControl1.Source = sql
 ArepEmpleadoDuplicados.Show 1

 Exit Sub
TipoErrs:
 MsgBox Err.Description

End Sub

Private Sub Command1_Click()
FrmTraslados.Show 1
End Sub

Private Sub Command2_Click()

End Sub



Private Sub Form_Load()
Dim SQLHistorial As String

 With Me.DtaPagos
  .ConnectionString = Conexion
 End With
 
  With Me.AdoBusca
  .ConnectionString = Conexion
 End With
 
  With Me.AdoHistorial
  .ConnectionString = Conexion
 End With
 
   With Me.AdoNomina
  .ConnectionString = Conexion
 End With
 
 With Me.DtaConsulta
  .ConnectionString = Conexion
 End With
 
  With Me.AdoHistoricos
  .ConnectionString = Conexion
 End With
 
  With Me.DtaEmpleadosNuevos
   .ConnectionString = Conexion
   .RecordSource = "SELECT Empleado.*  FROM Empleado ORDER BY Empleado.CodEmpleado1 "
   .Refresh
 End With
 
 With Me.DtaEmpleados
   .ConnectionString = Conexion
   .RecordSource = "SELECT Empleado.CodEmpleado1, Empleado.CodEmpleado,Empleado.Nombre1 +' '+ Empleado.Nombre2+' ' + Empleado.Apellido1+' ' + Empleado.Apellido2 AS Nombres, Empleado.Activo FROM Departamento INNER JOIN  Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento ORDER BY Empleado.CodEmpleado1 "
   .Refresh
 End With
 
CmdAnnoIni = Str(Year(Now))
CmdAnnoFin = Str(Year(Now))

MDIPrimero.Skin1.ApplySkin hWnd


Me.CmdActivar.BackColor = RGB(219, 226, 242)
Me.CmdDesactivar.BackColor = RGB(219, 226, 242)
Me.CmdCopiar.BackColor = RGB(219, 226, 242)

 Me.TDBGrid1.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.TDBGrid1.OddRowStyle.BackColor = &H80000005
 Me.TDBGrid1.AlternatingRowStyle = True
 
 
 Me.TDBGrid1.Columns(0).Width = 700
 Me.TDBGrid1.Columns(1).Width = 700
 Me.TDBGrid1.Columns(2).Width = 1200
 Me.TDBGrid1.Columns(3).Width = 1200
 Me.TDBGrid1.Columns(4).Width = 1200
 Me.TDBGrid1.Columns(5).Width = 1200
 Me.TDBGrid1.Columns(6).Width = 1200
 Me.TDBGrid1.Columns(7).Width = 1200
 Me.TDBGrid1.Columns(8).Width = 1200
 Me.TDBGrid1.Columns(9).Width = 1200
 Me.TDBGrid1.Columns(10).Width = 1200
 Me.TDBGrid1.Columns(11).Width = 1200
 Me.TDBGrid1.Columns(12).Width = 1200
 Me.TDBGrid1.Columns(13).Width = 1200
 Me.TDBGrid1.Columns(14).Width = 1200
 Me.TDBGrid1.Columns(15).Width = 1200
 Me.TDBGrid1.Columns(16).Width = 1400
 Me.TDBGrid1.Columns(17).Width = 1400
 Me.TDBGrid1.Columns(18).Width = 1400
 Me.TDBGrid1.Columns(19).Width = 1400
 Me.TDBGrid1.Columns(20).Width = 1400

 
 Me.TDBCombo1.Columns(0).Width = 1200
 Me.TDBCombo1.Columns(1).Width = 1200
 Me.TDBCombo1.Columns(2).Width = 4000
 Me.TDBCombo1.Columns(3).Width = 100
 
 
 
 
CmbMesIni = "Enero"
CmdMesFin = "Diciembre"



Me.Top = 1000
Me.Left = 1000


End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TDBCombo1_ItemChange()
Me.TxtCodEmpleado.Text = Me.TDBCombo1.Columns(1).Text
If Me.TDBCombo1.Columns(3).Text = -1 Then
   Me.SSCommand1.Caption = "Empleado Activo"
   Me.CmdDesactivar.Visible = True
   Me.CmdActivar.Visible = False
   Me.CmdCopiar.Visible = False
   Me.CmdConstanciaActivos.Visible = True
   EmpleadoActivo = True
Else
   Me.SSCommand1.Caption = "Empleado Inactivo"
   Me.CmdDesactivar.Visible = False
   Me.CmdActivar.Visible = True
   Me.CmdCopiar.Visible = True
   EmpleadoActivo = False
   Me.CmdConstanciaActivos.Visible = True
End If
End Sub

Private Sub TDBCombo1_KeyPress(KeyAscii As Integer)
 Dim Codigo As String
 If KeyAscii = 13 Then
  Me.AdoBusca.RecordSource = "SELECT CodEmpleado, CodEmpleado1, Activo From Empleado WHERE (CodEmpleado1 = '" & Me.TDBCombo1.Text & "') AND (Activo = 1)"
  Me.AdoBusca.Refresh
  If Not Me.AdoBusca.Recordset.EOF Then
     Me.TxtCodEmpleado.Text = Me.AdoBusca.Recordset("CodEmpleado")
   
  End If
 End If
End Sub

Private Sub txtCodEmpleado_Change()
On Error GoTo TipoErr
Dim SqlPagos As String
Dim SqlEmpleados As String
Dim MesIni As Byte
Dim Annoini As Integer
Dim MesFin As Byte
Dim AnnoFin As Integer
Dim CodEmpleado As Double
Annoini = val(CmdAnnoIni.Text)

CodEmpleado = Me.TxtCodEmpleado.Text



SqlEmpleados = "SELECT CodEmpleado1, CodEmpleado, Nombre1 + Nombre2 + Apellido1 + Apellido2 AS Nombres, Nombre1, Nombre2, Apellido1, Apellido2, NumHijos,Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula, Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss,NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria, PorcentajeComision, OtrosIngresos, DescripOtrIngre, ExentoInss,ExentoIr , PagoInssPatronal, SalarioMinimo, Observaciones, Liquidado, Ausente, Activo From Empleado Where (CodEmpleado = " & CodEmpleado & ") ORDER BY CodEmpleado1"
Me.AdoBusca.RecordSource = SqlEmpleados
Me.AdoBusca.Refresh

If Not Me.AdoBusca.Recordset.EOF Then
    TxtNombre1.Text = Me.AdoBusca.Recordset("Nombre1")
    TxtNombre2.Text = Me.AdoBusca.Recordset("Nombre2")
    TxtApellido1.Text = Me.AdoBusca.Recordset("Apellido1")
    TxtApellido2.Text = Me.AdoBusca.Recordset("Apellido2")
    Cargo = Me.AdoBusca.Recordset("Cargo")
    Sexo = Me.AdoBusca.Recordset("Sexo")
    Cedula = Me.AdoBusca.Recordset("NumCedula")
    Basico = Me.AdoBusca.Recordset("SueldoPeriodo")
'    TxtCargo.Text = Me.AdoBusca.Recordset("Cargo")
'    txtDepartamento = Me.AdoBusca.Recordset("departamento")
Else
    TxtNombre1.Text = ""
    TxtNombre2.Text = ""
    TxtApellido1.Text = ""
    TxtApellido2.Text = ""
    TxtCargo.Text = ""
    TxtDepartamento = ""
End If
'DbgrPagos.Columns(0).Visible = False

'SELECT DISTINCT
'                      TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)
'                      AS Destajo, SUM(dbo.DetalleNomina.Incentivos) AS Incentivos, SUM(dbo.DetalleNomina.HE) AS HE, SUM(dbo.DetalleNomina.HorasExtras)
'                      AS HorasExtras, SUM(dbo.DetalleNomina.Comisiones) AS Comisiones, SUM(dbo.DetalleNomina.OtrosIngresos) AS OtrosIngresos,
'                      SUM (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.OtrosIngresos)
'                       AS TotalIngresos, SUM(dbo.DetalleNomina.Deducciones) AS Deducciones, SUM(dbo.DetalleNomina.Prestamo) AS Prestamo,
'                      SUM(dbo.DetalleNomina.MontoINSS) AS MontoInss, SUM(dbo.DetalleNomina.MontoIR) AS MontoIR,
'                      SUM (dbo.DetalleNomina.Deducciones + dbo.DetalleNomina.Prestamo + dbo.DetalleNomina.MontoInss + dbo.DetalleNomina.MontoIR)
'                      AS TotalEgresos,
'                      SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.OtrosIngresos
'                       - dbo.DetalleNomina.Deducciones - dbo.DetalleNomina.Prestamo - dbo.DetalleNomina.MontoINSS - dbo.DetalleNomina.MontoIR) AS NetoPagar,
'                      SUM(dbo.DetalleNomina.INSSPatronal) AS INSSPATRONAL, SUM(dbo.DetalleNomina.IRPatronal) AS IRPATRONAL, SUM(dbo.DetalleNomina.INATEC)
'                      AS INATEC, SUM(dbo.DetalleNomina.IncetivoProduccion) AS INCENTIVOPRODUCCION, SUM(dbo.DetalleNomina.TarifaHoraria) AS TARIFA,
'                      MIN(dbo.Nomina.FechaNomina) AS Fecha, dbo.Nomina.Mes AS MES, dbo.Nomina.Ano AS AÑO
'FROM         dbo.DetalleNomina INNER JOIN
'                      dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina
'GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano
'Having (dbo.DetalleNomina.CodEmpleado = 2015)


SQLHistorial = "SELECT DISTINCT" & vbLf
SQLHistorial = SQLHistorial & "TOP 100 PERCENT dbo.DetalleNomina.CodEmpleado, SUM(dbo.DetalleNomina.SalarioBasico) AS SalarioBasico, SUM(dbo.DetalleNomina.Destajo)" & vbLf
SQLHistorial = SQLHistorial & "AS Destajo, SUM(dbo.DetalleNomina.Incentivos) AS Incentivos, SUM(dbo.DetalleNomina.HE) AS HE, SUM(dbo.DetalleNomina.HorasExtras)" & vbLf
SQLHistorial = SQLHistorial & "AS HorasExtras, SUM(dbo.DetalleNomina.Comisiones) AS Comisiones, SUM(dbo.DetalleNomina.OtrosIngresos) AS OtrosIngresos," & vbLf
SQLHistorial = SQLHistorial & "SUM (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.OtrosIngresos)" & vbLf
SQLHistorial = SQLHistorial & "AS TotalIngresos, SUM(dbo.DetalleNomina.Deducciones) AS Deducciones, SUM(dbo.DetalleNomina.Prestamo) AS Prestamo," & vbLf
SQLHistorial = SQLHistorial & "SUM(dbo.DetalleNomina.MontoINSS) AS MontoInss, SUM(dbo.DetalleNomina.MontoIR) AS MontoIR," & vbLf
SQLHistorial = SQLHistorial & "SUM (dbo.DetalleNomina.Deducciones + dbo.DetalleNomina.Prestamo + dbo.DetalleNomina.MontoInss + dbo.DetalleNomina.MontoIR)" & vbLf
SQLHistorial = SQLHistorial & "AS TotalEgresos," & vbLf
SQLHistorial = SQLHistorial & "SUM(dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.OtrosIngresos" & vbLf
SQLHistorial = SQLHistorial & "- dbo.DetalleNomina.Deducciones - dbo.DetalleNomina.Prestamo - dbo.DetalleNomina.MontoINSS - dbo.DetalleNomina.MontoIR) AS NetoPagar," & vbLf
SQLHistorial = SQLHistorial & "SUM(dbo.DetalleNomina.INSSPatronal) AS INSSPATRONAL, SUM(dbo.DetalleNomina.IRPatronal) AS IRPATRONAL, SUM(dbo.DetalleNomina.INATEC)" & vbLf
SQLHistorial = SQLHistorial & "AS INATEC, SUM(dbo.DetalleNomina.IncetivoProduccion) AS INCENTIVOPRODUCCION, SUM(dbo.DetalleNomina.TarifaHoraria) AS TARIFA," & vbLf
SQLHistorial = SQLHistorial & "MIN(dbo.Nomina.FechaNomina) AS Fecha, dbo.Nomina.Mes AS MES, dbo.Nomina.Ano AS AÑO" & vbLf
SQLHistorial = SQLHistorial & "FROM         dbo.DetalleNomina INNER JOIN" & vbLf
SQLHistorial = SQLHistorial & "dbo.Nomina ON dbo.DetalleNomina.NumNomina = dbo.Nomina.NumNomina" & vbLf
SQLHistorial = SQLHistorial & "GROUP BY dbo.DetalleNomina.CodEmpleado, dbo.Nomina.Mes, dbo.Nomina.Ano" & vbLf
SQLHistorial = SQLHistorial & "Having (dbo.DetalleNomina.CodEmpleado = " & CodEmpleado & ")"






Me.AdoHistorial.RecordSource = SQLHistorial
Me.AdoHistorial.Refresh

 Me.TDBGrid1.Columns(0).Width = 700
 Me.TDBGrid1.Columns(1).Width = 700
 Me.TDBGrid1.Columns(2).Width = 1200
 Me.TDBGrid1.Columns(3).Width = 1200
 Me.TDBGrid1.Columns(4).Width = 1200
 Me.TDBGrid1.Columns(5).Width = 1200
 Me.TDBGrid1.Columns(6).Width = 1200
 Me.TDBGrid1.Columns(7).Width = 1200
 Me.TDBGrid1.Columns(8).Width = 1200
 Me.TDBGrid1.Columns(9).Width = 1200
 Me.TDBGrid1.Columns(10).Width = 1200
 Me.TDBGrid1.Columns(11).Width = 1200
 Me.TDBGrid1.Columns(12).Width = 1200
 Me.TDBGrid1.Columns(13).Width = 1200
 Me.TDBGrid1.Columns(14).Width = 1200
 Me.TDBGrid1.Columns(15).Width = 1200
 Me.TDBGrid1.Columns(16).Width = 1400
 Me.TDBGrid1.Columns(17).Width = 1400
 Me.TDBGrid1.Columns(18).Width = 1400
 Me.TDBGrid1.Columns(19).Width = 1400
 Me.TDBGrid1.Columns(20).Width = 1400

Exit Sub
TipoErr:
    ControlErrores
End Sub
