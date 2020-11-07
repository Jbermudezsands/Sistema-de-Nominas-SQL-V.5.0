VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmCopiaMarcas 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar Marcas Entre Compañias"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8325
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Seleccion Nomina"
      TabPicture(0)   =   "FrmCopiaMarcas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "osProgress2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "osProgress1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdSalir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdIniciar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Conexion Origen"
      TabPicture(1)   =   "FrmCopiaMarcas.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1"
      Tab(1).Control(1)=   "TxtConexionString"
      Tab(1).Control(2)=   "Command3"
      Tab(1).Control(3)=   "Label1"
      Tab(1).ControlCount=   4
      Begin VB.CommandButton Command1 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   -69360
         TabIndex        =   19
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton CmdIniciar 
         Caption         =   "Copiar"
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   6600
         TabIndex        =   15
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   1935
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   7575
         Begin VB.TextBox TxtNumeroNomina 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   7
            Top             =   1005
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   480
            OleObjectBlob   =   "FrmCopiaMarcas.frx":0038
            TabIndex        =   8
            Top             =   1005
            Width           =   1335
         End
         Begin TrueOleDBList80.TDBCombo TDBTipo 
            Bindings        =   "FrmCopiaMarcas.frx":00B0
            DataSource      =   "AdoTipoNominas"
            Height          =   315
            Left            =   1200
            TabIndex        =   9
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
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
            ListField       =   ""
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
            _PropDict       =   $"FrmCopiaMarcas.frx":00CD
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   3120
            OleObjectBlob   =   "FrmCopiaMarcas.frx":0177
            TabIndex        =   10
            Top             =   645
            Width           =   255
         End
         Begin MSComCtl2.DTPicker DtpFechaINI 
            Height          =   300
            Left            =   960
            TabIndex        =   11
            Top             =   645
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            Format          =   79233025
            CurrentDate     =   40789
         End
         Begin MSComCtl2.DTPicker DTFechaFin 
            Height          =   300
            Left            =   3480
            TabIndex        =   12
            Top             =   645
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            Format          =   79233025
            CurrentDate     =   40789
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   480
            OleObjectBlob   =   "FrmCopiaMarcas.frx":01DB
            TabIndex        =   13
            Top             =   645
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   480
            OleObjectBlob   =   "FrmCopiaMarcas.frx":0245
            TabIndex        =   14
            Top             =   285
            Width           =   735
         End
      End
      Begin VB.TextBox TxtConexionString 
         Height          =   1515
         Left            =   -73800
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   5655
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   -67920
         Picture         =   "FrmCopiaMarcas.frx":02B1
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin XtremeSuiteControls.ProgressBar osProgress1 
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   2640
         Width           =   7575
         _Version        =   786432
         _ExtentX        =   13361
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   14737632
         Scrolling       =   1
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ProgressBar osProgress2 
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3120
         Visible         =   0   'False
         Width           =   4815
         _Version        =   786432
         _ExtentX        =   8493
         _ExtentY        =   450
         _StockProps     =   93
         BackColor       =   14737632
         Appearance      =   6
      End
      Begin VB.Label Label1 
         Caption         =   "Conexion"
         Height          =   375
         Left            =   -74760
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   8895
      TabIndex        =   0
      Top             =   -120
      Width           =   8895
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Copiar Marcas entre Compañias"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   4200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   8880
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image2 
         Height          =   1020
         Left            =   0
         Picture         =   "FrmCopiaMarcas.frx":0767
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1290
      End
   End
   Begin MSComDlg.CommonDialog CMRutaFoto 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   256
   End
   Begin MSAdodcLib.Adodc AdoTipoNominas 
      Height          =   375
      Left            =   840
      Top             =   6120
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
      Caption         =   "AdoTipoNominas"
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
   Begin MSAdodcLib.Adodc AdoConsultaNomina 
      Height          =   375
      Left            =   4560
      Top             =   6120
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
      Caption         =   "AdoConsultaNomina"
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
   Begin MSAdodcLib.Adodc AdoMarcasOrigen 
      Height          =   375
      Left            =   840
      Top             =   6840
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
      Caption         =   "AdoMarcasOrgien"
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
   Begin MSAdodcLib.Adodc AdoEmpleados 
      Height          =   375
      Left            =   4200
      Top             =   6960
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
   Begin MSAdodcLib.Adodc AdoMarcasDestino 
      Height          =   375
      Left            =   720
      Top             =   5400
      Width           =   3135
      _ExtentX        =   5530
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
      Caption         =   "AdoMarcasDestino"
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
   Begin MSAdodcLib.Adodc AdoMarcasDestino2 
      Height          =   375
      Left            =   3480
      Top             =   7920
      Width           =   3135
      _ExtentX        =   5530
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
      Caption         =   "AdoMarcasDestino2"
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
Attribute VB_Name = "FrmCopiaMarcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ConexionOrigen As String


Private Sub CmdIniciar_Click()
 Dim CodTipoNomina As String, Registros As Double, Registros2 As Double
 Dim FechaIni As String, FechaFin As String, CodEmpleado As Double, FechaMarca As String
 Dim CodEmpleado1 As String, CodEmpleado2 As Double
 
 CodTipoNomina = Me.TDBTipo.Columns(0).Text
 FechaIni = Format(Me.DTPFechaIni.Value, "yyyy-mm-dd")
 FechaFin = Format(Me.DTFechaFin.Value, "yyyy-mm-dd")
 
 Me.AdoEmpleados.RecordSource = "SELECT  * From Empleado WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (Activo = 1)"
 Me.AdoEmpleados.Refresh
 
 If Not Me.AdoEmpleados.Recordset.EOF Then
    Me.AdoEmpleados.Recordset.MoveLast
    Registros = Me.AdoEmpleados.Recordset.RecordCount
    Me.AdoEmpleados.Recordset.MoveFirst
 End If
 
 Me.osProgress1.Min = 0
 Me.osProgress1.Value = 0
 Me.osProgress1.Max = Registros
 Do While Not Me.AdoEmpleados.Recordset.EOF
   CodEmpleado1 = Me.AdoEmpleados.Recordset("CodEmpleado1")
   CodEmpleado = Me.AdoEmpleados.Recordset("CodEmpleado")
   
   Me.Caption = "Procesando : " & CodEmpleado1
   DoEvents
   '///////////////////////////////CONSULTA DE MARCAS ORIGEN ///////////////////////////////////////
   Me.AdoMarcasOrigen.RecordSource = "SELECT  * FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado  " & _
                                     "WHERE (AsistenciaEmpleado.FechaEntrada BETWEEN CONVERT(DATETIME, '" & FechaIni & "', 102) AND CONVERT(DATETIME, '" & FechaFin & "', 102)) AND (AsistenciaEmpleado.CodEmpleado1 = '" & CodEmpleado1 & "') AND (Empleado.Activo = 1)"
   Me.AdoMarcasOrigen.Refresh
   
    If Not Me.AdoMarcasOrigen.Recordset.EOF Then
      Me.AdoMarcasOrigen.Recordset.MoveLast
      Registros2 = Me.AdoMarcasOrigen.Recordset.RecordCount
      Me.AdoMarcasOrigen.Recordset.MoveFirst
    End If
    
    Me.osProgress2.Visible = True
    Me.osProgress2.Min = 0
    Me.osProgress2.Value = 0
    Me.osProgress2.Max = Registros2
   Do While Not Me.AdoMarcasOrigen.Recordset.EOF
     FechaMarca = Format(CDate(Me.AdoMarcasOrigen.Recordset("FechaEntrada")), "yyyy-mm-dd")
     Me.AdoMarcasDestino.RecordSource = "SELECT  * FROM AsistenciaEmpleado INNER JOIN Empleado ON AsistenciaEmpleado.CodEmpleado = Empleado.CodEmpleado WHERE (AsistenciaEmpleado.FechaEntrada = CONVERT(DATETIME, '" & FechaMarca & "', 102)) AND (AsistenciaEmpleado.CodEmpleado1 = '" & CodEmpleado1 & "') AND  (Empleado.Activo = 1)"
     Me.AdoMarcasDestino.Refresh
     If Me.AdoMarcasDestino.Recordset.EOF Then
      DoEvents
                 Me.AdoMarcasDestino2.RecordSource = "SELECT AsistenciaEmpleado.* From AsistenciaEmpleado "
                 Me.AdoMarcasDestino2.Refresh
                        Me.AdoMarcasDestino2.Recordset.AddNew
                        Me.AdoMarcasDestino2.Recordset("CodEmpleado") = CodEmpleado
                        Me.AdoMarcasDestino2.Recordset.Fields("CodEmpleado1") = CodEmpleado1
                        Me.AdoMarcasDestino2.Recordset.Fields("FechaEntrada") = Me.AdoMarcasOrigen.Recordset.Fields("FechaEntrada")
                        
                        If Not IsNull(Me.AdoMarcasOrigen.Recordset.Fields("HoraEntrada")) Then
                          Me.AdoMarcasDestino2.Recordset.Fields("HoraEntrada") = Me.AdoMarcasOrigen.Recordset.Fields("HoraEntrada")
                        End If
                        
                        If Not IsNull(Me.AdoMarcasOrigen.Recordset.Fields("HREntrada")) Then
                          Me.AdoMarcasDestino2.Recordset.Fields("HREntrada") = Me.AdoMarcasOrigen.Recordset.Fields("HREntrada")
                        End If
                        If Not IsNull(Me.AdoMarcasOrigen.Recordset.Fields("HRSalida")) Then
                          Me.AdoMarcasDestino2.Recordset.Fields("HRSalida") = Me.AdoMarcasOrigen.Recordset.Fields("HRSalida")
                        End If
                        If Not IsNull(Me.AdoMarcasOrigen.Recordset.Fields("FechaSalida")) Then
                          Me.AdoMarcasDestino2.Recordset.Fields("FechaSalida") = Me.AdoMarcasOrigen.Recordset.Fields("FechaSalida")
                        End If
                        If Not IsNull(Me.AdoMarcasOrigen.Recordset.Fields("HoraSalida")) Then
                          Me.AdoMarcasDestino2.Recordset.Fields("HoraSalida") = Me.AdoMarcasOrigen.Recordset.Fields("HoraSalida")
                        End If
                       
                      
                        
                        Me.AdoMarcasDestino2.Recordset.Fields("CodTurno") = Me.AdoMarcasOrigen.Recordset.Fields("CodTurno")
                        Me.AdoMarcasDestino2.Recordset.Fields("CodTipoNomina") = Me.AdoMarcasOrigen.Recordset.Fields("CodTipoNomina")
                        Me.AdoMarcasDestino2.Recordset.Fields("bActivo") = Me.AdoMarcasOrigen.Recordset.Fields("bActivo")
                        Me.AdoMarcasDestino2.Recordset.Update
                        
     Else

                 CodEmpleado2 = Me.AdoMarcasDestino.Recordset.Fields("CodEmpleado")
                 Me.AdoMarcasDestino2.RecordSource = "SELECT AsistenciaEmpleado.* From AsistenciaEmpleado WHERE (AsistenciaEmpleado.FechaEntrada = CONVERT(DATETIME, '" & FechaMarca & "', 102)) AND (AsistenciaEmpleado.CodEmpleado = " & CodEmpleado2 & ") "
                 Me.AdoMarcasDestino2.Refresh
                 If Not Me.AdoMarcasDestino2.Recordset.EOF Then
                        Me.AdoMarcasDestino2.Recordset.Fields("FechaEntrada") = Me.AdoMarcasOrigen.Recordset.Fields("FechaEntrada")
                        
                        If Not IsNull(Me.AdoMarcasOrigen.Recordset.Fields("HoraEntrada")) Then
                          Me.AdoMarcasDestino2.Recordset.Fields("HoraEntrada") = Me.AdoMarcasOrigen.Recordset.Fields("HoraEntrada")
                        End If
                        
                        If Not IsNull(Me.AdoMarcasOrigen.Recordset.Fields("HREntrada")) Then
                          Me.AdoMarcasDestino2.Recordset.Fields("HREntrada") = Me.AdoMarcasOrigen.Recordset.Fields("HREntrada")
                        End If
                        If Not IsNull(Me.AdoMarcasOrigen.Recordset.Fields("HRSalida")) Then
                          Me.AdoMarcasDestino2.Recordset.Fields("HRSalida") = Me.AdoMarcasOrigen.Recordset.Fields("HRSalida")
                        End If
                        If Not IsNull(Me.AdoMarcasOrigen.Recordset.Fields("FechaSalida")) Then
                          Me.AdoMarcasDestino2.Recordset.Fields("FechaSalida") = Me.AdoMarcasOrigen.Recordset.Fields("FechaSalida")
                        End If
                        If Not IsNull(Me.AdoMarcasOrigen.Recordset.Fields("HoraSalida")) Then
                          Me.AdoMarcasDestino2.Recordset.Fields("HoraSalida") = Me.AdoMarcasOrigen.Recordset.Fields("HoraSalida")
                        End If
                       
                      
                        
                        Me.AdoMarcasDestino2.Recordset.Fields("CodTurno") = Me.AdoMarcasOrigen.Recordset.Fields("CodTurno")
                        Me.AdoMarcasDestino2.Recordset.Fields("CodTipoNomina") = Me.AdoMarcasOrigen.Recordset.Fields("CodTipoNomina")
                        Me.AdoMarcasDestino2.Recordset.Fields("bActivo") = Me.AdoMarcasOrigen.Recordset.Fields("bActivo")
                        Me.AdoMarcasDestino2.Recordset.Update
     
                  End If
     End If
     
     DoEvents
     Me.osProgress2.Value = Me.osProgress2.Value + 1
     Me.AdoMarcasOrigen.Recordset.MoveNext
   Loop
   
   DoEvents
   Me.osProgress2.Visible = False
   Me.osProgress1.Value = Me.osProgress1.Value + 1
   Me.AdoEmpleados.Recordset.MoveNext
 Loop
 


End Sub

Private Sub CmdSalir_Click()
 Unload Me
End Sub

Private Sub Command1_Click()
  MDIPrimero.DtaEmpresa.Recordset("ConexionCopia") = Me.TxtConexionString.Text
  MDIPrimero.DtaEmpresa.Recordset.Update
  MsgBox "Registro Grabado", vbExclamation, "Zeus Nominas"


End Sub

Private Sub Command3_Click()
On Error GoTo TipoErrs
Dim mydlg As New MSDASC.DataLinks
Dim ADOcon As New ADODB.Connection

Me.TxtConexionString.Text = mydlg.PromptNew


Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub Form_Load()
  With Me.AdoTipoNominas
   .ConnectionString = Conexion
   .RecordSource = "SELECT  CodTipoNomina, Nomina, Periodo From TipoNomina Where (Activa = 1)"
   .Refresh
 End With
 
With Me.AdoConsultaNomina
  .ConnectionString = Conexion
End With

With Me.AdoEmpleados
  .ConnectionString = Conexion
End With

With Me.AdoMarcasDestino
  .ConnectionString = Conexion
End With

With Me.AdoMarcasDestino2
  .ConnectionString = Conexion
End With

'/////////////////////////////////ORGIEN DE REGISTROS ///////////////////////////////////////////////////

If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("ConexionCopia")) Then
  Me.TxtConexionString.Text = MDIPrimero.DtaEmpresa.Recordset("ConexionCopia")
End If

If Me.TxtConexionString.Text <> "" Then
  ConexionOrigen = Me.TxtConexionString.Text
  
  
  With Me.AdoMarcasOrigen
  .ConnectionString = ConexionOrigen
  End With
  
  
End If
End Sub

Private Sub TDBTipo_ItemChange()
  Me.AdoConsultaNomina.RecordSource = "SELECT  * From Nomina WHERE (Activa = 1) AND (CodTipoNomina = '" & Me.TDBTipo.Text & "')"
  Me.AdoConsultaNomina.Refresh
  If Not Me.AdoConsultaNomina.Recordset.EOF Then
    Me.DTPFechaIni.Value = Me.AdoConsultaNomina.Recordset("FechaNominaINI")
    Me.DTFechaFin.Value = Me.AdoConsultaNomina.Recordset("FechaNomina")
    Me.TxtNumeroNomina.Text = Me.AdoConsultaNomina.Recordset("NumNomina")
  End If
End Sub
