VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmActividades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrador de Actividades"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   12105
   Begin VB.Frame Frame4 
      Enabled         =   0   'False
      Height          =   735
      Left            =   6120
      TabIndex        =   16
      Top             =   240
      Width           =   5775
      Begin VB.OptionButton chkPago 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton chkPago 
         Caption         =   "Empresa"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmActividades.frx":0000
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   12
      Top             =   6720
      Width           =   5775
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   2040
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   4680
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdDelPA 
         Caption         =   "Eli&minar"
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraCuenta 
      Height          =   3615
      Left            =   6120
      TabIndex        =   3
      Top             =   1200
      Width           =   5775
      Begin VB.CommandButton cmdAgregarCta 
         Caption         =   "+"
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
         Left            =   5280
         TabIndex        =   10
         Tag             =   "1"
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdBorrarCta 
         Caption         =   "-"
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
         Left            =   5280
         TabIndex        =   5
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtCuenta 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   1995
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmActividades.frx":0084
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmActividades.frx":0100
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin TrueOleDBList80.TDBCombo tdbcNomina 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
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
         AutoCompletion  =   -1  'True
         LimitToList     =   0   'False
         ColumnHeaders   =   0   'False
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
         _PropDict       =   $"frmActividades.frx":016A
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
      Begin TrueOleDBGrid80.TDBGrid tdbgCuentas 
         Height          =   2535
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4471
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
         Splits(0).FilterBar=   -1  'True
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
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42,.ellipsis=0"
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
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12,.ellipsis=0"
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
         Left            =   120
         OleObjectBlob   =   "frmActividades.frx":0214
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6120
      TabIndex        =   1
      Top             =   5400
      Width           =   1335
      Begin VB.CommandButton cmdAddPA 
         Caption         =   "&Proc. / Activ."
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddArea 
         Caption         =   "Áre&as"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList ilnode 
      Left            =   4440
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActividades.frx":029C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActividades.frx":0A60
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvActividades 
      Height          =   7095
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   12515
      _Version        =   393217
      Indentation     =   353
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ilnode"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmActividades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnx As New ADODB.Connection
Private rs As New ADODB.Recordset
Private rs4 As New ADODB.Recordset
Private rs3 As New ADODB.Recordset
Private sql As String
Private modal As Boolean

Private Sub cmdAddArea_Click()
modal = True
frmNodes.prRaiz = 0
frmNodes.prLlave = -1
frmNodes.Show
End Sub

Private Sub cmdAddPA_Click()
If tvActividades.Nodes.Count = 0 Then _
    MsgBox "Agregue áreas para agregar procesos y actividades", vbInformation: Exit Sub
modal = True
frmNodes.prRaiz = val(Mid(Me.tvActividades.SelectedItem.Key, 2))
frmNodes.prLlave = -1
frmNodes.Show
End Sub

Private Sub cmdAgregarCta_Click()
On Error GoTo errcta
If Me.tdbgCuentas.Visible Then
    Me.tvActividades.Enabled = False
    Me.tdbgCuentas.Visible = False
    Me.SkinLabel6.Visible = True
    Me.SkinLabel7.Visible = True
    Me.tdbcNomina.Visible = True
    Me.txtCuenta.Visible = True
    Me.txtCuenta.Text = ""
    
    sql = "SELECT [Nomina], [CodTipoNomina] From [dbo].[TipoNomina] " & _
            "Where [Activa] = 1 AND [CodTipoNomina] NOT IN ( " & _
                "SELECT [TipoNomina] FROM [dbo].[ActividadesCuenta] "
    If tvActividades.SelectedItem.FullPath = tvActividades.SelectedItem.Text Then
        sql = sql & "WHERE codigo = " & Left(tvActividades.SelectedItem.Text, InStr(1, tvActividades.SelectedItem.Text, ".") - 1) & " AND raiz = 0)"
    Else
        sql = sql & "WHERE codigo = " & Left(tvActividades.SelectedItem.Text, InStr(1, tvActividades.SelectedItem.Text, ".") - 1) & " AND raiz = " & Mid(tvActividades.SelectedItem.Parent.Key, 2) & ")"
    End If
    
    With rs4
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .Open sql, cnx, adOpenDynamic, adLockOptimistic
    End With
    Me.tdbcNomina.RowSource = rs4
    Me.tdbcNomina.BoundColumn = "CodTipoNomina"
    Me.tdbcNomina.Refresh
    Me.tdbcNomina.Columns(1).Visible = False
    Me.tdbcNomina.Text = ""
Else
    If Me.tdbcNomina.BoundText = "" Then
        MsgBox "Seleccione una nomina", vbInformation
        Me.tdbcNomina.SetFocus
        Exit Sub
    End If
    If Trim(Me.txtCuenta) = "" Then
        MsgBox "Digite la cuenta que asignará", vbInformation
        Me.txtCuenta.SetFocus
        Exit Sub
    End If
    
    sql = "INSERT INTO [dbo].[ActividadesCuenta]([TipoNomina], [Cuenta], [Codigo], [Raiz]) " & _
            "Values('" & Trim(Me.tdbcNomina.BoundText) & "', '" & Trim(Me.txtCuenta) & "', " & Left(Me.tvActividades.SelectedItem.Text, InStr(1, Me.tvActividades.SelectedItem.Text, ".") - 1)
    If tvActividades.SelectedItem.FullPath = tvActividades.SelectedItem.Text Then
        sql = sql & ", 0)"
    Else
        sql = sql & ", " & Mid(tvActividades.SelectedItem.Parent.Key, 2) & ")"
    End If

    cnx.Execute sql
    MsgBox "Cuenta Agregada", vbInformation
    Call cmdBorrarCta_Click
End If

Exit Sub
errcta:
    MsgBox Err.Description
End Sub

Private Sub cmdBorrarCta_Click()
On Error GoTo errcta
If Me.tdbgCuentas.Visible Then
    If rs3.eof Then MsgBox "Operación Cancelada, no existen registros", vbInformation: Exit Sub
    If MsgBox("¿Desea eliminar el registro?", vbYesNo) = vbYes Then
        If tvActividades.SelectedItem.FullPath = tvActividades.SelectedItem.Text Then
            sql = "DELETE FROM [dbo].[ActividadesCuenta] WHERE codigo = " & Left(tvActividades.SelectedItem.Text, InStr(1, tvActividades.SelectedItem.Text, ".") - 1) & " AND raiz = 0"
        Else
            sql = "DELETE FROM [dbo].[ActividadesCuenta] WHERE codigo = " & Left(tvActividades.SelectedItem.Text, InStr(1, tvActividades.SelectedItem.Text, ".") - 1) & " AND raiz = " & Mid(tvActividades.SelectedItem.Parent.Key, 2)
        End If
        sql = sql & " and TipoNomina = '" & Trim(rs3!Id) & "'"
        cnx.Execute sql
        MsgBox "Registro eliminado", vbInformation
    End If
Else
    Me.tvActividades.Enabled = True
    Me.tdbgCuentas.Visible = True
    Me.SkinLabel6.Visible = False
    Me.SkinLabel7.Visible = False
    Me.tdbcNomina.Visible = False
    Me.txtCuenta.Visible = False
End If
Call tvActividades_NodeClick(Me.tvActividades.SelectedItem)
Exit Sub
errcta:
    MsgBox Err.Description
End Sub

Private Sub CmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdDelPA_Click()
On Error GoTo errdel
If rs.eof Then MsgBox "Operación Cancelada, no existen registros", vbInformation: Exit Sub
If tvActividades.SelectedItem.Children > 0 Then
    MsgBox "Operación cancelada, elimine los " & tvActividades.SelectedItem.Children & " items que tiene el nodo"
    Exit Sub
End If

If MsgBox("Esta seguro de eliminar el items", vbYesNo) = vbNo Then Exit Sub

If tvActividades.SelectedItem.FullPath = tvActividades.SelectedItem.Text Then
    sql = "DELETE FROM Actividades WHERE codigo = " & Left(tvActividades.SelectedItem.Text, InStr(1, tvActividades.SelectedItem.Text, ".") - 1) & " AND raiz = 0"
Else
    sql = "DELETE FROM Actividades WHERE codigo = " & Left(tvActividades.SelectedItem.Text, InStr(1, tvActividades.SelectedItem.Text, ".") - 1) & " AND raiz = " & Mid(tvActividades.SelectedItem.Parent.Key, 2)
End If

cnx.Execute sql
tvActividades.Nodes.Remove tvActividades.SelectedItem.Index

Exit Sub
errdel:
    MsgBox Err.Description
End Sub



Private Sub cmdEditar_Click()
If rs.eof Then MsgBox "Operación Cancelada, no existen registros", vbInformation: Exit Sub
modal = True
If tvActividades.SelectedItem.FullPath = tvActividades.SelectedItem.Text Then
    frmNodes.prRaiz = 0
Else
    frmNodes.prRaiz = Mid(tvActividades.SelectedItem.Parent.Key, 2)
End If
frmNodes.prLlave = Mid(Me.tvActividades.SelectedItem.Key, 2)
frmNodes.prCodigo = Left(Me.tvActividades.SelectedItem.Text, InStr(1, Me.tvActividades.SelectedItem.Text, ".") - 1)
frmNodes.prActividad = Mid(Me.tvActividades.SelectedItem.Text, InStr(1, Me.tvActividades.SelectedItem.Text, ".") + 3)
frmNodes.prCliente = chkPago(1).Value
frmNodes.Show
End Sub

Private Sub cmdImprimir_Click()
rptActividades.Show

End Sub

'Private Sub Form_Load()
'' Este código crea un árbol con 3 objetos Node.
'   TreeView1.Style = tvwTreelinesPlusMinusText ' Estilo 6.
'   TreeView1.LineStyle = tvwRootLines    'Estilo de línea 1.
'
'   ' Agrega varios objetos Node.
'   Dim nodX As Node    ' Crea variable.
'
'   Set nodX = TreeView1.Nodes.Add(, , "r", "Raíz")
'   Set nodX = TreeView1.Nodes.Add("r", tvwChild, "c1", "Secundario 1")
'
'   nodX.EnsureVisible ' Muestra todos los nodos.
'   Set nodX = TreeView1.Nodes.Add("c1", tvwChild, "c2", "Secundario 2")
'   Set nodX = TreeView1.Nodes.Add("c1", tvwChild, "c3", "Secundario 3")
'   nodX.EnsureVisible ' Muestra todos los nodos.
'End Sub
'
'Private Sub TreeView1_NodeClick(ByVal Node As Node)
'   ' Si el nodo tiene secundarios, se muestra el texto
'   ' del Node secundario.
'   If Node.Children Then
'      Caption = Node.Child.Text
'   End If
'End Sub


Private Sub Form_Activate()
If modal Then frmNodes.SetFocus
If Not modal Then Call tvActividades_NodeClick(Me.tvActividades.SelectedItem)

MDIPrimero.Skin1.ApplySkin hWnd
End Sub

Private Sub Form_Load()
On Error GoTo errload
Dim rs2 As New ADODB.Recordset

Me.Top = (MDIPrimero.ScaleHeight / 2) - (Me.Height / 2)
Me.Left = (MDIPrimero.ScaleWidth / 2) - (Me.Width / 2)

If cnx.State = adStateClosed Then
'    sql = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PRUEBA;Data Source=WEBMASTER\SQL2005"
    cnx.ConnectionString = Conexion
    cnx.Open
End If

sql = "SELECT [Llave], [Raiz], [Codigo], [Actividad], [PagaCliente] " & _
        "From [dbo].[Actividades]"
With rs
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

sql = "SELECT [Llave], [Raiz], [Codigo], [Actividad] " & _
        "From [dbo].[Actividades] " & _
        "Where Raiz = 0 " & _
        "order by codigo"
        
With rs2
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

If Not rs2.eof Then
    rs2.MoveFirst
    Do While Not rs2.eof
        tvActividades.Nodes.Add , tvwLast, "A" & Trim(Str(rs2!llave)), Trim(rs2!codigo) & ".- " & Trim(rs2!actividad), 1, 2
        Call addchild(rs2!llave)
        rs2.MoveNext
    Loop
End If

Set rs2 = Nothing
Exit Sub
errload:
    MsgBox Err.Description
End Sub

Private Sub addchild(llave As Integer)
Dim rs2 As New ADODB.Recordset

sql = "SELECT [Llave], [Raiz], [Codigo], [Actividad] " & _
        "From [dbo].[Actividades] " & _
        "Where Raiz = " & llave & _
        " order by codigo"

With rs2
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

If Not rs2.eof Then
    rs2.MoveFirst
    Do While Not rs2.eof
        tvActividades.Nodes.Add "A" & Trim(Str(rs2!raiz)), tvwChild, "A" & Trim(Str(rs2!llave)), Trim(rs2!codigo) & ".- " & Trim(rs2!actividad), 1, 2
        Call addchild(rs2!llave)
        rs2.MoveNext
    Loop
End If
Set rs2 = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set cnx = Nothing
Set rs = Nothing
Set rs3 = Nothing
Set rs4 = Nothing
End Sub

Public Property Let prModal(ByVal val As Boolean)
modal = val
End Property

Private Sub tdbgCuentas_FilterChange()
'Gets called when an action is performed on the filter bar
Dim col As TrueOleDBGrid80.Column
Dim cols As TrueOleDBGrid80.Columns

'On Error GoTo errHandler
On Error Resume Next
Set cols = tdbgCuentas.Columns
Dim c As Integer

c = tdbgCuentas.col
tdbgCuentas.HoldFields
sql = rs3.Filter
rs3.Filter = getFilter(col, cols)

tdbgCuentas.col = c
tdbgCuentas.EditActive = True
End Sub

Private Function getFilter(col As TrueOleDBGrid80.Column, cols As TrueOleDBGrid80.Columns) As String
'Creates the SQL statement in adodc1.recordset.filter
'and only filters text currently. It must be modified to
'filter other data types.
Dim tmp As String
Dim n As Integer
Dim X As Integer

For Each col In cols
    If Trim(col.FilterText) <> "" Then
        n = n + 1
        If n > 1 Then tmp = tmp & " AND "
        Select Case rs3.Fields(X).Type
        Case adVarWChar, adWChar: tmp = tmp & col.DataField & " LIKE '%" & col.FilterText & "%'"
        Case adInteger, adNumeric: tmp = tmp & col.DataField & " = " & col.FilterText
        Case adDBTimeStamp: tmp = tmp & col.DataField & " = #" & col.FilterText & "#"
        End Select
    End If
    X = X + 1
Next col
getFilter = tmp

End Function

Private Sub tvActividades_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo errclick
If tvActividades.Nodes.Count = 0 Then _
    MsgBox "Agregue actividades para continuar con las operaciones", vbInformation: Exit Sub
sql = "llave = " & Mid(tvActividades.SelectedItem.Key, 2)
rs.Requery
rs.Find sql
Me.chkPago(0).Value = IIf(rs!pagacliente = 0, True, False)
Me.chkPago(1).Value = IIf(rs!pagacliente = 0, False, True)

sql = "SELECT [TipoNomina] AS ID, N.[Nomina], [Cuenta] " & _
        "FROM [dbo].[ActividadesCuenta] C INNER JOIN [dbo].[TipoNomina] N ON C.[TipoNomina]= N.[CodTipoNomina] " & _
        "WHERE codigo = " & Left(tvActividades.SelectedItem.Text, InStr(1, tvActividades.SelectedItem.Text, ".") - 1)
        
If tvActividades.SelectedItem.FullPath = tvActividades.SelectedItem.Text Then
    sql = sql & " AND raiz = 0"
Else
    sql = sql & " AND raiz = " & Mid(tvActividades.SelectedItem.Parent.Key, 2)
End If

With rs3
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

Me.tdbgCuentas.DataSource = rs3

Exit Sub
errclick:
    MsgBox Err.Description

End Sub
