VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form frmProduccionReal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrador de Horas Laborales"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   11895
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
      Height          =   735
      Left            =   240
      TabIndex        =   27
      Top             =   8520
      Width           =   6015
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Ver Reporte"
         Height          =   375
         Left            =   4680
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   375
         Left            =   720
         TabIndex        =   28
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   81854465
         CurrentDate     =   40799
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProduccionReal.frx":0000
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   375
         Left            =   3000
         TabIndex        =   30
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   81854465
         CurrentDate     =   40799
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   2400
         OleObjectBlob   =   "frmProduccionReal.frx":0068
         TabIndex        =   31
         Top             =   240
         Width           =   1575
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
      Left            =   10440
      TabIndex        =   13
      Top             =   8520
      Width           =   1215
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Height          =   8175
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   11415
      Begin VB.TextBox txtCodPlantacion 
         Height          =   285
         Left            =   1800
         TabIndex        =   36
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtCodFinca 
         Height          =   285
         Left            =   1800
         TabIndex        =   35
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtCodActividad 
         Height          =   285
         Left            =   1200
         TabIndex        =   34
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtCodEmpleado 
         Height          =   285
         Left            =   1200
         TabIndex        =   33
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtPlantacion 
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1320
         Width           =   7575
      End
      Begin VB.TextBox txtCantHoras 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtFinca 
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   960
         Width           =   7575
      End
      Begin VB.TextBox txtActividad 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   8175
      End
      Begin VB.TextBox txtNomina 
         Height          =   285
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   4935
      End
      Begin VB.CommandButton cmdBorrar 
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
         Left            =   7560
         TabIndex        =   9
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdAgregar 
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
         Left            =   7080
         TabIndex        =   8
         Tag             =   "1"
         Top             =   1920
         Width           =   375
      End
      Begin TrueOleDBGrid80.TDBGrid tdbgProd 
         Height          =   5175
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   9128
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
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
      Begin MSComCtl2.DTPicker dtpFechaInicio 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   81854465
         CurrentDate     =   40799
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmProduccionReal.frx":00D0
         TabIndex        =   14
         Top             =   1680
         Width           =   1575
      End
      Begin TrueOleDBList80.TDBCombo tdbcFinca 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   960
         Visible         =   0   'False
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
         _PropDict       =   $"frmProduccionReal.frx":0138
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=167,.bold=0,.fontsize=825,.italic=0"
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmProduccionReal.frx":01E2
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin TrueOleDBList80.TDBCombo tdbcPlantacion 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
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
         _PropDict       =   $"frmProduccionReal.frx":024A
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=167,.bold=0,.fontsize=825,.italic=0"
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmProduccionReal.frx":02F4
         TabIndex        =   16
         Top             =   1320
         Width           =   1575
      End
      Begin TrueOleDBList80.TDBCombo tdbcEmpleado 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         _PropDict       =   $"frmProduccionReal.frx":0374
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=167,.bold=0,.fontsize=825,.italic=0"
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmProduccionReal.frx":041E
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   3600
         OleObjectBlob   =   "frmProduccionReal.frx":048C
         TabIndex        =   18
         Top             =   1680
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpEntrada 
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   81854466
         CurrentDate     =   40750
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   5280
         OleObjectBlob   =   "frmProduccionReal.frx":04F8
         TabIndex        =   19
         Top             =   1680
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpSalida 
         Height          =   375
         Left            =   5280
         TabIndex        =   7
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   81854466
         CurrentDate     =   40750
      End
      Begin TrueOleDBList80.TDBCombo tdbcActividad 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         _PropDict       =   $"frmProduccionReal.frx":0562
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=167,.bold=0,.fontsize=825,.italic=0"
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmProduccionReal.frx":060C
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   1920
         OleObjectBlob   =   "frmProduccionReal.frx":067C
         TabIndex        =   25
         Top             =   1680
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmProduccionReal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnx As New ADODB.Connection
Private rs As New ADODB.Recordset
Private rs2 As New ADODB.Recordset
Private rs3 As New ADODB.Recordset
Private rs4 As New ADODB.Recordset
Private rs5 As New ADODB.Recordset
Private sql As String

Private Sub cmdAgregar_Click()
On Error GoTo errAgregar
If Me.txtNombre.Text = "" Then
    MsgBox "Seleccione un Empleado", vbInformation
    Me.TxtCodEmpleado.SetFocus
    Exit Sub
ElseIf Me.txtActividad.Text = "" Then
    MsgBox "Seleccione una Actividad", vbInformation
    Me.txtCodActividad.SetFocus
    Exit Sub
ElseIf Me.txtFinca.Text = "" Then
    MsgBox "Seleccione una Finca", vbInformation
    Me.txtCodFinca.SetFocus
    Exit Sub
ElseIf Me.txtPlantacion.Text = "" Then
    MsgBox "Seleccione una Plantación - Año", vbInformation
    Me.txtCodPlantacion.SetFocus
    Exit Sub
ElseIf val(Me.txtCantHoras) < 0 Then
    MsgBox "Cantidad de Horas incorrectas...", vbInformation
    Me.txtCantHoras.SetFocus
    Exit Sub
ElseIf val(Me.txtCantHoras) = 0 And Me.dtpSalida.Value <= Me.dtpEntrada.Value Then
    MsgBox "La hora de salida debe ser mayor a la hora de entrada", vbInformation
    Me.dtpSalida = DateAdd("h", 1, Me.dtpEntrada)
    Me.dtpSalida.SetFocus
    Exit Sub
End If

sql = "insert into _ActividadesProduccion (IdActividad, CodEmpleado, IdFincaPlantacion, FechaRegistro, Fecha, " & _
      IIf(val(Me.txtCantHoras.Text) <= 0, "HEntrada, HSalida, CantidadHoras, ", "CantidadHoras, ") & "HExtras, Eliminar) " & _
      "values(" & rs2!IdActividad & ", " & rs5!CodEmpleado & "," & val(Me.txtCodPlantacion.Text) & ", '" & Format$(Now, "yyyyMMdd") & "', '" & Format$(Me.DtpFechaInicio.Value, "yyyyMMdd") & "', " & _
      IIf(val(Me.txtCantHoras.Text) <= 0, "'" & Format$(Me.dtpEntrada.Value, "yyyyMMdd hh:mm:ss") & "', '" & Format$(Me.dtpSalida.Value, "yyyyMMdd hh:mm:ss") & "', 0, ", val(Me.txtCantHoras.Text) & ", ") & " 0, 1) "
cnx.Execute sql


rs.Requery
Me.tdbgProd.ReBind
Me.tdbgProd.Refresh

With tdbgProd
    .DataSource = rs
    .Columns(1).Width = 700
    .Columns(2).Width = 2700
    .Columns(3).Width = 1000
    .Columns(4).Width = 550
    .Columns(5).Width = 1200
    .Columns(6).Width = 1200
    .Columns(7).Width = 1000
    .Columns(8).Visible = False
    .Columns(5).ValueItems.Translate = True
    .Columns(6).ValueItems.Translate = True
End With

Dim item As New TrueOleDBGrid80.ValueItem
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
    If Not IsNull(rs!HEntrada) Then
        item.Value = rs!HEntrada
        item.DisplayValue = Format$(rs!HEntrada, "HH:MM:SS AMPM")
        Me.tdbgProd.Columns(5).ValueItems.Add item
    End If
    If Not IsNull(rs!HSalida) Then
        item.Value = rs!HSalida
        item.DisplayValue = Format$(rs!HSalida, "HH:MM:SS AMPM")
        Me.tdbgProd.Columns(6).ValueItems.Add item
    End If
    rs.MoveNext
Loop
If rs.RecordCount > 0 Then rs.MoveFirst

'rs.Find "Código = " & Trim(Me.tdbcActRaiz.Text), 0

Me.TxtCodEmpleado.Text = ""
Me.txtCodActividad.Text = ""
Me.txtCodFinca.Text = ""
Me.txtCodPlantacion.Text = ""
Me.txtNombre.Text = ""
Me.txtNomina.Text = ""
Me.txtActividad.Text = ""
Me.txtFinca.Text = ""
Me.txtPlantacion.Text = ""
Me.txtCantHoras.Text = ""
Me.DtpFechaInicio = Now
Me.dtpEntrada = Me.DtpFechaInicio.Value
Me.dtpSalida = Me.DtpFechaInicio.Value
Me.TxtCodEmpleado.SetFocus

Exit Sub
errAgregar:
    MsgBox Err.Description
End Sub

Private Sub cmdBorrar_Click()
On Error GoTo errcta
If rs.EOF = rs.BOF And rs.EOF = True Then MsgBox "Operación Cancelada, no existen registros", vbInformation: Exit Sub
If MsgBox("¿Desea eliminar el registro?", vbYesNo) = vbYes Then
    sql = "delete from _ActividadesProduccion where IdActProduccion = " & Trim(rs!IdActProduccion)
    cnx.Execute sql
    MsgBox "Registro eliminado", vbInformation
        
    rs.Requery
    Me.tdbgProd.ReBind
    Me.tdbgProd.Refresh
    
    With tdbgProd
        .DataSource = rs
        .Columns(1).Width = 700
        .Columns(2).Width = 2700
        .Columns(3).Width = 1000
        .Columns(4).Width = 1200
        .Columns(5).Width = 1200
        .Columns(7).Width = 1000
        .Columns(8).Visible = False
        .Columns(4).ValueItems.Translate = True
        .Columns(5).ValueItems.Translate = True
    End With
    
    Dim item As New TrueOleDBGrid80.ValueItem
    If rs.RecordCount > 0 Then rs.MoveFirst
    Do While Not rs.EOF
        If Not IsNull(rs!HEntrada) Then
            item.Value = rs!HEntrada
            item.DisplayValue = Format$(rs!HEntrada, "HH:MM:SS AMPM")
            Me.tdbgProd.Columns(4).ValueItems.Add item
        End If
        If Not IsNull(rs!HSalida) Then
            item.Value = rs!HSalida
            item.DisplayValue = Format$(rs!HSalida, "HH:MM:SS AMPM")
            Me.tdbgProd.Columns(5).ValueItems.Add item
        End If
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then rs.MoveFirst

End If

Exit Sub
errcta:
    MsgBox Err.Description
End Sub

Private Sub cmdCerrar_Click()
Unload Me

End Sub

Private Sub cmdReporte_Click()
On Error GoTo errRep

'rptProduccion.prFechaIni = Me.dtpDesde.Value
'rptProduccion.prFechaFin = Me.dtpHasta.Value
'
'rptProduccion.Show
'ArepColillasPago.AdoColillas.Source = SQlReportes
'ArepColillasPago.LblPeriodo.Caption = FechaIni & " al " & FechaFin
'ArepColillasPago.LblTitulo.Caption = Titulo
'ArepColillasPago.AdoColillas.ConnectionString = ConexionReporte
'ArepColillasPago.Show 1

Exit Sub
errRep:
    MsgBox Err.Description
    
End Sub

Private Sub dtpFechaInicio_Change()
dtpEntrada.Value = Me.DtpFechaInicio.Value
dtpSalida.Value = Me.DtpFechaInicio.Value
End Sub

Private Sub Form_Activate()
On Error Resume Next
MDIPrimero.Skin1.ApplySkin hWnd

End Sub

Private Sub Form_Load()
On Error GoTo errload
Dim X As Integer

Me.Top = (MDIPrimero.ScaleHeight / 2) - (Me.Height / 2)
Me.Left = (MDIPrimero.ScaleWidth / 2) - (Me.Width / 2)

Me.DtpFechaInicio.Value = Now
dtpEntrada.Value = Me.DtpFechaInicio.Value
dtpSalida.Value = Me.DtpFechaInicio.Value

If cnx.State = adStateClosed Then
'    sql = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PRUEBA;Data Source=WEBMASTER\SQL2005"
    cnx.ConnectionString = Conexion
    cnx.Open
End If
   
sql = "SELECT     CASE WHEN Nombre2 IS NULL " & _
      "               THEN CodEmpleado1 + ' - ' + Nombre1 + ' ' + Apellido1 + ' ' + Apellido2 ELSE CodEmpleado1 + ' - ' + Nombre1 + ' ' + Nombre2 + ' ' + Apellido1 + ' ' + Apellido2 END AS [Nombre Completo], " & _
      "               RTRIM(_Actividades.Sufijo) + RTRIM(_Actividades.Codigo) AS Código, _Actividades.Actividad, " & _
      "               _ActividadesProduccion.Fecha, _ActividadesProduccion.CantidadHoras as Horas, _ActividadesProduccion.HEntrada, " & _
      "               _ActividadesProduccion.HSalida, _Finca.Finca, RTRIM(_Plantacion.Plantacion) + ' - ' + CAST(_FincaPlantacion.Anio AS nchar(4)) AS Plantación, _ActividadesProduccion.IdActProduccion " & _
      "FROM         _ActividadesProduccion INNER JOIN " & _
      "               Empleado ON _ActividadesProduccion.CodEmpleado = Empleado.CodEmpleado INNER JOIN " & _
      "               _Actividades ON _ActividadesProduccion.IdActividad = _Actividades.IdActividad INNER JOIN " & _
      "               _FincaPlantacion ON _ActividadesProduccion.IdFincaPlantacion = _FincaPlantacion.IdFincaPlantacion INNER JOIN " & _
      "               _Finca ON _FincaPlantacion.IdFinca = _Finca.IdFinca INNER JOIN " & _
      "               _Plantacion ON _FincaPlantacion.IdPlantacion = _Plantacion.IdPlantacion "
      
With rs
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

With tdbgProd
    .DataSource = rs
    .Columns(1).Width = 700
    .Columns(2).Width = 2700
    .Columns(3).Width = 1000
    .Columns(4).Width = 550
    .Columns(5).Width = 1200
    .Columns(6).Width = 1200
    .Columns(7).Width = 1000
    .Columns(8).Visible = False
    .Columns(5).ValueItems.Translate = True
    .Columns(6).ValueItems.Translate = True
End With

Dim item As New TrueOleDBGrid80.ValueItem
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
    If Not IsNull(rs!HEntrada) Then
        item.Value = rs!HEntrada
        item.DisplayValue = Format$(rs!HEntrada, "HH:MM:SS AMPM")
        Me.tdbgProd.Columns(5).ValueItems.Add item
    End If
    If Not IsNull(rs!HSalida) Then
        item.Value = rs!HSalida
        item.DisplayValue = Format$(rs!HSalida, "HH:MM:SS AMPM")
        Me.tdbgProd.Columns(6).ValueItems.Add item
    End If
    rs.MoveNext
Loop
If rs.RecordCount > 0 Then rs.MoveFirst

sql = "SELECT RTRIM(Sufijo) + RTRIM(Codigo) AS Codigo, Actividad , IdActividad " & _
        "FROM _Actividades WHERE LEN(RTRIM(Sufijo) + RTRIM(Codigo)) > 3 "
With rs2
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

Me.tdbcActividad.RowSource = rs2
Me.tdbcActividad.BoundColumn = "IdActividad"
Me.tdbcActividad.Refresh
Me.tdbcActividad.Columns(0).Width = 800
Me.tdbcActividad.Columns(2).Visible = False
Me.tdbcActividad.Text = ""

sql = "SELECT IdFinca, Finca from _finca"
With rs3
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

Me.tdbcFinca.RowSource = rs3
Me.tdbcFinca.BoundColumn = "IdFinca"
Me.tdbcFinca.Refresh
'Me.tdbcFinca.Columns(1).Visible = False
Me.tdbcFinca.Text = ""


sql = "SELECT  CodEmpleado1 as Código, 'Nombre Completo' = case when Nombre2 is null then Nombre1 + ' ' + Apellido1 + ' ' + Apellido2 " & _
                  "else Nombre1 + ' ' + Nombre2 + ' ' + Apellido1 + ' ' + Apellido2 " & _
            "end, CodEmpleado, tiponomina.nomina " & _
        "FROM         Empleado inner join tiponomina on empleado.CodTipoNomina = tiponomina.CodTipoNomina"
With rs5
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

Me.tdbcEmpleado.RowSource = rs5
Me.tdbcEmpleado.BoundColumn = "CodEmpleado"
Me.tdbcEmpleado.Refresh
Me.tdbcEmpleado.Columns(2).Visible = False
Me.tdbcEmpleado.Columns(3).Visible = False
Me.tdbcEmpleado.Text = ""


'para tdropdown


'With tddNiveles
'    .DataSource = rs2
'    .ListField = "IdActividad"
'End With

Exit Sub
errload:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set cnx = Nothing
Set rs = Nothing
Set rs2 = Nothing
Set rs3 = Nothing
Set rs4 = Nothing
Set rs5 = Nothing

End Sub

Private Sub tdbcActividad_ItemChange()
On Error GoTo errItem
rs2.Bookmark = Me.tdbcActividad.Bookmark
Me.txtActividad.Text = rs2!actividad

Exit Sub
errItem:
    MsgBox Err.Description

End Sub

Private Sub tdbcEmpleado_ItemChange()
On Error GoTo errItem
rs5.Bookmark = Me.tdbcEmpleado.Bookmark
Me.txtNombre.Text = rs5![Nombre Completo]
Me.txtNomina.Text = rs5!nomina

Exit Sub
errItem:
    MsgBox Err.Description

End Sub

Private Sub tdbcFinca_ItemChange()
On Error GoTo errItem
rs3.Bookmark = Me.tdbcFinca.Bookmark
Me.txtFinca.Text = rs3!Finca

sql = "SELECT    _FincaPlantacion.IdFincaPlantacion, _Plantacion.Plantacion as Plantación, _FincaPlantacion.Anio as Año " & _
        "FROM         _FincaPlantacion INNER JOIN _Plantacion ON _FincaPlantacion.IdPlantacion = _Plantacion.IdPlantacion " & _
        "where _FincaPlantacion.IdFinca = " & Trim(Me.tdbcFinca.BoundText)
With rs4
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

Me.tdbcPlantacion.RowSource = rs4
Me.tdbcPlantacion.BoundColumn = "IdFincaPlantacion"
Me.tdbcPlantacion.Refresh
'Me.tdbcPlantacion.Columns(2).Visible = False
Me.tdbcPlantacion.Text = ""
Exit Sub
errItem:
    MsgBox Err.Description
    
End Sub

Private Sub tdbcPlantacion_ItemChange()
On Error GoTo errItem
rs4.Bookmark = Me.tdbcPlantacion.Bookmark
Me.txtPlantacion.Text = rs4!Plantación & " - " & rs4!Año

Exit Sub
errItem:
    Me.txtPlantacion.Text = ""

End Sub

Private Sub tdbgProd_FilterChange()
'Gets called when an action is performed on the filter bar
Dim col As TrueOleDBGrid80.Column
Dim cols As TrueOleDBGrid80.Columns

'On Error GoTo errHandler
On Error Resume Next
Set cols = tdbgProd.Columns
Dim c As Integer

c = tdbgProd.col
tdbgProd.HoldFields
sql = rs.Filter
rs.Filter = getFilter(col, cols)

tdbgProd.col = c
tdbgProd.EditActive = True
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
        Select Case rs.Fields(X).Type
        Case adVarWChar, adVarChar: tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
        Case adInteger, adNumeric: tmp = tmp & "[" & col.DataField & "] = " & col.FilterText
        Case adDBTimeStamp: tmp = tmp & "[" & col.DataField & "] = #" & col.FilterText & "#"
        End Select
    End If
    X = X + 1
Next col
getFilter = tmp

End Function

Private Sub txtCantHoras_Change()
If val(Me.txtCantHoras) > 0 Then
    Me.dtpEntrada.Enabled = False
    Me.dtpSalida.Enabled = False
Else
    Me.dtpEntrada.Enabled = True
    Me.dtpSalida.Enabled = True
End If

End Sub

Private Sub txtCodActividad_Change()
On Error GoTo errAct
rs2.Find "Codigo = '" & Trim(Me.txtCodActividad.Text) & "'", , , 1
Me.txtActividad.Text = rs2!actividad

Exit Sub
errAct:
    Me.txtActividad.Text = ""
End Sub

Private Sub txtCodEmpleado_Change()
On Error GoTo errEmp
rs5.Find "Código = '" & Trim(Me.TxtCodEmpleado.Text) & "'", , , 1
Me.txtNombre.Text = rs5![Nombre Completo]
Me.txtNomina.Text = rs5!nomina

Exit Sub
errEmp:
    Me.txtNombre.Text = ""
    Me.txtNomina.Text = ""
    
End Sub

Private Sub txtCodFinca_Change()
On Error GoTo errFin
rs3.Find "IdFinca = " & Trim(Me.txtCodFinca.Text), , , 1
Me.txtFinca.Text = rs3!Finca

Exit Sub
errFin:
    Me.txtFinca.Text = ""
End Sub


Private Sub txtCodPlantacion_Change()
On Error GoTo errPlan
rs4.Find "IdFincaPlantacion = " & Trim(Me.txtCodPlantacion.Text), , , 1
Me.txtPlantacion.Text = rs4!Plantación & " - " & rs4!Año


Exit Sub
errPlan:
    Me.txtPlantacion.Text = ""
End Sub

Private Sub txtFinca_Change()
On Error GoTo errItem
If Trim(txtFinca.Text) <> "" Then
    sql = "SELECT    _FincaPlantacion.IdFincaPlantacion, _Plantacion.Plantacion as Plantación, _FincaPlantacion.Anio as Año " & _
            "FROM         _FincaPlantacion INNER JOIN _Plantacion ON _FincaPlantacion.IdPlantacion = _Plantacion.IdPlantacion " & _
            "where _FincaPlantacion.IdFinca = " & Trim(Me.txtCodFinca.Text)
    With rs4
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .Open sql, cnx, adOpenDynamic, adLockOptimistic
    End With
    
    Me.tdbcPlantacion.RowSource = rs4
    Me.tdbcPlantacion.BoundColumn = "IdFincaPlantacion"
    Me.tdbcPlantacion.Refresh
    'Me.tdbcPlantacion.Columns(2).Visible = False
    Me.tdbcPlantacion.Text = ""
End If

Exit Sub
errItem:
    MsgBox Err.Description
    
End Sub

