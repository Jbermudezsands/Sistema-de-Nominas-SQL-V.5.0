VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmReportes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmSeleccion 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   3360
      TabIndex        =   33
      Top             =   2040
      Visible         =   0   'False
      Width           =   3735
      Begin XtremeSuiteControls.RadioButton OptProduccionPagada 
         Height          =   390
         Left            =   1320
         TabIndex        =   34
         Top             =   240
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   688
         _StockProps     =   79
         Caption         =   "Produccion Pagada"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptProduccionNoPagada 
         Height          =   390
         Left            =   2520
         TabIndex        =   35
         Top             =   240
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   688
         _StockProps     =   79
         Caption         =   "Produccion No Pagada"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptProduccionCompleta 
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Produccion Completa"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Departamento"
      Height          =   2415
      Left            =   2760
      TabIndex        =   65
      Top             =   840
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton Command3 
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
         Left            =   5280
         Picture         =   "FrmReportes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command2 
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
         Left            =   5280
         Picture         =   "FrmReportes.frx":014E
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   720
         Width           =   375
      End
      Begin TrueOleDBList80.TDBCombo TDBDepartamentoIni 
         Bindings        =   "FrmReportes.frx":029C
         Height          =   315
         Left            =   600
         TabIndex        =   68
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
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
         _PropDict       =   $"FrmReportes.frx":02BA
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
      Begin TrueOleDBList80.TDBCombo TDBDepartamentoFin 
         Bindings        =   "FrmReportes.frx":0364
         Height          =   315
         Left            =   600
         TabIndex        =   69
         Top             =   720
         Width           =   4695
         _ExtentX        =   8281
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
         ListField       =   "CodDepartamento"
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
         _PropDict       =   $"FrmReportes.frx":0382
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmReportes.frx":042C
         TabIndex        =   70
         Top             =   360
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmReportes.frx":0496
         TabIndex        =   71
         Top             =   720
         Width           =   375
      End
   End
   Begin MSAdodcLib.Adodc AdoDeducciones 
      Height          =   375
      Left            =   5400
      Top             =   6240
      Width           =   2760
      _ExtentX        =   4868
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
      Caption         =   "AdoDeducciones"
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
   Begin TrueOleDBList80.TDBCombo TDBCombo2 
      Bindings        =   "FrmReportes.frx":04FA
      Height          =   315
      Left            =   240
      TabIndex        =   62
      Top             =   4080
      Width           =   2535
      _ExtentX        =   4471
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
      _PropDict       =   $"FrmReportes.frx":0517
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=200,.bold=0,.fontsize=825,.italic=0"
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
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   3480
      TabIndex        =   61
      Top             =   3720
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc AdoAuxiliar 
      Height          =   375
      Left            =   6720
      Top             =   5640
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
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton Command1 
      Height          =   255
      Left            =   6120
      TabIndex        =   60
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblTitulo 
      Height          =   375
      Left            =   3360
      OleObjectBlob   =   "FrmReportes.frx":05C1
      TabIndex        =   53
      Top             =   840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6840
      TabIndex        =   39
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton CmdVerreporte 
      Caption         =   "Ver Reporte"
      Height          =   375
      Left            =   240
      TabIndex        =   37
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Todos los Empleados"
      Height          =   375
      Left            =   3000
      TabIndex        =   32
      Top             =   3360
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSAdodcLib.Adodc AdoSuspenciones 
      Height          =   375
      Left            =   360
      Top             =   9360
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
      Caption         =   "AdoSuspenciones"
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
   Begin MSAdodcLib.Adodc AdoNuevoIngreso 
      Height          =   375
      Left            =   1200
      Top             =   9960
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
      Caption         =   "AdoNuevoIngreso"
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
   Begin MSAdodcLib.Adodc AdoVacaciones 
      Height          =   330
      Left            =   1200
      Top             =   9600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
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
      Caption         =   "AdoVacaciones"
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
   Begin MSAdodcLib.Adodc AdoTarifa 
      Height          =   495
      Left            =   1920
      Top             =   9480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   873
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
      Caption         =   "AdoTarifa"
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
   Begin MSAdodcLib.Adodc AdoBajas 
      Height          =   375
      Left            =   1320
      Top             =   10920
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
      Caption         =   "AdoBajas"
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
   Begin MSAdodcLib.Adodc AdoDatosEmpresa 
      Height          =   375
      Left            =   1680
      Top             =   10320
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
      Caption         =   "AdoDatosEmpresa"
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   2415
      Left            =   2760
      TabIndex        =   8
      Top             =   840
      Width           =   5775
      Begin VB.CommandButton CmdBuscarEmpleadoFin 
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
         Left            =   5280
         Picture         =   "FrmReportes.frx":0641
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton CmdBuscarEmpleadoIni 
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
         Left            =   5280
         Picture         =   "FrmReportes.frx":078F
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   720
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmReportes.frx":08DD
         TabIndex        =   40
         Top             =   240
         Width           =   495
      End
      Begin TrueOleDBList80.TDBCombo DBPeriodos 
         Bindings        =   "FrmReportes.frx":0945
         Height          =   315
         Left            =   1320
         TabIndex        =   23
         Top             =   2640
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
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
         _PropDict       =   $"FrmReportes.frx":095E
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
      Begin TrueOleDBList80.TDBCombo DBTipoNominas 
         Bindings        =   "FrmReportes.frx":0A08
         Height          =   315
         Left            =   1080
         TabIndex        =   22
         Top             =   1440
         Width           =   4575
         _ExtentX        =   8070
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
         _PropDict       =   $"FrmReportes.frx":0A1E
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
      Begin TrueOleDBList80.TDBCombo DataCombo1 
         Bindings        =   "FrmReportes.frx":0AC8
         Height          =   315
         Left            =   600
         TabIndex        =   18
         Top             =   720
         Width           =   4695
         _ExtentX        =   8281
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
         _PropDict       =   $"FrmReportes.frx":0AE2
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
      Begin VB.TextBox TxtNumeros 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         TabIndex        =   15
         Top             =   1560
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtNumNominas 
         Height          =   375
         Left            =   4080
         TabIndex        =   14
         Top             =   1680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmReportes.frx":0B8C
         Left            =   600
         List            =   "FrmReportes.frx":0BB4
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmReportes.frx":0C1D
         Left            =   3600
         List            =   "FrmReportes.frx":0C45
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DBCAño 
         Bindings        =   "FrmReportes.frx":0CAE
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker TxtFecha1 
         Height          =   375
         Left            =   720
         TabIndex        =   16
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   16777215
         Format          =   17039361
         CurrentDate     =   37257
      End
      Begin MSComCtl2.DTPicker TxtFecha2 
         Height          =   375
         Left            =   3600
         TabIndex        =   17
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   16777215
         Format          =   17039361
         CurrentDate     =   37257
      End
      Begin TrueOleDBList80.TDBCombo DataCombo2 
         Bindings        =   "FrmReportes.frx":0CC3
         Height          =   315
         Left            =   600
         TabIndex        =   19
         Top             =   1080
         Width           =   4695
         _ExtentX        =   8281
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
         _PropDict       =   $"FrmReportes.frx":0CDD
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
      Begin MSDataListLib.DataCombo DBAño2 
         Bindings        =   "FrmReportes.frx":0D87
         Height          =   315
         Left            =   4680
         TabIndex        =   20
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "FrmReportes.frx":0D9C
         TabIndex        =   41
         Top             =   240
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmReportes.frx":0E04
         TabIndex        =   42
         Top             =   720
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmReportes.frx":0E6E
         TabIndex        =   43
         Top             =   1080
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmReportes.frx":0ED2
         TabIndex        =   44
         Top             =   1440
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmReportes.frx":0F46
         TabIndex        =   45
         Top             =   1800
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "FrmReportes.frx":0FAE
         TabIndex        =   46
         Top             =   1800
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmReportes.frx":1016
         TabIndex        =   47
         Top             =   2640
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   2760
      TabIndex        =   24
      Top             =   1080
      Visible         =   0   'False
      Width           =   5535
      Begin TrueOleDBList80.TDBCombo TDBCombo1 
         Bindings        =   "FrmReportes.frx":1082
         Height          =   315
         Left            =   1200
         TabIndex        =   31
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
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
         _PropDict       =   $"FrmReportes.frx":1098
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=248,.bold=0,.fontsize=825,.italic=0"
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
      Begin VB.TextBox TxtNumero 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3600
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxtNNomina 
         Height          =   375
         Left            =   3600
         TabIndex        =   26
         Top             =   720
         Width           =   615
      End
      Begin MSDataListLib.DataCombo DBAño 
         Bindings        =   "FrmReportes.frx":1142
         Height          =   315
         Left            =   600
         TabIndex        =   25
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker MtxtFechaini 
         Height          =   375
         Left            =   840
         TabIndex        =   28
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   16777215
         Format          =   17039361
         CurrentDate     =   37257
      End
      Begin MSComCtl2.DTPicker MtxtFecha 
         Height          =   375
         Left            =   3480
         TabIndex        =   29
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   16777215
         Format          =   17039361
         CurrentDate     =   37257
      End
      Begin MSDataListLib.DataCombo DBComboPeriodo 
         Bindings        =   "FrmReportes.frx":1157
         Height          =   315
         Left            =   2520
         TabIndex        =   30
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin ACTIVESKINLibCtl.SkinLabel DbTipoNomina 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmReportes.frx":1170
         TabIndex        =   48
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmReportes.frx":11E4
         TabIndex        =   49
         Top             =   720
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "FrmReportes.frx":124A
         TabIndex        =   50
         Top             =   720
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmReportes.frx":12B6
         TabIndex        =   51
         Top             =   1200
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   2880
         OleObjectBlob   =   "FrmReportes.frx":131E
         TabIndex        =   52
         Top             =   1200
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblInicio 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmReportes.frx":1386
         TabIndex        =   56
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
      End
      Begin TrueOleDBList80.TDBCombo TDBCodigo1 
         Bindings        =   "FrmReportes.frx":13F0
         Height          =   315
         Left            =   840
         TabIndex        =   57
         Top             =   1680
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
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
         _PropDict       =   $"FrmReportes.frx":1410
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
      Begin ACTIVESKINLibCtl.SkinLabel LblFinal 
         Height          =   255
         Left            =   3000
         OleObjectBlob   =   "FrmReportes.frx":14BA
         TabIndex        =   58
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin TrueOleDBList80.TDBCombo TDBCodigo2 
         Bindings        =   "FrmReportes.frx":151E
         Height          =   315
         Left            =   3480
         TabIndex        =   59
         Top             =   1680
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
         _PropDict       =   $"FrmReportes.frx":153E
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
   Begin MSAdodcLib.Adodc AdoAño 
      Height          =   495
      Left            =   3720
      Top             =   10200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
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
      Caption         =   "AdoAño"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2400
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SmartButtonProject.SmartButton CmdExportar 
      Height          =   975
      Left            =   3480
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      Caption         =   "Exportar  CSV"
      Picture         =   "FrmReportes.frx":15E8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoEmpleado 
      Height          =   375
      Left            =   1680
      Top             =   9960
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
      Caption         =   "AdoEmpleado"
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
   Begin MSAdodcLib.Adodc AdoBusca 
      Height          =   375
      Left            =   1080
      Top             =   9600
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
   Begin MSAdodcLib.Adodc AdoTipo 
      Height          =   375
      Left            =   2640
      Top             =   9360
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
      Caption         =   "AdoTipo"
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
   Begin MSAdodcLib.Adodc AdoPeriodo 
      Height          =   375
      Left            =   720
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
      Caption         =   "AdoPeriodo"
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
   Begin MSAdodcLib.Adodc DtaNomSubsidio 
      Height          =   375
      Left            =   1440
      Top             =   10320
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
      Caption         =   "DtaNomSubsidio"
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
      Left            =   3840
      Top             =   9600
      Width           =   3495
      _ExtentX        =   6165
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
   Begin MSComCtl2.MonthView Mes 
      Height          =   2370
      Left            =   5040
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   17039361
      TitleBackColor  =   12632256
      TrailingForeColor=   12632256
      CurrentDate     =   37838
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   3000
      TabIndex        =   4
      Top             =   1200
      Width           =   3855
      Begin MSComCtl2.DTPicker DTFecha2 
         Height          =   290
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   17039361
         CurrentDate     =   37837
      End
      Begin MSComCtl2.DTPicker DTFecha1 
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   17039361
         CurrentDate     =   37837
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmReportes.frx":21BA
         TabIndex        =   54
         Top             =   240
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   2160
         OleObjectBlob   =   "FrmReportes.frx":2224
         TabIndex        =   55
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton CmdVerreporte1 
      Caption         =   "Ver Reporte"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   10920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox CmbReportes 
      Height          =   2400
      ItemData        =   "FrmReportes.frx":2288
      Left            =   120
      List            =   "FrmReportes.frx":228A
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton CmdRptOtros 
      Caption         =   "Personalizados"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton CmdSalir1 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   10920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin SmartButtonProject.SmartButton CmdExportarExcel 
      Height          =   975
      Left            =   4800
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      Caption         =   "Exportar Excel"
      Picture         =   "FrmReportes.frx":228C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CmdAcercade 
      Height          =   435
      Left            =   120
      TabIndex        =   21
      Top             =   240
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   767
      _Version        =   196610
      Font3D          =   2
      MarqueeStyle    =   4
      ForeColor       =   8388608
      MarqueeDelay    =   5
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "REPORTES"
      ButtonStyle     =   4
      AutoRepeat      =   -1  'True
   End
   Begin XtremeSuiteControls.ProgressBar Barra 
      Height          =   375
      Left            =   240
      TabIndex        =   36
      Top             =   4440
      Width           =   3135
      _Version        =   786432
      _ExtentX        =   5530
      _ExtentY        =   661
      _StockProps     =   93
      Scrolling       =   1
      Appearance      =   6
   End
   Begin MSAdodcLib.Adodc AdoReportes 
      Height          =   375
      Left            =   3840
      Top             =   6960
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
      Caption         =   "AdoReportes"
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
      Left            =   120
      Top             =   6720
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
   Begin MSAdodcLib.Adodc AdoEmpleadoActivo 
      Height          =   375
      Left            =   1440
      Top             =   6120
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
      Caption         =   "AdoEmpleadoActivo"
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
   Begin MSAdodcLib.Adodc AdoDepartamento 
      Height          =   375
      Left            =   4920
      Top             =   7800
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
      Caption         =   "AdoDepartamento"
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
Attribute VB_Name = "FrmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Function MesIni(MesLetra As String, Ano As Double) As Date
  Select Case MesLetra
     Case "Enero": MesIni = "01/01/" & Ano
     Case "Febrero": MesIni = "01/02/" & Ano
     Case "Marzo": MesIni = "01/03/" & Ano
     Case "Abril": MesIni = "01/04/" & Ano
     Case "Mayo": MesIni = "01/05/" & Ano
     Case "Junio": MesIni = "01/06/" & Ano
     Case "Julio": MesIni = "01/07/" & Ano
     Case "Agosto": MesIni = "01/08/" & Ano
     Case "Septiembre": MesIni = "01/09/" & Ano
     Case "Octubre": MesIni = "01/10/" & Ano
     Case "Noviembre": MesIni = "01/11/" & Ano
     Case "Diciembre": MesIni = "01/12/" & Ano

  End Select

End Function
Private Sub BDAño2_Click(Area As Integer)
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String
If Not Me.TxtNumero.Text = "" Then
    Me.AdoPeriodo.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & Me.TxtNumeros.Text & "')AND (año = '" & Me.DBCAño.Text & "')"
'    InputBox "", "", Me.AdoPeriodo.RecordSource
    Me.AdoPeriodo.Refresh
    Me.DBPeriodos.ListField = "Periodo"
End If

FMes (Combo3.Text)
Mes1 = Format(Nmes, "0#")
FMes (Combo4.Text)
Mes2 = Format(Nmes, "0#")
Año1 = val(Me.DBAño.Text)
CodTipoNomina = Me.TDBCombo1.Columns(0).Text


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.MtxtFechaini.Value = Me.AdoBusca.Recordset("Inicio")
 End If
 
Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.MtxtFecha.Value = Me.AdoBusca.Recordset("Final")
 End If
End Sub

Private Sub CmbReportes_Click()
    Me.Check1.Visible = False
    Me.Frame2.Visible = False
    Me.Frame3.Visible = False
    Me.Frame1.Visible = False
    Me.CmdExportar.Visible = False
    Me.CmdExportarExcel.Visible = False
    Me.Mes.Visible = False
    Me.CmdVerreporte.Visible = True
    Me.FrmSeleccion.Visible = False
    Me.LblInicio.Visible = False
    Me.LblFinal.Visible = False
    Me.TDBCodigo1.Visible = False
    Me.TDBCodigo2.Visible = False
    Me.txtCantidad.Visible = False
    Me.Frame4.Visible = False
    
Select Case CmbReportes.Text
  Case "Exportar Lista Empleados"
    Me.CmdExportarExcel.Visible = True
    Me.Frame3.Visible = True
    Me.CmdVerreporte.Visible = False
    Me.txtCantidad.Visible = False
    Me.TDBCombo1.Visible = False
    
    
 Case "Reporte Carnet Empleados"
   Me.Frame3.Visible = False
   Me.Frame4.Visible = True
   
 Case "Reporte Horas Extra"
        Me.Frame3.Visible = True
        Me.txtCantidad.Visible = False
 Case "Reporte GRAL INGRESOS"
    Me.Frame3.Visible = True
    Me.Check1.Visible = True
    Me.txtCantidad.Visible = False
 Case "Reporte Proyeccion Vacaciones"
    Me.Frame3.Visible = True
    Me.Combo1.Visible = False
    Me.Combo2.Visible = False
    Me.DataCombo1.Visible = False
    Me.DataCombo2.Visible = False
    Me.DbTipoNomina.Visible = False
    Me.DBTipoNominas.Visible = False
    Me.DBCAño.Visible = False
    Me.DBAño2.Visible = False
    Me.SkinLabel1.Visible = False
    Me.SkinLabel2.Visible = False
    Me.SkinLabel3.Visible = False
    Me.SkinLabel4.Visible = False
    Me.SkinLabel5.Visible = False
    Me.TxtFecha1.Value = Now
    Me.TxtFecha2.Value = Now
    Me.txtCantidad.Visible = False
 Case "Reporte Estimado Vacaciones"
     Me.Frame3.Visible = True
    Me.Check1.Visible = True
    Me.CmdExportarExcel.Visible = True
    Me.txtCantidad.Visible = False
    
 Case "Listado de Empleados FHM"
    Me.Frame3.Visible = True
    Me.Check1.Visible = False
    Me.CmdExportarExcel.Visible = False
    Me.txtCantidad.Visible = False
 
 Case "Analisis Produccion Resumen"
    Me.CmdExportar.Visible = True
    Me.Frame1.Visible = False
    Me.Mes.Visible = False
    Me.lblTitulo.Visible = False
    Me.Frame2.Visible = True
    Me.CmdExportarExcel.Visible = True
    Me.txtCantidad.Visible = False
 Case "Analisis Produccion"
    Me.CmdExportar.Visible = True
    Me.Frame1.Visible = False
    Me.Mes.Visible = False
    Me.lblTitulo.Visible = False
    Me.Frame2.Visible = True
    Me.CmdExportarExcel.Visible = True
    Me.LblInicio.Visible = True
    Me.LblFinal.Visible = True
    Me.TDBCodigo1.Visible = True
    Me.TDBCodigo2.Visible = True
    Me.txtCantidad.Visible = False
 Case "Reporte Dias Acumulados"
     Me.Frame3.Visible = True
    Me.Check1.Visible = True
    Me.CmdExportarExcel.Visible = True
    Me.txtCantidad.Visible = False
 Case "Reporte INSS E IR MENSUAL"
    Me.Frame3.Visible = True
    Me.Check1.Visible = True
    Me.CmdExportarExcel.Visible = True
    Me.txtCantidad.Visible = False
 Case "EXPORTACION INSS"
    Me.CmdExportarExcel.Visible = True
    Me.Frame3.Visible = True
    Me.CmdVerreporte.Visible = False
    Me.txtCantidad.Visible = False
 Case "Listado Maestro de Empleados"
'    Me.CmdExportar.Visible = True
    Me.Frame1.Visible = False
    Me.Mes.Visible = False
    Me.lblTitulo.Visible = False
    Me.Frame2.Visible = True
    Me.CmdExportarExcel.Visible = True
    Me.txtCantidad.Visible = False

 Case "Reporte Ir"
    Me.Frame3.Visible = True
    Me.CmdExportar.Visible = True
    Me.txtCantidad.Visible = False
    
 Case "Reporte Registro Vacaciones"
    Me.Frame3.Visible = True
    Me.Check1.Visible = False
    Me.CmdExportarExcel.Visible = False
    Me.txtCantidad.Visible = False
    
 Case "Reporte Total Vacaciones"
    Me.Frame3.Visible = True
    Me.Check1.Visible = False
    Me.CmdExportarExcel.Visible = False
    Me.txtCantidad.Visible = False
    
 Case "Reporte x Provision"
    Me.Frame3.Visible = True
    Me.Check1.Visible = False
    Me.CmdExportarExcel.Visible = False
    Me.txtCantidad.Visible = False
    
 Case "Reporte Consolidado Vacaciones"
    Me.Frame3.Visible = True
    Me.Check1.Visible = True
    Me.Check1.Caption = "Saldo a fecha mayor a:"
    Me.txtCantidad.Visible = True
    Me.CmdExportarExcel.Visible = False

 Case "Reporte IR MENSUAL"
    Me.Frame3.Visible = True
    Me.Check1.Visible = True
    Me.CmdExportarExcel.Visible = True

 Case "Reporte Detalle Ir"
    Me.Frame3.Visible = True
    
  Case "Reporte Inss 2"
    Me.Frame3.Visible = True
    Me.CmdExportarExcel.Visible = True
    
  Case "Reporte Inss"
   Me.Frame3.Visible = True
   Me.Frame1.Visible = False
   Me.CmdExportar.Visible = True
 

    
  Case "Reporte Detalle Inss"
    Me.Frame3.Visible = True
    
  Case "Salario Basico Vrs Produccion"
   Me.Frame3.Visible = True
   Me.Frame1.Visible = False
   Me.CmdExportarExcel.Visible = True
    
  Case "Resumen-Pago Mensual"
   Me.Frame3.Visible = True
   Me.Frame1.Visible = False
   
   Case "Total-Pago Mensual"
   Me.Frame3.Visible = True
   Me.Frame1.Visible = False
   
   Case "Detalle Deducciones"
   Me.Frame3.Visible = True
   Me.Frame1.Visible = False
   
   Case "Reporte Detalle Deducciones"
   Me.Frame3.Visible = True
   Me.Frame1.Visible = False
   
  Case "Lista de Empleados Activos"
    Me.CmdExportarExcel.Visible = True
    Me.Frame2.Visible = True
    Me.Frame1.Visible = False
    Me.Mes.Visible = False
    
  Case "Listado de Empleados"
    Me.Frame1.Visible = False
    Me.Mes.Visible = False
    
  Case "Reporte x Produccion Basico"
    Me.CmdExportar.Visible = True
    Me.Frame1.Visible = False
    Me.Mes.Visible = False
    Me.lblTitulo.Visible = False
    Me.Frame2.Visible = True
    Me.CmdExportarExcel.Visible = True
    Me.FrmSeleccion.Visible = False
  Case "Reporte x Produccion"
    Me.CmdExportar.Visible = True
    Me.Frame1.Visible = False
    Me.Mes.Visible = False
    Me.lblTitulo.Visible = False
    Me.Frame2.Visible = True
    Me.CmdExportarExcel.Visible = True
    Me.FrmSeleccion.Visible = True

  Case "Reporte x Produccion Linea"
    Me.CmdExportar.Visible = True
    Me.Frame1.Visible = False
    Me.Mes.Visible = False
    Me.lblTitulo.Visible = False
    Me.Frame2.Visible = True
    Me.CmdExportarExcel.Visible = True
    
  Case "Reporte INSS"
    Me.Frame1.Visible = False
    Me.Mes.Visible = True
    Me.lblTitulo.Visible = True
  Case "Reporte IR"
    Me.Frame1.Visible = False
    Me.Mes.Visible = True
    Me.lblTitulo.Visible = True
  Case Else
'   Me.Frame1.Visible = True
'    Me.Mes.Visible = False
    Me.lblTitulo.Visible = False
End Select
End Sub

Private Sub CmdDeducciones_Click()
If Not IsNumeric(DbcListNominas.Text) Then
   MsgBox "Error en la seleccion de la nómina"
   DbcListNominas.SetFocus
   Exit Sub
Else
    NumNomina = val(DbcListNominas.Text)
    ARPagoDeducciones.Show
End If
End Sub

Private Sub CmdIncentivos_Click()
    If Not IsNumeric(DbcListNominas.Text) Then
       MsgBox "Error en la seleccion de la nómina"
       DbcListNominas.SetFocus
       Exit Sub
    Else
        NumNomina = val(DbcListNominas.Text)
        ARPagoIncentivos.Show
    End If
End Sub

Private Sub CmdPagoPrestamo_Click()
    If Not IsNumeric(DbcListNominas.Text) Then
       MsgBox "Error en la seleccion de la nómina"
       DbcListNominas.SetFocus
       Exit Sub
    Else
        NumNomina = val(DbcListNominas.Text)
        ARPagoPrestamo.Show
    End If
End Sub

Private Sub CmdPagoSubsidio_Click()

If Not IsNumeric(DbcNomSubsidio.Text) Then
   MsgBox "Error en la seleccion de la nómina de Subsidios"
   DbcNomSubsidio.SetFocus
   Exit Sub
Else
    NumNominaSubsidio = val(DbcNomSubsidio.Text)
    ARPagoSubsidios.Show
End If



End Sub

Private Sub CmdBuscarEmpleadoFin_Click()
QueProducto = "CodigoEmpleadoReportesFin"
FrmConsulta.Show 1
Me.DataCombo2.Text = FrmConsulta.CodigoEmpleado1
End Sub

Private Sub CmdBuscarEmpleadoIni_Click()
QueProducto = "CodigoEmpleadoReportesIni"
FrmConsulta.Show 1
Me.DataCombo1.Text = FrmConsulta.CodigoEmpleado1
End Sub

Private Sub CmdExportar_Click()
On Error GoTo TipoErrs
Dim SQLExporta As String, Longitud As Integer, Respuesta As Integer
Dim Cadena As String, Mes As String, Dia As String, Ano As String
Dim TextoMonto As String, TipoMovimiento As String, j As Integer
Dim Codigo As String
salir = False
Me.Barra.Visible = True
Me.CommonDialog1.ShowSave
Directorio = ""
Directorio = Me.CommonDialog1.FileName + ".csv"
Fecha1 = Year(Me.TxtFecha1.Value) & "-" & Month(Me.TxtFecha1.Value) & "-" & Day(Me.TxtFecha1.Value)
Fecha2 = Year(Me.TxtFecha2.Value) & "-" & Month(Me.TxtFecha2.Value) & "-" & Day(Me.TxtFecha2.Value)

Select Case CmbReportes.Text
  Case "Listado Maestro de Empleados"
      sql = "SELECT     TOP 100 PERCENT Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico," & vbLf
    sql = sql & "                    Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
    sql = sql & "                     Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
    sql = sql & "                    Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
    sql = sql & "                    DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
    sql = sql & "                    Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
    sql = sql & "                    DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
    sql = sql & "                    DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
    sql = sql & "                    DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
    sql = sql & "                    DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
    sql = sql & "                    Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
    sql = sql & "                    DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
    sql = sql & "                     DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia AS TotalDevengado," & vbLf
    sql = sql & "                    DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
    sql = sql & "                    (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
    sql = sql & "                     DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia)" & vbLf
    sql = sql & "                    - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
    sql = sql & "                    Empleado.TarifaHoraria, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, Empleado.Activo, Empleado.NumeroInss, Empleado.CodEmpleado1," & vbLf
    sql = sql & "                    Historico.FechaContrato , Empleado.Sexo" & vbLf
    sql = sql & "FROM         Nomina INNER JOIN" & vbLf
    sql = sql & "                    Grupo INNER JOIN" & vbLf
    sql = sql & "                    Cargo INNER JOIN" & vbLf
    sql = sql & "                    TipoNomina INNER JOIN" & vbLf
    sql = sql & "                    Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
    sql = sql & "                    DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
    sql = sql & "                    TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN" & vbLf
    sql = sql & "                    Historico ON Empleado.CodEmpleado = Historico.Codempleado" & vbLf
    sql = sql & "Where (Nomina.NumNomina = '" & Me.TxtNNomina.Text & "')" & vbLf
    sql = sql & "ORDER BY Nomina.NumNomina, DetalleNomina.CodEmpleado"


  Case "Salario Basico Vrs Produccion"
     sql = "SELECT Empleado.CodEmpleado1, Empleado.CodEmpleado, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres,DetalleNomina.NumNomina, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HE, DetalleNomina.HorasExtras,DetalleNomina.OtrosIngresos , DetalleNomina.Comisiones, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, Nomina.FechaNomina FROM Empleado INNER JOIN" & vbLf
     sql = sql & "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN" & vbLf
     sql = sql & "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina" & vbLf
     sql = sql & "WHERE(Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102))AND (dbo.DetalleNomina.SalarioBasico <> 0)"
     sql = sql & "ORDER BY Empleado.CodEmpleado,DetalleNomina.NumNomina"
  Case "Reporte x Produccion"
     Numero = val(Me.TxtNNomina.Text)
     sql = "SELECT DetalleProduccion.CodEmpleado, DetalleProduccion.NumNomina, DetalleProduccion.CodReferencia, DetalleProduccion.CodProceso,DetalleProduccion.Ref, DetalleProduccion.Lunes, DetalleProduccion.Martes, DetalleProduccion.Miercoles, DetalleProduccion.Jueves,DetalleProduccion.Viernes, DetalleProduccion.Sabado, DetalleProduccion.Domingo, DetalleProduccion.TotalUnidades, DetalleProduccion.SalarioPieza,DetalleProduccion.Precio , DetalleProduccion.unidad, DetalleProduccion.Pagado, Empleado.CodEmpleado1 FROM DetalleProduccion INNER JOIN Empleado ON DetalleProduccion.CodEmpleado = Empleado.CodEmpleado Where (DetalleProduccion.NumNomina = " & Numero & ")"

   Case "Reporte Inss"
        sql = "SELECT     TOP 100 PERCENT dbo.Empleado.Nombre1 + N' ' + dbo.Empleado.Nombre2 + N' ' + dbo.Empleado.Apellido1 + N' ' + dbo.Empleado.Apellido2 AS Nombres," & vbLf
        sql = sql & "                       dbo.DetalleNomina.CodEmpleado, dbo.DetalleNomina.MontoINSS, dbo.DetalleNomina.INSSPatronal," & vbLf
        sql = sql & "                dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
        sql = sql & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.IncetivoProduccion AS TotalDevengado," & vbLf
        sql = sql & "                       dbo.DetalleNomina.INATEC, dbo.Empleado.CodInss, dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.INSSPatronal AS TotalInss," & vbLf
        sql = sql & "                      dbo.Empleado.CodEmpleado1 , dbo.DetalleNomina.NumNomina, dbo.Nomina.FechaNomina, dbo.Cargo.Cargo" & vbLf
        sql = sql & "FROM         dbo.Nomina INNER JOIN" & vbLf
        sql = sql & "                      dbo.Grupo INNER JOIN" & vbLf
        sql = sql & "                      dbo.Cargo INNER JOIN" & vbLf
        sql = sql & "                      dbo.TipoNomina INNER JOIN" & vbLf
        sql = sql & "                      dbo.Empleado ON dbo.TipoNomina.CodTipoNomina = dbo.Empleado.CodTipoNomina ON dbo.Cargo.CodCargo = dbo.Empleado.CodCargo INNER JOIN" & vbLf
        sql = sql & "                      dbo.DetalleNomina ON dbo.Empleado.CodEmpleado = dbo.DetalleNomina.CodEmpleado ON dbo.Grupo.CodGrupo = dbo.Empleado.CodGrupo ON" & vbLf
        sql = sql & "                      dbo.TipoNomina.CodTipoNomina = dbo.Nomina.CodTipoNomina And dbo.Nomina.NumNomina = dbo.DetalleNomina.NumNomina" & vbLf
        sql = sql & "WHERE(Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102))"
        sql = sql & "ORDER BY dbo.Empleado.CodEmpleado1, dbo.Nomina.FechaNomina"
   Case "Reporte Ir"
        sql = "SELECT     TOP 100 PERCENT dbo.Empleado.Nombre1 + N' ' + dbo.Empleado.Nombre2 + N' ' + dbo.Empleado.Apellido1 + N' ' + dbo.Empleado.Apellido2 AS Nombres," & vbLf
        sql = sql & "                       dbo.DetalleNomina.CodEmpleado, dbo.DetalleNomina.MontoINSS, dbo.DetalleNomina.INSSPatronal," & vbLf
        sql = sql & "                dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
        sql = sql & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.IncetivoProduccion AS TotalDevengado," & vbLf
        sql = sql & "                       dbo.DetalleNomina.INATEC, dbo.Empleado.CodInss, dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.INSSPatronal AS TotalInss," & vbLf
        sql = sql & "                      dbo.Empleado.CodEmpleado1 , dbo.DetalleNomina.NumNomina, dbo.Nomina.FechaNomina, dbo.Cargo.Cargo,DetalleNomina.MontoIr, Empleado.Codir" & vbLf
        sql = sql & "FROM         dbo.Nomina INNER JOIN" & vbLf
        sql = sql & "                      dbo.Grupo INNER JOIN" & vbLf
        sql = sql & "                      dbo.Cargo INNER JOIN" & vbLf
        sql = sql & "                      dbo.TipoNomina INNER JOIN" & vbLf
        sql = sql & "                      dbo.Empleado ON dbo.TipoNomina.CodTipoNomina = dbo.Empleado.CodTipoNomina ON dbo.Cargo.CodCargo = dbo.Empleado.CodCargo INNER JOIN" & vbLf
        sql = sql & "                      dbo.DetalleNomina ON dbo.Empleado.CodEmpleado = dbo.DetalleNomina.CodEmpleado ON dbo.Grupo.CodGrupo = dbo.Empleado.CodGrupo ON" & vbLf
        sql = sql & "                      dbo.TipoNomina.CodTipoNomina = dbo.Nomina.CodTipoNomina And dbo.Nomina.NumNomina = dbo.DetalleNomina.NumNomina" & vbLf
        sql = sql & "WHERE(Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102))AND(dbo.DetalleNomina.MontoIr <> 0)"
        sql = sql & "ORDER BY dbo.Empleado.CodEmpleado1, DetalleNomina.NumNomina"
     

End Select
Me.AdoBusca.RecordSource = sql
AdoBusca.Refresh
If AdoBusca.Recordset.EOF Then
  MsgBox "No Existen Registros", vbCritical, "Sistema de Nominas"
  Me.Barra.Visible = False
  Exit Sub
End If
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
                   
                Select Case CmbReportes.Text
                   Case "Salario Basico Vrs Produccion"
                      Cadena = "Codigo Empleado" & "," & "Nombres" & "," & "Numero Nomina" & "," & "Salario Basico" & "," & "Produccion" & "," & "Horas Extra" & "," & "Otros Ingresos" & "," & "Comisiones" & "," & "Horas Trabajada" & "," & "Septimo Dia" & "," & "Fecha Nomina"
                   Case "Reporte x Produccion"
                      Cadena = "Codigo Empleado" & "," & "Referencia" & "," & "Proceso" & "," & "Precio" & "," & "Referencia" & "," & "Lunes" & "," & "Martes" & "," & "Miercoles" & "," & "Jueves" & "," & "Viernes" & "," & "Sabado" & "," & "Domingo" & "," & "TotalUnidades" & "," & "SalarioPieza" & "," & "Unidad"
                   Case "Reporte Inss"
                      Cadena = "CodEmpleado" & "," & "Nombres" & "," & "Cargo" & "," & "TotalDevengado" & "," & "MontoINSS" & "," & "INSSPatronal" & "," & "TotalINSS" & "," & "INATEC"
                   Case "Reporte Ir"
                      Cadena = "CodEmpleado" & "," & "Nombres" & "," & "Cargo" & "," & "TotalDevengado" & "," & "MontoIr" & "," & "INATEC" & "," & "CodIr"

                 End Select
                 Print #1, Cadena
                 
                 
                 Do While Not AdoBusca.Recordset.EOF
                 '////////Inicialiso las variables/////////////////
                 Select Case CmbReportes.Text
                   Case "Salario Basico Vrs Produccion"
                      Cadena = AdoBusca.Recordset("CodEmpleado1") & "," & AdoBusca.Recordset("Nombres") & "," & AdoBusca.Recordset("NumNomina") & "," & AdoBusca.Recordset("SalarioBasico") & "," & AdoBusca.Recordset("Destajo") & "," & AdoBusca.Recordset("HorasExtras") & "," & AdoBusca.Recordset("OtrosIngresos") & "," & AdoBusca.Recordset("Comisiones") & "," & AdoBusca.Recordset("HTrabajada") & "," & AdoBusca.Recordset("SeptimoDia") & "," & AdoBusca.Recordset("FechaNomina")
                   Case "Reporte x Produccion"
                      Cadena = AdoBusca.Recordset("CodEmpleado1") & "," & AdoBusca.Recordset("CodReferencia") & "," & AdoBusca.Recordset("CodProceso") & "," & AdoBusca.Recordset("Precio") & "," & AdoBusca.Recordset("Ref") & "," & AdoBusca.Recordset("Lunes") & "," & AdoBusca.Recordset("Martes") & "," & AdoBusca.Recordset("Miercoles") & "," & AdoBusca.Recordset("Jueves") & "," & AdoBusca.Recordset("Viernes") & "," & AdoBusca.Recordset("Sabado") & "," & AdoBusca.Recordset("Domingo") & "," & AdoBusca.Recordset("TotalUnidades") & "," & AdoBusca.Recordset("SalarioPieza") & "," & AdoBusca.Recordset("Unidad")
                   Case "Reporte Inss"
                      Cadena = AdoBusca.Recordset("CodEmpleado1") & "," & AdoBusca.Recordset("Nombres") & "," & AdoBusca.Recordset("Cargo") & "," & AdoBusca.Recordset("TotalDevengado") & "," & AdoBusca.Recordset("MontoINSS") & "," & AdoBusca.Recordset("INSSPatronal") & "," & AdoBusca.Recordset("TotalINSS") & "," & AdoBusca.Recordset("INATEC")
                   Case "Reporte Ir"
                      Cadena = AdoBusca.Recordset("CodEmpleado1") & "," & AdoBusca.Recordset("Nombres") & "," & AdoBusca.Recordset("Cargo") & "," & AdoBusca.Recordset("TotalDevengado") & "," & AdoBusca.Recordset("MontoIr") & "," & AdoBusca.Recordset("INATEC") & "," & AdoBusca.Recordset("CodIr")

                 End Select
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
                   
                Select Case CmbReportes.Text
                   Case "Salario Basico Vrs Produccion"
                      Cadena = "Codigo Empleado" & "," & "Nombres" & "," & "Numero Nomina" & "," & "Salario Basico" & "," & "Produccion" & "," & "Horas Extra" & "," & "Otros Ingresos" & "," & "Comisiones" & "," & "Horas Trabajada" & "," & "Septimo Dia" & "," & "Fecha Nomina"
                   Case "Reporte x Produccion"
                      Cadena = "Codigo Empleado" & "," & "Referencia" & "," & "Proceso" & "," & "Precio" & "," & "Referencia" & "," & "Lunes" & "," & "Martes" & "," & "Miercoles" & "," & "Jueves" & "," & "Viernes" & "," & "Sabado" & "," & "Domingo" & "," & "TotalUnidades" & "," & "SalarioPieza" & "," & "Unidad"
                   Case "Reporte Inss"
                      Cadena = "CodEmpleado" & "," & "Nombres" & "," & "Cargo" & "," & "TotalDevengado" & "," & "MontoINSS" & "," & "INSSPatronal" & "," & "TotalINSS" & "," & "INATEC"
                   Case "Reporte Ir"
                      Cadena = "CodEmpleado" & "," & "Nombres" & "," & "Cargo" & "," & "TotalDevengado" & "," & "MontoIr" & "," & "INATEC" & "," & "CodIr"

                 End Select
                 Print #1, Cadena
                   
                 Do While Not AdoBusca.Recordset.EOF
                  Select Case CmbReportes.Text
                   Case "Salario Basico Vrs Produccion"
                      Cadena = AdoBusca.Recordset("CodEmpleado1") & "," & AdoBusca.Recordset("Nombres") & "," & AdoBusca.Recordset("NumNomina") & "," & AdoBusca.Recordset("SalarioBasico") & "," & AdoBusca.Recordset("Destajo") & "," & AdoBusca.Recordset("HorasExtras") & "," & AdoBusca.Recordset("OtrosIngresos") & "," & AdoBusca.Recordset("Comisiones") & "," & AdoBusca.Recordset("HTrabajada") & "," & AdoBusca.Recordset("SeptimoDia") & "," & AdoBusca.Recordset("FechaNomina")
                   Case "Reporte x Produccion"
                      Cadena = AdoBusca.Recordset("CodEmpleado1") & "," & AdoBusca.Recordset("CodReferencia") & "," & AdoBusca.Recordset("CodProceso") & "," & AdoBusca.Recordset("Precio") & "," & AdoBusca.Recordset("Ref") & "," & AdoBusca.Recordset("Lunes") & "," & AdoBusca.Recordset("Martes") & "," & AdoBusca.Recordset("Miercoles") & "," & AdoBusca.Recordset("Jueves") & "," & AdoBusca.Recordset("Viernes") & "," & AdoBusca.Recordset("Sabado") & "," & AdoBusca.Recordset("Domingo") & "," & AdoBusca.Recordset("TotalUnidades") & "," & AdoBusca.Recordset("SalarioPieza") & "," & AdoBusca.Recordset("Unidad")
                    Case "Reporte Inss"
                      Cadena = AdoBusca.Recordset("CodEmpleado1") & "," & AdoBusca.Recordset("Nombres") & "," & AdoBusca.Recordset("Cargo") & "," & AdoBusca.Recordset("TotalDevengado") & "," & AdoBusca.Recordset("MontoINSS") & "," & AdoBusca.Recordset("INSSPatronal") & "," & AdoBusca.Recordset("TotalINSS") & "," & AdoBusca.Recordset("INATEC")
                   Case "Reporte Ir"
                      Cadena = AdoBusca.Recordset("CodEmpleado1") & "," & AdoBusca.Recordset("Nombres") & "," & AdoBusca.Recordset("Cargo") & "," & AdoBusca.Recordset("TotalDevengado") & "," & AdoBusca.Recordset("MontoIr") & "," & AdoBusca.Recordset("INATEC") & "," & AdoBusca.Recordset("CodIr")
                   End Select

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

                MsgBox "La Exportacion, fue Creada con Exito", vbExclamation, "Sistema de Enlace"
                Me.Barra.Visible = False
  End If
Exit Sub
TipoErrs:
  MsgBox Err.Description
End Sub

Private Sub CmdExportarExcel_Click()
Dim Semanas As String, Novedad As String, TarifaHoraria As Double
Dim departamento As String, Periodo As Integer, Semana() As Variant, PeriodoNomina As String
Dim cnDB As New ADODB.Connection, MontoInssVaca As Double, SalarioVaca As Double
Dim rsBD As New ADODB.Recordset, SueldoPeriodo As Double, FechaBaja As Date
Dim iCont As Integer, TipoNomia As String, agregar As Boolean, TotalAjuste As Double
Dim FechaContrato As Date, SemanaAjuste() As Double, SqlString As String, MontoInssBaja As Double, MontoIrBaja As Double
Dim DiasMes As Double, CodEmpleado As Double
Dim CodTipoNomina As String, AjusteINSS As Double, MontoInssBasico As Double, MontoInss As Double
Dim TarifaHorariaBasico As Double, TasaInss As Double, NumNomina As Double

If Me.CmbReportes.Text <> "EXPORTACION INSS" Then
 Me.CommonDialog1.ShowSave
 Directorio = ""
 Directorio = Me.CommonDialog1.FileName + ".xls"
End If

Select Case CmbReportes.Text

 Case "Reporte IR MENSUAL"
 
 Fecha1 = Year(Me.TxtFecha1.Value) & "-" & Month(Me.TxtFecha1.Value) & "-" & Day(Me.TxtFecha1.Value)
 Fecha2 = Year(Me.TxtFecha2.Value) & "-" & Month(Me.TxtFecha2.Value) & "-" & Day(Me.TxtFecha2.Value)
   
    
  If Me.Check1.Value = 0 Then
   
'     sql = "SELECT     TOP 100 PERCENT MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, " & _
'                      "SUM(DetalleNomina.SalarioBasico  + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos " & _
'                       "+ DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion+ DetalleNomina.BonoProduccion + dbo.DetalleNomina.Antiguedad) AS TotalDevengado, " & _
'                       "Empleado.CodEmpleado1, MAX(DetalleNomina.NumNomina) AS NumNomina, SUM(DetalleNomina.MontoIR) AS MontoIR, MAX(Empleado.NumCedula) AS NumCedula, SUM(DetalleNomina.MontoINSS) AS MontoINSS, " & _
'                      "Nomina.CodTipoNomina AS CodTipoNomina " & _
'            "FROM         Nomina INNER JOIN " & _
'                      "Grupo INNER JOIN " & _
'                      "Cargo INNER JOIN " & _
'                      "TipoNomina INNER JOIN " & _
'                      "Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN " & _
'                      "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON " & _
'                      "TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina " & _
'            "WHERE     (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
'            "GROUP BY Empleado.CodEmpleado1, Nomina.CodTipoNomina " & _
'            "HAVING      (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') AND (SUM(DetalleNomina.MontoIR) <> 0) " & _
'            "ORDER BY Empleado.CodEmpleado1, MAX(DetalleNomina.NumNomina) "
  
       sql = "SELECT    TOP (100) PERCENT MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad) AS TotalDevengado, Empleado.CodEmpleado, Empleado.CodEmpleado1, MAX(DetalleNomina.NumNomina) AS NumNomina, Nomina.CodTipoNomina, MAX(Empleado.NumCedula) AS NumCedula, CASE WHEN SUM(Bajas.MontoIR) IS NULL THEN SUM(DetalleNomina.MontoIR) ELSE SUM(DetalleNomina.MontoIR + Bajas.MontoIR) END AS MontoIR, CASE WHEN SUM(Bajas.MontoINSS) IS NULL THEN SUM(DetalleNomina.MontoINSS) ELSE SUM(DetalleNomina.MontoINSS + Bajas.MontoINSS) END As MontoInss FROM  Nomina INNER JOIN  Grupo INNER JOIN  Cargo INNER JOIN TipoNomina INNER JOIN " & _
            "Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN  DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON  TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina LEFT OUTER JOIN Bajas ON Empleado.CodEmpleado = Bajas.CodEmpleado  " & _
            "WHERE  (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Empleado.CodEmpleado, Empleado.CodEmpleado1, Nomina.CodTipoNomina HAVING (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') AND (SUM(DetalleNomina.MontoIR) <> 0) ORDER BY Empleado.CodEmpleado1, NumNomina"
  
  Else
  
   '//////////////////////////////////////////////////////////BUSCO SI EXISTEN BAJAS PARA ESTE PERIODO ///////////////////////////////////////////////////////////////////////////////////////////////
   
  
   
     sql = "SELECT    TOP (100) PERCENT MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad) AS TotalDevengado, Empleado.CodEmpleado, Empleado.CodEmpleado1, MAX(DetalleNomina.NumNomina) AS NumNomina, Nomina.CodTipoNomina, MAX(Empleado.NumCedula) AS NumCedula, CASE WHEN SUM(Bajas.MontoIR) IS NULL THEN SUM(DetalleNomina.MontoIR) ELSE SUM(DetalleNomina.MontoIR) END AS MontoIR, CASE WHEN SUM(Bajas.MontoINSS) IS NULL THEN SUM(DetalleNomina.MontoINSS) ELSE SUM(DetalleNomina.MontoINSS) END As MontoInss FROM  Nomina INNER JOIN  Grupo INNER JOIN  Cargo INNER JOIN TipoNomina INNER JOIN " & _
            "Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN  DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON  TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina LEFT OUTER JOIN Bajas ON Empleado.CodEmpleado = Bajas.CodEmpleado  " & _
            "WHERE  (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Empleado.CodEmpleado,Empleado.CodEmpleado1, Nomina.CodTipoNomina HAVING (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') ORDER BY Empleado.CodEmpleado1, NumNomina"
  End If
  
      Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
    
        V = 4
        H = 0
        i = 1

 
 '///////////////////////////////////////////////////////////////////////////////////////
'////////////////////ENCABEZADOS//////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////
            objExcel.ActiveSheet.Cells(1, 1) = Titulo
            objExcel.ActiveSheet.Range("A1:J1").Merge
            objExcel.ActiveSheet.Range("A1", "J1").HorizontalAlignment = xlHAlignCenter
            With objExcel.ActiveSheet.Cells(1, 1)
                  .Font.Size = 20             ' tamaño de letra
                  .Font.Bold = True           ' Fuente en negrita
            End With
            
            objExcel.ActiveSheet.Cells(2, 1) = "EXPORTACION DESDE " & Me.TxtFecha1.Value & " HASTA " & Me.TxtFecha2.Value
            objExcel.ActiveSheet.Range("A2:J2").Merge
            objExcel.ActiveSheet.Range("A2", "J2").HorizontalAlignment = xlHAlignCenter
            With objExcel.ActiveSheet.Cells(2, 1)
                  .Font.Size = 14             ' tamaño de letra
                  .Font.Bold = True           ' Fuente en negrita
            End With
            
            objExcel.ActiveSheet.Cells(3, 1) = "CodEmpleado"
            objExcel.ActiveSheet.Cells(3, 2) = "Nombre y Apellido"
            objExcel.ActiveSheet.Cells(3, 3) = "NumerCedula"
            objExcel.ActiveSheet.Cells(3, 4) = "Devengado"
            objExcel.ActiveSheet.Cells(3, 5) = "MontoInss"
            objExcel.ActiveSheet.Cells(3, 6) = "MontoIr"
            
       Me.AdoConsulta.RecordSource = sql
       Me.AdoConsulta.Refresh
       Do While Not Me.AdoConsulta.Recordset.EOF
       
       '///////////////////////////////////////////////////////////////////////////////////////////////////////
       '///////////////////CONSULTO SI TIENE BAJAS ////////////////////////////////////////////////////////////
       '//////////////////////////////////////////////////////////////////////////////////////////////////////
       CodEmpleado = Me.AdoConsulta.Recordset("CodEmpleado")
       MontoInssBaja = 0
       MontoIrBaja = 0
       
       SqlString = "SELECT MontoINSS, MontoIR From Bajas WHERE (CodEmpleado = " & CodEmpleado & ") AND (FechaBaja BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102))"
       MDIPrimero.AdoConsulta.RecordSource = SqlString
       MDIPrimero.AdoConsulta.ConnectionString = Conexion
       MDIPrimero.AdoConsulta.Refresh
       If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
         MontoInssBaja = MDIPrimero.AdoConsulta.Recordset("MontoINSS")
         MontoIrBaja = MDIPrimero.AdoConsulta.Recordset("MontoIR")
       End If
       
       
                objExcel.ActiveSheet.Cells(V, H + 1) = Me.AdoConsulta.Recordset("CodEmpleado1")
                objExcel.ActiveSheet.Cells(V, H + 2) = Me.AdoConsulta.Recordset("Nombres")
                objExcel.ActiveSheet.Cells(V, H + 3) = Me.AdoConsulta.Recordset("NumCedula")
                objExcel.ActiveSheet.Cells(V, H + 4) = Format(Me.AdoConsulta.Recordset("TotalDevengado"), "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 5) = Format(Me.AdoConsulta.Recordset("MontoINSS") + MontoInssBaja, "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 6) = Format(Me.AdoConsulta.Recordset("MontoIR") + MontoIrBaja, "##,##0.00")
                    
      
         V = V + 1
         Me.AdoConsulta.Recordset.MoveNext
       Loop
       
       
        objExcel.ActiveSheet.Columns("A").ColumnWidth = 13.3
        objExcel.ActiveSheet.Columns("B").ColumnWidth = 50
        objExcel.ActiveSheet.Columns("C").ColumnWidth = 17

 


 Case "Exportar Lista Empleados"
 
'           Me.Barra.Visible = True
'    Me.AdoBusca.RecordSource = SQl
'    AdoBusca.Refresh
'    If AdoBusca.Recordset.EOF Then
'         MsgBox "No Existen Registros", vbCritical, "Sistema de Nominas"
'         Me.Barra.Visible = False
'         Exit Sub
'    End If
    
'    Me.AdoBusca.Recordset.MoveLast
'    Maximo = AdoBusca.Recordset.RecordCount
    Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
    
        V = 4
        H = 0
        i = 1

'///////////////////////////////////////////////////////////////////////////////////////
'////////////////////ENCABEZADOS//////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////
            objExcel.ActiveSheet.Cells(1, 1) = Titulo
            objExcel.ActiveSheet.Range("A1:J1").Merge
            objExcel.ActiveSheet.Range("A1", "J1").HorizontalAlignment = xlHAlignCenter
            With objExcel.ActiveSheet.Cells(1, 1)
                  .Font.Size = 20             ' tamaño de letra
                  .Font.Bold = True           ' Fuente en negrita
            End With
            
            objExcel.ActiveSheet.Cells(2, 1) = "EXPORTACION DESDE " & Me.TxtFecha1.Value & " HASTA " & Me.TxtFecha2.Value
            objExcel.ActiveSheet.Range("A2:J2").Merge
            objExcel.ActiveSheet.Range("A2", "J2").HorizontalAlignment = xlHAlignCenter
            With objExcel.ActiveSheet.Cells(2, 1)
                  .Font.Size = 14             ' tamaño de letra
                  .Font.Bold = True           ' Fuente en negrita
            End With
            
            objExcel.ActiveSheet.Cells(3, 1) = "CodEmpleado"
            objExcel.ActiveSheet.Cells(3, 2) = "Nombre y Apellido"
            objExcel.ActiveSheet.Cells(3, 3) = "NumeroInss"
            objExcel.ActiveSheet.Cells(3, 4) = "NumCedula"
            objExcel.ActiveSheet.Cells(3, 5) = "FechaContrato"
            objExcel.ActiveSheet.Cells(3, 6) = "Cargo"
            objExcel.ActiveSheet.Cells(3, 7) = "SalarioBasico"
            objExcel.ActiveSheet.Cells(3, 8) = "Destajo"
            objExcel.ActiveSheet.Cells(3, 9) = "Antiguedad"
            objExcel.ActiveSheet.Cells(3, 10) = "SalarioMensual"
            
 
       If Me.DBTipoNominas.Text = "" Then
             sql = "SELECT  Empleado.CodEmpleado1,  Empleado.Apellido1 + ' ' +  Empleado.Apellido2 + ' ' +  Empleado.Nombre1 + ' ' +  Empleado.Nombre2  AS Nombres, Empleado.NumeroInss, Empleado.NumCedula, MAX(Historico.FechaContrato) AS FechaContrato, MAX(Cargo.Cargo) AS Cargo, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo  +  DetalleNomina.Antiguedad) AS SalarioMensual, SUM(DetalleNomina.Antiguedad) AS Antiguedad, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SalarioBasico) As SalarioBasico FROM Empleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN  DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN " & _
                  "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina WHERE (Nomina.FechaNominaINI >= CONVERT(DATETIME, '" & Format(Me.TxtFecha1.Value, "yyyy-mm-dd") & "', 102)) AND (Nomina.FechaNomina <= CONVERT(DATETIME, '" & Format(Me.TxtFecha2.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Empleado.CodEmpleado1,  Empleado.Apellido1 + ' ' +  Empleado.Apellido2 + ' ' +  Empleado.Nombre1 + ' ' +  Empleado.Nombre2, Empleado.NumeroInss, Empleado.NumCedula, Empleado.Apellido1 ORDER BY Empleado.Apellido1"  'REPLACE(STR(Empleado.CodEmpleado1), ' ', '0')
       Else
            sql = "SELECT  Empleado.CodEmpleado1,  Empleado.Apellido1 + ' ' +  Empleado.Apellido2 + ' ' +  Empleado.Nombre1 + ' ' +  Empleado.Nombre2 AS Nombres, Empleado.NumeroInss, Empleado.NumCedula, MAX(Historico.FechaContrato) AS FechaContrato, MAX(Cargo.Cargo) AS Cargo, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo  + DetalleNomina.Antiguedad) AS SalarioMensual, SUM(DetalleNomina.Antiguedad) AS Antiguedad, SUM(DetalleNomina.Destajo) AS Destajo, SUM(DetalleNomina.SalarioBasico) As SalarioBasico FROM Empleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN  DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN " & _
                  "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina WHERE (Nomina.FechaNominaINI >= CONVERT(DATETIME, '" & Format(Me.TxtFecha1.Value, "yyyy-mm-dd") & "', 102)) AND (Nomina.FechaNomina <= CONVERT(DATETIME, '" & Format(Me.TxtFecha2.Value, "yyyy-mm-dd") & "', 102)) AND (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') GROUP BY Empleado.CodEmpleado1,  Empleado.Apellido1 + ' ' +  Empleado.Apellido2 + ' ' +  Empleado.Nombre1 + ' ' +  Empleado.Nombre2, Empleado.NumeroInss, Empleado.NumCedula, Empleado.Apellido1 ORDER BY Empleado.Apellido1"  'REPLACE(STR(Empleado.CodEmpleado1), ' ', '0')
       End If
       
       Me.AdoConsulta.RecordSource = sql
       Me.AdoConsulta.Refresh
       Do While Not Me.AdoConsulta.Recordset.EOF
         
                objExcel.ActiveSheet.Cells(V, H + 1) = Me.AdoConsulta.Recordset("CodEmpleado1")
                objExcel.ActiveSheet.Cells(V, H + 2) = Me.AdoConsulta.Recordset("Nombres")
                objExcel.ActiveSheet.Cells(V, H + 3) = Me.AdoConsulta.Recordset("NumeroInss")
                objExcel.ActiveSheet.Cells(V, H + 4) = Me.AdoConsulta.Recordset("NumCedula")
                objExcel.ActiveSheet.Cells(V, H + 5) = Format(Me.AdoConsulta.Recordset("FechaContrato"), "dd/mm/yyyy")
                objExcel.ActiveSheet.Cells(V, H + 6) = Me.AdoConsulta.Recordset("Cargo")
                objExcel.ActiveSheet.Cells(V, H + 7) = Format(Me.AdoConsulta.Recordset("SalarioBasico"), "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 8) = Format(Me.AdoConsulta.Recordset("Destajo"), "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 9) = Format(Me.AdoConsulta.Recordset("Antiguedad"), "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 10) = Format(Me.AdoConsulta.Recordset("SalarioMensual"), "##,##0.00")
     
      
         V = V + 1
         Me.AdoConsulta.Recordset.MoveNext
       Loop
     
  


        objExcel.ActiveSheet.Columns("A").ColumnWidth = 13.57
        objExcel.ActiveSheet.Columns("B").ColumnWidth = 37.71
        objExcel.ActiveSheet.Columns("D").ColumnWidth = 19.5
        objExcel.ActiveSheet.Columns("G").ColumnWidth = 13.57
 
 Set objExcel = Nothing
 

 Case "Reporte INSS E IR MENSUAL"
    Fecha1 = Year(Me.TxtFecha1.Value) & "-" & Month(Me.TxtFecha1.Value) & "-" & Day(Me.TxtFecha1.Value)
    Fecha2 = Year(Me.TxtFecha2.Value) & "-" & Month(Me.TxtFecha2.Value) & "-" & Day(Me.TxtFecha2.Value)

    Exportar = True
    ArepInssIr.DataControl1.ConnectionString = ConexionReporte
    ArepInssIr.lblTitulo.Caption = Titulo
    ArepInssIr.LblSubtitulo.Caption = "REPORTE DETALLADO DEDUCCIONES SEGUN NOMINA"
    ArepInssIr.LblFecha.Caption = "Impreso desde: " & Me.TxtFecha1.Value & " Hasta: " & Me.TxtFecha2.Value
    ArepInssIr.LblFechaHoy.Caption = Format(Now, "Long Date")
    
          
    
'   sql = "SELECT Empleado.CodEmpleado1 AS CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) AS TotalIngresos,MAX(DetalleNomina.NumNomina) AS NumNomina, Fecha_Planilla.mes AS Mes, Fecha_Planilla.año AS Año, SUM(DetalleNomina.MontoINSS)AS MontoInss, SUM(DetalleNomina.MontoIR) AS MontoIR, Empleado.NumCedula FROM  Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN Fecha_Planilla ON DetalleNomina.NumNomina = Fecha_Planilla.NumNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina  " & _
'         "WHERE (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) AND (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') GROUP BY Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2, Fecha_Planilla.año, Empleado.CodEmpleado1, Fecha_Planilla.Mes , Empleado.NumCedula HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos                        + DetalleNomina.Incentivos + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) <> 0) ORDER BY MAX(DetalleNomina.NumNomina), Empleado.CodEmpleado1"

    sql = "SELECT Empleado.CodEmpleado1 AS CodEmpleado1, Empleado.Nombre1 + N' ' + Empleado.Nombre2 AS Nombres, Empleado.Apellido1, Empleado.Apellido2, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) AS TotalIngresos, MAX(DetalleNomina.NumNomina) AS NumNomina, Fecha_Planilla.mes AS Mes, Fecha_Planilla.año AS Año, SUM(DetalleNomina.MontoINSS) AS MontoInss, SUM(DetalleNomina.MontoIR) AS MontoIR, Empleado.NumCedula, Historico.FechaContrato FROM Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN  Fecha_Planilla ON DetalleNomina.NumNomina = Fecha_Planilla.NumNomina INNER JOIN  Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado " & _
          "WHERE (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) AND (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') GROUP BY Empleado.Nombre1 + N' ' + Empleado.Nombre2, Fecha_Planilla.año, Empleado.CodEmpleado1, Fecha_Planilla.mes, Empleado.NumCedula, Empleado.Apellido1 , Empleado.Apellido2, Historico.FechaContrato HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) <> 0) ORDER BY MAX(DetalleNomina.NumNomina), Empleado.CodEmpleado1"
    ArepInssIr.ImgLogo.Picture = LoadPicture(RutaLogo)
    ArepInssIr.DataControl1.Source = sql
    ArepInssIr.Show 1
  

 Case "Lista de Empleados Activos"
      Exportar = True
     ArepActivos.DataControl1.ConnectionString = ConexionReporte
     ArepActivos.lblTitulo.Caption = Titulo
     ArepActivos.LblSubtitulo.Caption = SubTitulo
     ArepActivos.ImgLogo.Picture = LoadPicture(RutaLogo)
'     ArepActivos.DataControl1.Source = "SELECT Empleado.CodEmpleado1,Empleado.CodEmpleado, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres,departamento.departamento , TipoNomina.Nomina, Empleado.Activo FROM Empleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina ORDER BY Empleado.CodEmpleado1"
     ArepActivos.DataControl1.Source = "SELECT     Empleado.CodEmpleado1, Empleado.CodEmpleado,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Departamento.Departamento, TipoNomina.Nomina , Empleado.Activo, Empleado.CodTipoNomina FROM  Empleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina WHERE     (Empleado.Activo = 1) AND (Empleado.CodTipoNomina =  '" & Me.TDBCombo1.Columns(0).Text & "') ORDER BY Empleado.CodEmpleado1"
     ArepActivos.Show 1

 Case "Reporte x Produccion Basico"
     Exportar = True
     Numero = val(Me.TxtNNomina.Text)
     
     ArepProduccionBasico.DataControl1.ConnectionString = ConexionReporte
     ArepProduccionBasico.lblTitulo.Caption = Titulo
     ArepProduccionBasico.LblSubtitulo.Caption = SubTitulo
     ArepProduccionBasico.ImgLogo.Picture = LoadPicture(RutaLogo)

     ArepProduccionBasico.DataControl1.Source = "SELECT Empleado.CodEmpleado1, DetalleProduccion.CodEmpleado, DetalleProduccion.NumNomina, SUM(DetalleProduccion.SalarioPieza) AS SalarioPieza, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Departamento.Departamento, Cargo.Cargo FROM DetalleProduccion INNER JOIN Empleado ON DetalleProduccion.CodEmpleado = Empleado.CodEmpleado INNER JOIN Nomina ON DetalleProduccion.NumNomina = Nomina.NumNomina INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN Cargo ON Empleado.CodCargo = Cargo.CodCargo GROUP BY Empleado.CodEmpleado1, DetalleProduccion.CodEmpleado, DetalleProduccion.NumNomina, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2, Departamento.Departamento, Cargo.Cargo Having (DetalleProduccion.NumNomina = " & Numero & ") And (Sum(DetalleProduccion.SalarioPieza) <> 0) ORDER BY Empleado.CodEmpleado1"


     ArepProduccionBasico.Show 1
 Case "Reporte x Produccion"
     Exportar = True
     Numero = val(Me.TxtNNomina.Text)
     ArepProduccion.LblTitulo3.Caption = "Reporte de Produccion de la Nomina No " & Numero
     ArepProduccion.DataControl1.ConnectionString = ConexionReporte
     ArepProduccion.lblTitulo.Caption = Titulo
     ArepProduccion.LblSubtitulo.Caption = SubTitulo
     ArepProduccion.ImgLogo.Picture = LoadPicture(RutaLogo)


    If Me.OptProduccionCompleta.Value = True Then
     ArepProduccion.DataControl1.Source = "SELECT Empleado.CodEmpleado1 AS CodEmpleado1, *, Nomina.FechaNominaINI AS FechaNominaINI, Nomina.FechaNomina AS FechaNomina FROM DetalleProduccion INNER JOIN Empleado ON DetalleProduccion.CodEmpleado = Empleado.CodEmpleado INNER JOIN Nomina ON DetalleProduccion.NumNomina = Nomina.NumNomina  " & _
                                          "Where (DetalleProduccion.NumNomina = " & Numero & ") ORDER BY Empleado.CodEmpleado1"
    ElseIf Me.OptProduccionNoPagada.Value = True Then
     ArepProduccion.DataControl1.Source = "SELECT Empleado.CodEmpleado1 AS CodEmpleado1, *, Nomina.FechaNominaINI AS FechaNominaINI, Nomina.FechaNomina AS FechaNomina FROM DetalleProduccion INNER JOIN Empleado ON DetalleProduccion.CodEmpleado = Empleado.CodEmpleado INNER JOIN Nomina ON DetalleProduccion.NumNomina = Nomina.NumNomina  " & _
                                          "WHERE (DetalleProduccion.NumNomina = " & Numero & ") AND (Empleado.CodEmpleado IN (SELECT CodEmpleado From DetalleNomina WHERE (NumNomina = " & Numero & ") AND (Destajo = 0) AND (BonoProduccion = 0))) ORDER BY Empleado.CodEmpleado1"
    ElseIf Me.OptProduccionPagada.Value = True Then
     ArepProduccion.DataControl1.Source = "SELECT Empleado.CodEmpleado1 AS CodEmpleado1, *, Nomina.FechaNominaINI AS FechaNominaINI, Nomina.FechaNomina AS FechaNomina FROM DetalleProduccion INNER JOIN Empleado ON DetalleProduccion.CodEmpleado = Empleado.CodEmpleado INNER JOIN Nomina ON DetalleProduccion.NumNomina = Nomina.NumNomina  " & _
                                          "WHERE (DetalleProduccion.NumNomina = " & Numero & ") AND (Empleado.CodEmpleado NOT IN (SELECT CodEmpleado From DetalleNomina WHERE (NumNomina = " & Numero & ") AND (Destajo = 0) AND (BonoProduccion = 0))) ORDER BY Empleado.CodEmpleado1"
    End If

     ArepProduccion.Show 1


 Case "Reporte x Produccion Linea"
     Exportar = True
     Numero = val(Me.TxtNNomina.Text)
     ArepProduccionLinea.LblTitulo3.Caption = "Reporte de Produccion de la Nomina No " & Numero
     ArepProduccionLinea.DataControl1.ConnectionString = ConexionReporte
     ArepProduccionLinea.lblTitulo.Caption = Titulo
     ArepProduccionLinea.LblSubtitulo.Caption = SubTitulo
     ArepProduccionLinea.ImgLogo.Picture = LoadPicture(RutaLogo)

    ArepProduccionLinea.DataControl1.Source = "SELECT DetalleNomina.NumNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HE, " & _
                      "DetalleNomina.DD, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, " & _
                      "DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, " & _
                      "DetalleNomina.Vacaciones, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.INATEC, DetalleNomina.Mes13, " & _
                      "DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.TotalSubsidio, DetalleNomina.VacacionesPagadas, " & _
                      "DetalleNomina.DiasVacaciones, DetalleNomina.AdelantosVacaciones, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, " & _
                      "DetalleNomina.IncetivoProduccion, DetalleNomina.TarifaHoraria, Nomina.FechaNomina, Nomina.FechaNominaINI, Empleado.CodEmpleado1, " & _
                      "DetalleProduccion.Linea,DetalleProduccion.CodReferencia, DetalleProduccion.CodReferencia1, DetalleProduccion.CodProceso, DetalleProduccion.Ref, DetalleProduccion.Lunes, DetalleProduccion.Martes, " & _
                      "DetalleProduccion.Miercoles, DetalleProduccion.Jueves, DetalleProduccion.Viernes, DetalleProduccion.Sabado, DetalleProduccion.Domingo, " & _
                      "DetalleProduccion.TotalUnidades, DetalleProduccion.SalarioPieza, DetalleProduccion.Precio, DetalleProduccion.Unidad, " & _
                      "DetalleProduccion.Pagado FROM DetalleNomina INNER JOIN " & _
                      "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN " & _
                      "Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado INNER JOIN " & _
                      "DetalleProduccion ON Empleado.CodEmpleado = DetalleProduccion.CodEmpleado " & _
                      "WHERE (DetalleNomina.NumNomina = " & Numero & ")AND (DetalleProduccion.NumNomina = " & Numero & ") ORDER BY DetalleProduccion.Linea,Empleado.CodEmpleado1, DetalleProduccion.CodReferencia, DetalleProduccion.CodProceso"
                       ' ORDER BY Empleado.CodEmpleado1"
     
'     ArepProduccion.DataControl1.Source = "SELECT DetalleProduccion.CodEmpleado, DetalleProduccion.NumNomina, DetalleProduccion.CodReferencia, DetalleProduccion.CodProceso,DetalleProduccion.Ref, DetalleProduccion.Lunes, DetalleProduccion.Martes, DetalleProduccion.Miercoles, DetalleProduccion.Jueves,DetalleProduccion.Viernes, DetalleProduccion.Sabado, DetalleProduccion.Domingo, DetalleProduccion.TotalUnidades, DetalleProduccion.SalarioPieza,DetalleProduccion.Precio , DetalleProduccion.unidad, DetalleProduccion.Pagado, Empleado.CodEmpleado1 FROM DetalleProduccion INNER JOIN Empleado ON DetalleProduccion.CodEmpleado = Empleado.CodEmpleado Where (DetalleProduccion.NumNomina = " & Numero & ")"
     ArepProduccionLinea.Show 1

 
 Case "EXPORTACION INSS"
    Dim MesAnterior As String, TotalDevengado As Double, DiaFinal As Date
    Dim NombreEmpresa As String, SemanaPeriodo As Double, MontoBasico As Double
    Fecha1 = Year(Me.TxtFecha1.Value) & "-" & Month(Me.TxtFecha1.Value) & "-" & Day(Me.TxtFecha1.Value)
    Fecha2 = Year(Me.TxtFecha2.Value) & "-" & Month(Me.TxtFecha2.Value) & "-" & Day(Me.TxtFecha2.Value)


     Me.AdoDatosEmpresa.RecordSource = "SELECT  Numero, NombreEmpresa, NumeroRUC, Direccion, Telefono, Fax, Email, RutaLogo From DatosEmpresa "
     Me.AdoDatosEmpresa.Refresh
     If Not Me.AdoDatosEmpresa.Recordset.EOF Then
       NombreEmpresa = Me.AdoDatosEmpresa.Recordset("NombreEmpresa")
     End If
     
     MDIPrimero.DtaControles.Refresh
     DiasMes = MDIPrimero.DtaControles.Recordset("DiasMes")
     CodTipoNomina = Me.DBTipoNominas.Columns(0).Text
     
     MDIPrimero.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, PorcientoInss, TasaInssPatronal, TasaInss, PorcientoIr, TasaIr From TipoNomina WHERE (PorcientoInss = 1) AND (CodTipoNomina = '" & CodTipoNomina & "' )"
     MDIPrimero.DtaConsulta.Refresh
     If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
       TasaInss = MDIPrimero.DtaConsulta.Recordset("TasaInss")
     End If
     

'    MDIPrimero.DtaConsulta.RecordSource = "SELECT * From TipoNomina WHERE (CodTipoNomina = '" & CodTipoNomina & "')"
'    MDIPrimero.DtaConsulta.Refresh
'    If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
'     If Not IsNull(MDIPrimero.DtaConsulta.Recordset("TarifaHoraria")) Then
'      TarifaHorariaBasico = MDIPrimero.DtaConsulta.Recordset("TarifaHoraria")
'     End If
'    Else
'      TarifaHorariaBasico = 0
'    End If
     
        
     
     '////////////////BUSCO LA FECHA DEL MES ANTERIOR///////////

    Select Case Me.Combo1.Text
    Case "Enero"
        MesAnterior = "Diciembre"
    Case "Febrero"
        MesAnterior = "Enero"
    Case "Marzo"
        MesAnterior = "Febrero"
    Case "Abril"
        MesAnterior = "Marzo"
    Case "Mayo"
        MesAnterior = "Abril"
    Case "Junio"
        MesAnterior = "Mayo"
    Case "Julio"
        MesAnterior = "Junio"
    Case "Agosto"
        MesAnterior = "Julio"
    Case "Septiembre"
         MesAnterior = "Agosto"
    Case "Octubre"
         MesAnterior = "Septiembre"
    Case "Noviembre"
          MesAnterior = "Octubre"
    Case "Diciembre"
         MesAnterior = "Noviembre"

    End Select

    FMes (MesAnterior)
    Mes1 = Format(Nmes, "0#")
    FMes (MesAnterior)
    Mes2 = Format(Nmes, "0#")
    If MesAnterior = "Diciembre" Then
        Año1 = val(Me.DBCAño.Text) - 1
        Año2 = val(Me.DBCAño.Text) - 1
    Else
        Año1 = val(Me.DBCAño.Text)
        Año2 = val(Me.DBCAño.Text)
    End If
    CodTipoNomina = Me.DBTipoNominas.Columns(0).Text


    Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
    Me.AdoBusca.Refresh
    If Not Me.AdoBusca.Recordset.EOF Then
        Fecha1Reporte = Format(Me.AdoBusca.Recordset("Inicio"), "yyyy/mm/dd")
    End If
 
    Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
    Me.AdoBusca.Refresh
    If Not Me.AdoBusca.Recordset.EOF Then
        Me.AdoBusca.Recordset.MoveLast
        Fecha2Reporte = Format(Me.AdoBusca.Recordset("Final"), "yyyy/mm/dd")
    End If

    Mes1 = Month(Fecha2)
    ConvertirMes (Mes1)
    
  
     
   
                   sql = "SELECT     TOP (100) PERCENT Empleado.Nombre1, Empleado.Apellido1, DetalleNomina.CodEmpleado, DetalleNomina.MontoINSS AS MontoInss, DetalleNomina.INSSPatronal AS InssPatronal,"
                   sql = sql + "   DetalleNomina.SalarioBasico  + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas"
                   sql = sql + "     + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad - ISNULL"
                   sql = sql + "        ((SELECT     SUM(DetalleIncentivo.Valor) AS Valor"
                   sql = sql + "            FROM         Incentivo INNER JOIN"
                   sql = sql + "                                  DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo"
                   sql = sql + "            WHERE     (Incentivo.CodTipoIncentivo = '14') OR"
                   sql = sql + "                                  (Incentivo.CodTipoIncentivo = '15') OR"
                   sql = sql + "                                  (Incentivo.CodTipoIncentivo = '16') OR"
                   sql = sql + "                                  (Incentivo.CodTipoIncentivo = '17') OR"
                   sql = sql + "                                  (Incentivo.CodTipoIncentivo = '18') OR"
                   sql = sql + "                                  (Incentivo.CodTipoIncentivo = '19')"
                   sql = sql + "            GROUP BY DetalleIncentivo.NumNomina, Incentivo.CodEmpleado"
                   sql = sql + "            HAVING      (DetalleIncentivo.NumNomina = DetalleNomina.NumNomina) AND (Incentivo.CodEmpleado = DetalleNomina.CodEmpleado)), 0) AS TotalDevengado,"
                   sql = sql + "    DetalleNomina.INATEC AS MontoInatec, Empleado.NumeroInss, DetalleNomina.MontoINSS + DetalleNomina.INSSPatronal AS TotalInss, Empleado.CodEmpleado1, DetalleNomina.NumNomina,"
                   sql = sql + "    Nomina.FechaNomina AS Nomina, Cargo.Cargo, Nomina.Mes, Nomina.Ano, Nomina.Periodo, TipoNomina.Periodo AS PeriodoNomina, Empleado.NumCedula, DetalleNomina.AjusteINSS,"
                   sql = sql + "    Empleado.TarifaHoraria , departamento.CodDepartamento, departamento.departamento"
                   sql = sql + "  FROM         Nomina INNER JOIN"
                   sql = sql + "    Grupo INNER JOIN"
                   sql = sql + "    Cargo INNER JOIN"
                   sql = sql + "    TipoNomina INNER JOIN"
                   sql = sql + "    Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN"
                   sql = sql + "    DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND"
                   sql = sql + "    Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento "
                   sql = sql + "  WHERE     (Nomina.FechaNomina BETWEEN '" & Format(Me.TxtFecha1.Value, "yyyymmdd") & "' AND '" & Format(Me.TxtFecha2.Value, "yyyymmdd") & "') AND (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "')"
                   sql = sql + "  ORDER BY Empleado.CodEmpleado1, Nomina"
     
 
'                   sql = "SELECT Empleado.Nombre1, Empleado.Apellido1, DetalleNomina.CodEmpleado, DetalleNomina.MontoINSS AS MontoInss, DetalleNomina.INSSPatronal AS InssPatronal, " & _
'                         "DetalleNomina.SalarioBasico DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas " & _
'                         "+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad - ISNULL " & _
'                         "    ((SELECT        SUM(DetalleIncentivo.Valor) AS Valor " & _
'                         "       FROM            Incentivo INNER JOIN " & _
'                         "                                DetalleIncentivo ON Incentivo.NumIncentivo = DetalleIncentivo.NumIncentivo " & _
'                         "       WHERE        (Incentivo.CodTipoIncentivo = '14') OR " & _
'                         "                                 (Incentivo.CodTipoIncentivo = '15') OR " & _
'                         "                                 (Incentivo.CodTipoIncentivo = '16') OR " & _
'                         "                                 (Incentivo.CodTipoIncentivo = '17') OR " & _
'                         "                                 (Incentivo.CodTipoIncentivo = '18') OR " & _
'                         "                                 (Incentivo.CodTipoIncentivo = '19') " & _
'                         "        GROUP BY DetalleIncentivo.NumNomina, Incentivo.CodEmpleado " & _
'                         "        HAVING        (DetalleIncentivo.NumNomina = DetalleNomina.NumNomina) AND (Incentivo.CodEmpleado = DetalleNomina.CodEmpleado)), 0) AS TotalDevengado, " & _
'                         "DetalleNomina.INATEC AS MontoInatec, Empleado.NumeroInss, DetalleNomina.MontoINSS + DetalleNomina.INSSPatronal AS TotalInss, Empleado.CodEmpleado1, " & _
'                         "DetalleNomina.NumNomina, Nomina.FechaNomina AS Nomina, Cargo.Cargo, Nomina.Mes, Nomina.Ano, Nomina.Periodo, TipoNomina.Periodo AS PeriodoNomina, " & _
'                         "Empleado.NumCedula , DetalleNomina.AjusteINSS, Empleado.TarifaHoraria, departamento.CodDepartamento, departamento.departamento " & _
'                         "FROM            Nomina INNER JOIN " & _
'                         "Grupo INNER JOIN " & _
'                         "Cargo INNER JOIN " & _
'                         "TipoNomina INNER JOIN " & _
'                         "Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN " & _
'                         "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON " & _
'                         "TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN " & _
'                         "Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento WHERE   (Nomina.FechaNomina BETWEEN '20181021' AND '20181117') AND (Nomina.CodTipoNomina = '02') ORDER BY Empleado.CodEmpleado1, Nomina"
         
' '//////////////////////////////////////////////////////////////////////////////////////////////
' '/////////////////////BUSCO EL MONTO DEL INSS PARA LAS VACACIONES/////////////////////////////
' '/////////////////////////////////////////////////////////////////////////////////////////////

' SqlVacaciones = "SELECT     NomVaca.NumNomVaca, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, " & _
'                 "DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, " & _
'                 "DetalleNomVaca.SalarioMensual * (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento)/ 30 - DetalleNomVaca.AdelantoVacaciones AS MontoAPagar " & _
'                 ", NomVaca.FechaAplica, NomVaca.CodTipoNomina, DetalleNomVaca.Inss,DetalleNomVaca.CodEmpleado " & _
'                 "FROM  NomVaca INNER JOIN " & _
'                 "Empleado INNER JOIN " & _
'                 "DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca " & _
'                 "WHERE     (NomVaca.FechaAplica BETWEEN '" & Format(Me.TxtFecha1.Value, "yyyymmdd") & "' AND '" & Format(Me.TxtFecha2.Value, "yyyymmdd") & "') " & _
'                 " AND (NomVaca.CodTipoNomina =  '" & Me.DBTipoNominas.Columns(0).Text & "') AND (DetalleNomVaca.CodEmpleado = 1235)" & _
'                 "ORDER BY Empleado.CodEmpleado1 "
'
'
'
' Me.AdoVacaciones.RecordSource = SqlVacaciones
' Me.AdoVacaciones.Refresh
 
 
 
 
    
    Me.Barra.Visible = True
    Me.AdoBusca.RecordSource = sql
    AdoBusca.Refresh
    If AdoBusca.Recordset.EOF Then
         MsgBox "No Existen Registros", vbCritical, "Sistema de Nominas"
         Me.Barra.Visible = False
         Exit Sub
    End If
    
    Me.AdoBusca.Recordset.MoveLast
    Maximo = AdoBusca.Recordset.RecordCount
    Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
'    Call Formato_Excel(8, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
V = 2
H = 0
i = 1

'///////////////////////////////////////////////////////////////////////////////////////
'////////////////////ENCABEZADOS//////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////

            objExcel.ActiveSheet.Cells(1, 1) = "INSSBI"
'            objExcel.ActiveSheet.Cells(1, 2) = "CODIGO"
'            objExcel.ActiveSheet.Cells(1, 3) = "CEDULA"
            objExcel.ActiveSheet.Cells(1, 2) = "NOMBRES"
            objExcel.ActiveSheet.Cells(1, 3) = "APELLIDOS"
            objExcel.ActiveSheet.Cells(1, 4) = "NOMINA"
            objExcel.ActiveSheet.Cells(1, 5) = "TIPO DE NOVEDAD"
            objExcel.ActiveSheet.Cells(1, 6) = "FECHA DE NOVEDAD"
'            objExcel.ActiveSheet.Cells(1, 9) = "SEMANA1"
'            objExcel.ActiveSheet.Cells(1, 10) = "AJUSTE S1"
'            objExcel.ActiveSheet.Cells(1, 11) = "SEMANA2"
'            objExcel.ActiveSheet.Cells(1, 12) = "AJUSTE S2"
'            objExcel.ActiveSheet.Cells(1, 13) = "SEMANA3"
'            objExcel.ActiveSheet.Cells(1, 14) = "AJUSTE S3"
'            objExcel.ActiveSheet.Cells(1, 15) = "SEMANA4"
'            objExcel.ActiveSheet.Cells(1, 16) = "AJUSTE S4"
'            objExcel.ActiveSheet.Cells(1, 17) = "SEMANA5"
'            objExcel.ActiveSheet.Cells(1, 18) = "AJUSTE S5"
            objExcel.ActiveSheet.Cells(1, 7) = "SALDEVENGADO"
            objExcel.ActiveSheet.Cells(1, 8) = "SALMENSUAL"
'            objExcel.ActiveSheet.Cells(1, 21) = "SALVAC"
            objExcel.ActiveSheet.Cells(1, 9) = "APORTE"
            objExcel.ActiveSheet.Cells(1, 10) = "SEMANAS"
            objExcel.ActiveSheet.Cells(1, 11) = "CENTRO DE COSTO"
            objExcel.ActiveSheet.Cells(1, 12) = "TIPEMPL"
            objExcel.ActiveSheet.Cells(1, 12) = "DEPARTAMENTO"
            
            
            
            


            

     Me.AdoBusca.Refresh
     If Not Me.AdoBusca.Recordset.EOF Then
      CodEmpleado = AdoBusca.Recordset("CodEmpleado")
     End If
     
    If Not IsNull(AdoBusca.Recordset("TarifaHoraria")) Then
      TarifaHorariaBasico = AdoBusca.Recordset("TarifaHoraria")
    
    End If
     
     ReDim Semana(5) As Variant
      For j = 1 To 5
        Semana(j) = "0"
      Next
      '/////////////////////DEFINO LA MATRIZ PARA LOS AJUSTES ///////////////////////////
      ReDim SemanaAjuste(5) As Double
      For j = 1 To 5
        SemanaAjuste(j) = "0"
      Next
      

        cnDB.ConnectionString = Conexion
        cnDB.Open

        rsBD.Open "SELECT * FROM Fecha_Planilla WHERE Inicio >= CONVERT(DATETIME, '" & Format(Me.TxtFecha1.Value, "yyyy-mm-dd") & " 00:00:00', 102) AND Final <=CONVERT(DATETIME, '" & Format(Me.TxtFecha2.Value, "yyyy-mm-dd") & " 00:00:00', 102) AND CodTipoNomina ='" & Me.DBTipoNominas.Columns(0).Text & "' ORDER BY Periodo ASC", cnDB

        iCont = 1

        Do While Not rsBD.EOF
            saPeriodos(iCont) = rsBD.Fields("Periodo")
            iCont = iCont + 1
   
            rsBD.MoveNext
        Loop

    rsBD.Close
    cnDB.Close


     
  Do While Not Me.AdoBusca.Recordset.EOF 'esto nos sirve pa leer los datos desde
            If Not IsNull(Me.AdoBusca.Recordset("Periodo")) Then
               Periodo = Me.AdoBusca.Recordset("Periodo")
            End If
            
            If Not IsNull(Me.AdoBusca.Recordset("PeriodoNomina")) Then
               PeriodoNomina = Me.AdoBusca.Recordset("PeriodoNomina")
            End If
            
           If Not IsNull(Me.AdoBusca.Recordset("TarifaHoraria")) Then
                TarifaHorariaBasico = Me.AdoBusca.Recordset("TarifaHoraria")
           End If
       
       
            CantSabados = SemanasPeriodos(Me.AdoBusca.Recordset("Ano"), Me.AdoBusca.Recordset("Mes"), CodTipoNomina)

             If AdoBusca.Recordset("CodEmpleado") = 10825 Then
              AdoBusca.Recordset("CodEmpleado") = 10825
             End If
       
             SalarioVaca = 0
             MontoInssVaca = 0
             
'           SqlVacaciones = "SELECT  MAX(NomVaca.NumNomVaca) AS NumNom13Mes, SUM(DetalleNomVaca.Inss) AS Inss, Empleado.CodEmpleado1, MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, SUM(DetalleNomVaca.SalarioMensual) AS SalarioMensual, SUM(DetalleNomVaca.DiasAPagar) AS DiasAPagar, SUM(DetalleNomVaca.DiasDescuento) AS DiasDescuento, SUM(DetalleNomVaca.AdelantoVacaciones) AS Adelanto13vo, SUM(DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss) AS MontoAPagar, SUM(DetalleNomVaca.TotalDevengado) AS TotalDevengado, MAX(Historico.FechaContrato) AS FechaContrato, MAX(Empleado.TarifaHoraria) AS TarifaHoraria, SUM(DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss) AS TotalDeducir, MAX(NomVaca.CodTipoNomina) AS CodTipoNomina, MAX(NomVaca.FechaAplica) As FechaAplica " & _
'                 "FROM  NomVaca INNER JOIN Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado  " & _
'                 "Where (Empleado.CodEmpleado = " & AdoBusca.Recordset("CodEmpleado") & ") GROUP BY Empleado.CodEmpleado1 HAVING (MAX(NomVaca.CodTipoNomina) = '" & Me.DBTipoNominas.Columns(0).Text & "') AND (MAX(NomVaca.FechaAplica) BETWEEN CONVERT(DATETIME, '" & Format(Me.TxtFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME,'" & Format(Me.TxtFecha2.Value, "yyyy-mm-dd") & "', 102)) ORDER BY Empleado.CodEmpleado1"

           SqlVacaciones = "SELECT  NomVaca.NumNomVaca, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, DetalleNomVaca.SalarioMensual * (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento) / '30' - DetalleNomVaca.AdelantoVacaciones AS MontoAPagar, NomVaca.CodTipoNomina , Empleado.CodEmpleado, NomVaca.FechaAplica, DetalleNomVaca.Inss FROM   NomVaca INNER JOIN Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca  " & _
                           "WHERE  (NomVaca.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') AND (Empleado.CodEmpleado = " & AdoBusca.Recordset("CodEmpleado") & ") AND (NomVaca.FechaAplica BETWEEN CONVERT(DATETIME, '" & Format(Me.TxtFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.TxtFecha2.Value, "yyyy-mm-dd") & "', 102)) ORDER BY Empleado.CodEmpleado1, NomVaca.FechaAplica DESC"
           
           Me.AdoVacaciones.RecordSource = SqlVacaciones
           Me.AdoVacaciones.Refresh
           If Not Me.AdoVacaciones.Recordset.EOF Then
             If Not IsNull(Me.AdoVacaciones.Recordset("Inss")) Then
              MontoInssVaca = Me.AdoVacaciones.Recordset("Inss")
             End If
             If Not IsNull(Me.AdoVacaciones.Recordset("MontoAPagar")) Then
               SalarioVaca = Me.AdoVacaciones.Recordset("MontoAPagar")
             End If
             
            '//////////////////////////////////////////SI DIA ES CERO NO TIENE SALARIO VACA/////////////////////
            If Not IsNull(Me.AdoVacaciones.Recordset("DiasAPagar")) Then
             If Me.AdoVacaciones.Recordset("DiasAPagar") = 0 Then
               SalarioVaca = 0
               MontoInssVaca = 0
             End If
            End If
           Else
             SalarioVaca = 0
             MontoInssVaca = 0
           End If
           
           
           

 'la tabla de access para despues colocarlos en las celdas correspondientes
            
  If CodEmpleado = AdoBusca.Recordset("CodEmpleado") Then
        
'            SqlVacaciones = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo, DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoAPagar, DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina , NomVaca.FechaAplica  " & _
'                   "FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado  " & _
'                   "WHERE (NomVaca.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') AND (DetalleNomVaca.CodEmpleado = " & CodEmpleado & ") AND (NomVaca.FechaAplica BETWEEN '" & Format(Me.TxtFecha1.Value, "yyyymmdd") & "' AND '" & Format(Me.TxtFecha2.Value, "yyyymmdd") & "') ORDER BY Empleado.CodEmpleado1"

       
'           SqlVacaciones = "SELECT  MAX(NomVaca.NumNomVaca) AS NumNom13Mes, SUM(DetalleNomVaca.Inss) AS Inss, Empleado.CodEmpleado1, MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, SUM(DetalleNomVaca.SalarioMensual) AS SalarioMensual, SUM(DetalleNomVaca.DiasAPagar) AS DiasAPagar, SUM(DetalleNomVaca.DiasDescuento) AS DiasDescuento, SUM(DetalleNomVaca.AdelantoVacaciones) AS Adelanto13vo, SUM(DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss) AS MontoAPagar, SUM(DetalleNomVaca.TotalDevengado) AS TotalDevengado, MAX(Historico.FechaContrato) AS FechaContrato, MAX(Empleado.TarifaHoraria) AS TarifaHoraria, SUM(DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss) AS TotalDeducir, MAX(NomVaca.CodTipoNomina) AS CodTipoNomina, MAX(NomVaca.FechaAplica) As FechaAplica " & _
'                 "FROM  NomVaca INNER JOIN Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado  " & _
'                 "Where (Empleado.CodEmpleado = " & CodEmpleado & ") GROUP BY Empleado.CodEmpleado1 HAVING (MAX(NomVaca.CodTipoNomina) = '" & Me.DBTipoNominas.Columns(0).Text & "') AND (MAX(NomVaca.FechaAplica) BETWEEN CONVERT(DATETIME, '" & Format(Me.TxtFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME,'" & Format(Me.TxtFecha2.Value, "yyyy-mm-dd") & "', 102)) ORDER BY Empleado.CodEmpleado1"

         SqlVacaciones = "SELECT  NomVaca.NumNomVaca, Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones, DetalleNomVaca.SalarioMensual * (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento) / '30' - DetalleNomVaca.AdelantoVacaciones AS MontoAPagar, NomVaca.CodTipoNomina , Empleado.CodEmpleado, NomVaca.FechaAplica, DetalleNomVaca.Inss FROM   NomVaca INNER JOIN Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca  " & _
                           "WHERE  (NomVaca.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') AND (Empleado.CodEmpleado = " & AdoBusca.Recordset("CodEmpleado") & ") AND (NomVaca.FechaAplica BETWEEN CONVERT(DATETIME, '" & Format(Me.TxtFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.TxtFecha2.Value, "yyyy-mm-dd") & "', 102)) ORDER BY Empleado.CodEmpleado1, NomVaca.FechaAplica DESC"
                           
                           
           Me.AdoVacaciones.RecordSource = SqlVacaciones
           Me.AdoVacaciones.Refresh
           If Not Me.AdoVacaciones.Recordset.EOF Then
             If Not IsNull(Me.AdoVacaciones.Recordset("Inss")) Then
              MontoInssVaca = Me.AdoVacaciones.Recordset("Inss")
             End If
             If Not IsNull(Me.AdoVacaciones.Recordset("MontoAPagar")) Then
               SalarioVaca = Me.AdoVacaciones.Recordset("MontoAPagar")
             End If
             
            '//////////////////////////////////////////SI DIA ES CERO NO TIENE SALARIO VACA/////////////////////
            If Not IsNull(Me.AdoVacaciones.Recordset("DiasAPagar")) Then
             If Me.AdoVacaciones.Recordset("DiasAPagar") = 0 Then
               SalarioVaca = 0
               MontoInssVaca = 0
             End If
            End If
           Else
             SalarioVaca = 0
             MontoInssVaca = 0
           End If
           
           If CodEmpleado = "11328" Then
             CodEmpleado = "11328"
           End If

        
          TotalDevengado = Me.AdoBusca.Recordset("TotalDevengado") + TotalDevengado
'          TotalAjuste = TotalAjuste + Me.AdoBusca.Recordset("AjusteINSS")
          MontoInss = Me.AdoBusca.Recordset("MontoInss")
          
         MontoInssBasico = ((TarifaHorariaBasico * 8 * DiasMes) * (TasaInss / 100) / CantSabados)
         
         'If MontoInss > MontoInssBasico Then
         MontoBasico = (TarifaHorariaBasico * 8 * DiasMes) / CantSabados
         If Me.AdoBusca.Recordset("TotalDevengado") > MontoBasico Then
           AjusteINSS = 0
         Else
'           AjusteINSS = MontoInssBasico - MontoInss
            AjusteINSS = MontoBasico - Me.AdoBusca.Recordset("TotalDevengado")
         End If
          SemanaAjuste(i) = AjusteINSS
          
          Select Case PeriodoNomina
           Case "Semanal Viernes"
'            Select Case i 'Periodo
'             Case saPeriodos(1): Semana(1) = "1"
'             Case saPeriodos(2): Semana(2) = "1"
'             Case saPeriodos(3): Semana(3) = "1"
'             Case saPeriodos(4): Semana(4) = "1"
'             Case saPeriodos(5): Semana(5) = "1"
'            End Select
            Select Case i 'Periodo
             Case 1: If Me.AdoBusca.Recordset("TotalDevengado") <> 0 Then Semana(1) = "1" Else Semana(1) = "0"
             Case 2: If Me.AdoBusca.Recordset("TotalDevengado") <> 0 Then Semana(2) = "1" Else Semana(2) = "0"
             Case 3: If Me.AdoBusca.Recordset("TotalDevengado") <> 0 Then Semana(3) = "1" Else Semana(3) = "0"
             Case 4: If Me.AdoBusca.Recordset("TotalDevengado") <> 0 Then Semana(4) = "1" Else Semana(4) = "0"
             Case 5: If Me.AdoBusca.Recordset("TotalDevengado") <> 0 Then Semana(5) = "1" Else Semana(5) = "0"
            End Select
           Case "Semanal Sabado"

            Select Case i 'Periodo
             Case 1: If Me.AdoBusca.Recordset("TotalDevengado") <> 0 Then Semana(1) = "1" Else Semana(1) = "0"
             Case 2: If Me.AdoBusca.Recordset("TotalDevengado") <> 0 Then Semana(2) = "1" Else Semana(2) = "0"
             Case 3: If Me.AdoBusca.Recordset("TotalDevengado") <> 0 Then Semana(3) = "1" Else Semana(3) = "0"
             Case 4: If Me.AdoBusca.Recordset("TotalDevengado") <> 0 Then Semana(4) = "1" Else Semana(4) = "0"
             Case 5: If Me.AdoBusca.Recordset("TotalDevengado") <> 0 Then Semana(5) = "1" Else Semana(5) = "0"
            End Select
            
           Case "Quincenal"
            Select Case i 'Periodo
             Case saPeriodos(1)
                Semana(1) = "1"
                Semana(2) = "1"
             Case saPeriodos(2)
                Semana(3) = "1"
                Semana(4) = "1"
            
            End Select
           
          End Select
            Semanas = Semana(1) & Semana(2) & Semana(3) & Semana(4) & Semana(5)
          
         
        
           CodEmpleado = AdoBusca.Recordset("CodEmpleado")


     
       
        '/////////////////////////TH//////////////////////////////////
        Me.AdoTarifa.RecordSource = "SELECT   CodEmpleado, CodEmpleado1, Liquidado, TarifaHoraria,SueldoPeriodo From Empleado Where (CodEmpleado = " & CodEmpleado & ")"
        Me.AdoTarifa.Refresh
        If Not Me.AdoTarifa.Recordset.EOF Then
             TipoNomina = Me.DBTipoNominas.Columns(2).Text
            Select Case TipoNomina
                Case "Quincenal"
                   TarifaHoraria = Me.AdoTarifa.Recordset("TarifaHoraria")
                   SueldoPeriodo = Me.AdoTarifa.Recordset("SueldoPeriodo") * 2
                Case "Mensual"
                   TarifaHoraria = Me.AdoTarifa.Recordset("TarifaHoraria")
                   SueldoPeriodo = Me.AdoTarifa.Recordset("SueldoPeriodo")
                Case Else
                    TarifaHoraria = Me.AdoTarifa.Recordset("TarifaHoraria")
                    SueldoPeriodo = 30.4167 * TarifaHoraria * 8
                End Select
        End If
        
        '//////////////////////////////////////////////////////////////////////
        '/////ASIGNO LA NOVEDAD 03 A TODOS PARA MIENTRAS CAMBIAN/////////////
        '/////////////////////////////////////////////////////////////////////
           Novedad = "03"
          DiaFinal = DateSerial(Year(Me.TxtFecha2.Value), Month(Me.TxtFecha2.Value) + 1, 1 - 1)

           
           If TotalDevengado = (TarifaHoraria * 8 * 30.4167) Then
             Novedad = 0
           Else
             Novedad = 3
             
           End If
           
        '/////////////////////////////////////////////////////////////////////////
        '////////////////BUSCO SI EL EMPLEADO ES NUEVO INGRESO///////////////////
        '/////////////////////////////////////////////////////////////////////////
        Me.AdoNuevoIngreso.RecordSource = "SELECT Codempleado, FechaContrato " & _
        "From Historico WHERE  (Codempleado = " & CodEmpleado & ") " & _
        "AND (FechaContrato BETWEEN '" & Format(Me.TxtFecha1.Value, "yyyymmdd") & "' And '" & Format(Me.TxtFecha2.Value, "yyyymmdd") & "')"
        
        Me.AdoNuevoIngreso.Refresh
        If Not Me.AdoNuevoIngreso.Recordset.EOF Then
         If Not IsNull(Me.AdoNuevoIngreso.Recordset("FechaContrato")) Then
           FechaContrato = Me.AdoNuevoIngreso.Recordset("FechaContrato")
           
              Novedad = "01"
              DiaFinal = FechaContrato
         End If
        End If
           
           
        '///////////////////////////////////////////////////////////////////////////////////
        '/////////BUSCP SI EL EMPLEADO ESTA DADO DE BAJA///////////////
        '//////////////////////////////////////////////////////////////
                  
         Me.AdoBajas.RecordSource = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Activo, Bajas.FechaBaja, Empleado.Liquidado " & _
                                   "FROM  Empleado INNER JOIN Bajas ON Empleado.CodEmpleado = Bajas.CodEmpleado " & _
                                   "Where (Empleado.Activo = 0) And (Empleado.CodEmpleado = " & CodEmpleado & ") "
        'Me.AdoBajas.RecordSource = "SELECT  CodEmpleado, CodEmpleado1, Activo From Empleado WHERE (CodEmpleado = " & CodEmpleado & ")"
        Me.AdoBajas.Refresh
        If Me.AdoBajas.Recordset.EOF Then
'          Novedad = "03"
'          DiaFinal = DateSerial(Year(Me.TxtFecha2.Value), Month(Me.TxtFecha2.Value) + 1, 1 - 1)
        Else
          Novedad = "02"
          DiaFinal = Me.AdoBajas.Recordset("FechaBaja")
          SueldoPeriodo = 0
        End If
        


           
          '//////////////////////////////////////////////////////////////////////////////////
          '////SI EL EMPLEADO NO TIENE SALARIO DEVENGADO Y SUELDO DEL PERIODO////////////////
          '///LA NOVEDAD ES 9 Y LAS SEMANAS ES "00000"//////////////////////////////////////
           
           If val(TotalDevengado) = 0 And val(SueldoPeriodo) = 0 Then
             agregar = False
'              Novedad = "88"
              Semanas = "00000"
           Else
             agregar = True
           End If
           
           
           '///////////////////////////////////////////////////////////////////////////////
           '//////BUSCO LAS SUSPENCIONES DE LOS EMPLEADOS PARA MANTENER/////////////////////
           '/////LA NOVEDAD 09 ////////////////////////////////////////////////////////////
            Me.AdoSuspenciones.RecordSource = "SELECT CodEmpleado, CodEmpleado1, Fechaini, FechaFin, Motivo, Activo, Ultimo " & _
            "From Subsidios Where (Activo = 1) And (CodEmpleado = " & CodEmpleado & ")"
            Me.AdoSuspenciones.Refresh
            If Not Me.AdoSuspenciones.Recordset.EOF Then
              Novedad = "09"
            End If
            

            
            If agregar = True Then
                objExcel.ActiveSheet.Columns("D").NumberFormat = "0#"
                objExcel.ActiveSheet.Columns("E").NumberFormat = "0#"
               objExcel.ActiveSheet.Columns("F").NumberFormat = "yyyy-MM-dd"
                objExcel.ActiveSheet.Columns("L").NumberFormat = "0.00"
                objExcel.ActiveSheet.Columns("J").NumberFormat = "@"
                objExcel.ActiveSheet.Columns("J").HorizontalAlignment = 3
                objExcel.ActiveSheet.Cells(V, H + 1) = Me.AdoBusca.Recordset("NumeroInss")
'                objExcel.ActiveSheet.Cells(V, H + 2) = Me.AdoBusca.Recordset("CodEmpleado1")
'                objExcel.ActiveSheet.Cells(V, H + 3) = Me.AdoBusca.Recordset("NumCedula")
                objExcel.ActiveSheet.Cells(V, H + 2) = Me.AdoBusca.Recordset("Nombre1")
                objExcel.ActiveSheet.Cells(V, H + 3) = Me.AdoBusca.Recordset("Apellido1")
                objExcel.ActiveSheet.Cells(V, H + 4) = 1
                objExcel.ActiveSheet.Cells(V, H + 5) = Novedad
                objExcel.ActiveSheet.Cells(V, H + 6) = Format(DiaFinal, "yyyy-mm-dd")
'                Select Case i
'                   Case 1: objExcel.ActiveSheet.Cells(V, H + 9) = Me.AdoBusca.Recordset("TotalDevengado")
'                   Case 2: objExcel.ActiveSheet.Cells(V, H + 11) = Me.AdoBusca.Recordset("TotalDevengado")
'                   Case 3: objExcel.ActiveSheet.Cells(V, H + 13) = Me.AdoBusca.Recordset("TotalDevengado")
'                   Case 4: objExcel.ActiveSheet.Cells(V, H + 15) = Me.AdoBusca.Recordset("TotalDevengado")
'                   Case 5: objExcel.ActiveSheet.Cells(V, H + 17) = Me.AdoBusca.Recordset("TotalDevengado")
'                End Select
                objExcel.ActiveSheet.Cells(V, H + 7) = Format(TotalDevengado + SalarioVaca, "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 8) = Format(SueldoPeriodo, "##,##0.00")
'                objExcel.ActiveSheet.Cells(V, H + 21) = Format(SalarioVaca, "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 9) = Format(MontoInss + MontoInssVaca, "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 10) = Semanas
               objExcel.ActiveSheet.Cells(V, H + 11) = Me.AdoBusca.Recordset("Departamento")
'                objExcel.ActiveSheet.Cells(V, H + 25) = ""
'                Select Case i
'                   Case 1: objExcel.ActiveSheet.Cells(V, H + 10) = Format(SemanaAjuste(i), "##,##0.00")
'                   Case 2: objExcel.ActiveSheet.Cells(V, H + 12) = Format(SemanaAjuste(i), "##,##0.00")
'                   Case 3: objExcel.ActiveSheet.Cells(V, H + 14) = Format(SemanaAjuste(i), "##,##0.00")
'                   Case 4: objExcel.ActiveSheet.Cells(V, H + 16) = Format(SemanaAjuste(i), "##,##0.00")
'                   Case 5: objExcel.ActiveSheet.Cells(V, H + 18) = Format(SemanaAjuste(i), "##,##0.00")
'                End Select
                objExcel.ActiveSheet.Cells(V, H + 12) = Me.AdoBusca.Recordset("NumeroInss") & ";" & Me.AdoBusca.Recordset("Nombre1") & ";" & Me.AdoBusca.Recordset("Apellido1") & ";" & "01" & ";" & Novedad & ";" & Format(Me.TxtFecha2.Value, "yyyy-mm-dd") & ";" & Format(TotalDevengado, "####0.00") & ";" & Format((30.4167 * TarifaHoraria * 8), "####0.00") & ";" & "0.00" & ";" & Semanas & ";" & "0"
                

            
            End If
            i = i + 1
            Me.AdoBusca.Recordset.MoveNext

         Else
         
        
                  
                 If agregar = True Then
                   V = V + 1
                 End If
                 
                 i = 1
                  
                For j = 1 To 5
                   Semana(j) = "0"
                Next
            
                Select Case PeriodoNomina
                 Case "Semanal Viernes"
                Select Case i
                 Case 1: If Me.AdoBusca.Recordset("TotalDevengado") <> 0 Then Semana(1) = "1" Else Semana(1) = "0"
                 Case 2: If Me.AdoBusca.Recordset("TotalDevengado") <> 0 Then Semana(2) = "1" Else Semana(2) = "0"
                 Case 3: If Me.AdoBusca.Recordset("TotalDevengado") <> 0 Then Semana(3) = "1" Else Semana(3) = "0"
                 Case 4: If Me.AdoBusca.Recordset("TotalDevengado") <> 0 Then Semana(4) = "1" Else Semana(4) = "0"
                 Case 5: If Me.AdoBusca.Recordset("TotalDevengado") <> 0 Then Semana(5) = "1" Else Semana(5) = "0"
                End Select
                 
                 Case "Quincenal"
                  Select Case i 'Periodo
                   Case saPeriodos(1)
                      Semana(1) = "1"
                      Semana(2) = "1"
                   Case saPeriodos(2)
                      Semana(3) = "1"
                      Semana(4) = "1"
                  
                  End Select
                 
                End Select
            
                   Semanas = Semana(1) & Semana(2) & Semana(3) & Semana(4) & Semana(5)
        
           '------------------------------------------------------------------------------------------------------------------------------
           '------------------------------------------LLENO DE CERO LOS PRIMEROS REGISTROS -------------------------------------------------
           '--------------------------------------------------------------------------------------------------------------------------------
           
              If i = 1 Then
                objExcel.ActiveSheet.Columns("D").NumberFormat = "0#"
                objExcel.ActiveSheet.Columns("F").NumberFormat = "yyyy-MM-dd"
                objExcel.ActiveSheet.Columns("E").NumberFormat = "0#"
                objExcel.ActiveSheet.Columns("L").NumberFormat = "0.00"
                objExcel.ActiveSheet.Columns("J").NumberFormat = "@"
                objExcel.ActiveSheet.Columns("J").HorizontalAlignment = 3
                objExcel.ActiveSheet.Cells(V, H + 1) = Me.AdoBusca.Recordset("NumeroInss")
'                objExcel.ActiveSheet.Cells(V, H + 2) = Me.AdoBusca.Recordset("CodEmpleado1")
'                objExcel.ActiveSheet.Cells(V, H + 3) = Me.AdoBusca.Recordset("NumCedula")
                objExcel.ActiveSheet.Cells(V, H + 2) = Me.AdoBusca.Recordset("Nombre1")
                objExcel.ActiveSheet.Cells(V, H + 3) = Me.AdoBusca.Recordset("Apellido1")
                objExcel.ActiveSheet.Cells(V, H + 4) = 1
                objExcel.ActiveSheet.Cells(V, H + 5) = 0
'                objExcel.ActiveSheet.Cells(V, H + 9) = Format(0, "##,##0.00")
'                objExcel.ActiveSheet.Cells(V, H + 10) = Format(0, "##,##0.00")
'                objExcel.ActiveSheet.Cells(V, H + 11) = Format(0, "##,##0.00")
'                objExcel.ActiveSheet.Cells(V, H + 12) = Format(0, "##,##0.00")
'                objExcel.ActiveSheet.Cells(V, H + 13) = Format(0, "##,##0.00")
'                objExcel.ActiveSheet.Cells(V, H + 14) = Format(0, "##,##0.00")
'                objExcel.ActiveSheet.Cells(V, H + 15) = Format(0, "##,##0.00")
'                objExcel.ActiveSheet.Cells(V, H + 16) = Format(0, "##,##0.00")
'                objExcel.ActiveSheet.Cells(V, H + 17) = Format(0, "##,##0.00")
'                objExcel.ActiveSheet.Cells(V, H + 18) = Format(0, "##,##0.00")
'                objExcel.ActiveSheet.Cells(V, H + 19) = Format(0, "##,##0.00")
'                objExcel.ActiveSheet.Cells(V, H + 20) = Format(0, "##,##0.00")
'                objExcel.ActiveSheet.Cells(V, H + 21) = Format(0, "##,##0.00")
'                objExcel.ActiveSheet.Cells(V, H + 22) = "0.00"
'                objExcel.ActiveSheet.Cells(V, H + 23) = "0"
'                objExcel.ActiveSheet.Cells(V, H + 24) = Me.AdoBusca.Recordset("Cargo")
'                objExcel.ActiveSheet.Cells(V, H + 25) = ""
                objExcel.ActiveSheet.Cells(V, H + 10) = Format(0, "##,##0.00")
              End If
        
                
                   TotalDevengado = Me.AdoBusca.Recordset("TotalDevengado")
'                   TotalAjuste = Me.AdoBusca.Recordset("AjusteINSS")
                   CodEmpleado = AdoBusca.Recordset("CodEmpleado")

                     MontoInss = Me.AdoBusca.Recordset("MontoInss")
                     
                    MontoInssBasico = ((TarifaHorariaBasico * 8 * DiasMes) * (TasaInss / 100) / CantSabados)
                     'If MontoInss > MontoInssBasico Then
                     MontoBasico = TarifaHorariaBasico * 56
                     If Me.AdoBusca.Recordset("TotalDevengado") > MontoBasico Then
                       AjusteINSS = 0
                     Else
            '           AjusteINSS = MontoInssBasico - MontoInss
                        AjusteINSS = (TarifaHorariaBasico * 56) - Me.AdoBusca.Recordset("TotalDevengado")
                     End If
                     SemanaAjuste(i) = AjusteINSS
           
                '/////////////////////////TH//////////////////////////////////
                Me.AdoTarifa.RecordSource = "SELECT   CodEmpleado, CodEmpleado1, Liquidado, TarifaHoraria,SueldoPeriodo From Empleado Where (CodEmpleado = " & CodEmpleado & ")"
                Me.AdoTarifa.Refresh
                If Not Me.AdoTarifa.Recordset.EOF Then
                     TipoNomina = Me.DBTipoNominas.Columns(2).Text
                    Select Case TipoNomina
                        Case "Quincenal"
                           SueldoPeriodo = Me.AdoTarifa.Recordset("SueldoPeriodo") * 2
                        Case "Mensual"
                           SueldoPeriodo = Me.AdoTarifa.Recordset("SueldoPeriodo")
                        Case Else
                            TarifaHoraria = Me.AdoTarifa.Recordset("TarifaHoraria")
                            SueldoPeriodo = 30.416667 * TarifaHoraria * 8
                        End Select
                End If
        
             '//////////////////////////////////////////////////////////////////////
             '/////ASIGNO LA NOVEDAD 03 A TODOS PARA MIENTRAS CAMBIAN/////////////
             '/////////////////////////////////////////////////////////////////////
               Novedad = "03"
               DiaFinal = DateSerial(Year(Me.TxtFecha2.Value), Month(Me.TxtFecha2.Value) + 1, 1 - 1)
             
            
            
               If TotalDevengado = (TarifaHoraria * 8 * 30.416667) Then
                  Novedad = 0
                Else
                  Novedad = 3
                  
                End If
           
             '/////////////////////////////////////////////////////////////////////////
            '////////////////BUSCO SI EL EMPLEADO ES NUEVO INGRESO///////////////////
            '/////////////////////////////////////////////////////////////////////////
            Me.AdoNuevoIngreso.RecordSource = "SELECT Codempleado, FechaContrato " & _
            "From Historico WHERE  (Codempleado = " & CodEmpleado & ") " & _
            "AND (FechaContrato BETWEEN '" & Format(Me.TxtFecha1.Value, "yyyymmdd") & "' And '" & Format(Me.TxtFecha2.Value, "yyyymmdd") & "')"
            
            Me.AdoNuevoIngreso.Refresh
            If Not Me.AdoNuevoIngreso.Recordset.EOF Then
             If Not IsNull(Me.AdoNuevoIngreso.Recordset("FechaContrato")) Then
               FechaContrato = Me.AdoNuevoIngreso.Recordset("FechaContrato")
               
                  Novedad = "01"
                  DiaFinal = FechaContrato
             End If
            End If
         
           
           
                '///////////////////////////////////////////////////////////////////////////////////
                '/////////BUSCP SI EL EMPLEADO ESTA DADO DE BAJA///////////////
                '//////////////////////////////////////////////////////////////
                Me.AdoBajas.RecordSource = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Activo, Bajas.FechaBaja, Empleado.Liquidado " & _
                                           "FROM  Empleado INNER JOIN Bajas ON Empleado.CodEmpleado = Bajas.CodEmpleado " & _
                                           "Where (Empleado.Activo = 0) And (Empleado.CodEmpleado = " & CodEmpleado & ") "
                Me.AdoBajas.Refresh
                If Me.AdoBajas.Recordset.EOF Then
        '          Novedad = "03"
        '          DiaFinal = DateSerial(Year(Me.TxtFecha2.Value), Month(Me.TxtFecha2.Value) + 1, 1 - 1)
        
                Else
                  Novedad = "02"
                  DiaFinal = Me.AdoBajas.Recordset("FechaBaja")
                  SueldoPeriodo = 0
                End If
        

           
                  '//////////////////////////////////////////////////////////////////////////////////
                  '////SI EL EMPLEADO NO TIENE SALARIO DEVENGADO Y SUELDO DEL PERIODO////////////////
                  '///LA NOVEDAD ES 9 Y LAS SEMANAS ES "00000"//////////////////////////////////////
                   
                   If val(TotalDevengado) = 0 And val(SueldoPeriodo) = 0 Then
                      agregar = False
        '              Novedad = 9
                      Semanas = "0"
                   Else
                      agregar = True
                   End If
           
                '///////////////////////////////////////////////////////////////////////////////
                '//////BUSCO LAS SUSPECIONES DE LOS EMPLEADOS PARA MANTENER/////////////////////
                '/////LA NOVEDAD 09 ////////////////////////////////////////////////////////////
                 Me.AdoSuspenciones.RecordSource = "SELECT CodEmpleado, CodEmpleado1, Fechaini, FechaFin, Motivo, Activo, Ultimo " & _
                 "From Subsidios Where (Activo = 1) And (CodEmpleado = " & CodEmpleado & ")"
                 Me.AdoSuspenciones.Refresh
                 If Not Me.AdoSuspenciones.Recordset.EOF Then
                   Novedad = "09"
                 End If
           
            
            If agregar = True Then
                objExcel.ActiveSheet.Columns("D").NumberFormat = "0#"
                objExcel.ActiveSheet.Columns("E").NumberFormat = "0#"
                objExcel.ActiveSheet.Columns("F").NumberFormat = "yyyy-MM-dd"
                objExcel.ActiveSheet.Columns("L").NumberFormat = "0.00"
                objExcel.ActiveSheet.Columns("J").NumberFormat = "@"
                objExcel.ActiveSheet.Columns("J").HorizontalAlignment = 3
                objExcel.ActiveSheet.Cells(V, H + 1) = Me.AdoBusca.Recordset("NumeroInss")
'                objExcel.ActiveSheet.Cells(V, H + 2) = Me.AdoBusca.Recordset("CodEmpleado1")
'                objExcel.ActiveSheet.Cells(V, H + 3) = Me.AdoBusca.Recordset("NumCedula")
                objExcel.ActiveSheet.Cells(V, H + 2) = Me.AdoBusca.Recordset("Nombre1")
                objExcel.ActiveSheet.Cells(V, H + 3) = Me.AdoBusca.Recordset("Apellido1")
                objExcel.ActiveSheet.Cells(V, H + 4) = 1
                objExcel.ActiveSheet.Cells(V, H + 5) = Novedad
                objExcel.ActiveSheet.Cells(V, H + 6) = Format(DiaFinal, "yyyy-MM-dd")
'                Select Case i
'                   Case 1: objExcel.ActiveSheet.Cells(V, H + 9) = Me.AdoBusca.Recordset("TotalDevengado")
'                   Case 2: objExcel.ActiveSheet.Cells(V, H + 11) = Me.AdoBusca.Recordset("TotalDevengado")
'                   Case 3: objExcel.ActiveSheet.Cells(V, H + 13) = Me.AdoBusca.Recordset("TotalDevengado")
'                   Case 4: objExcel.ActiveSheet.Cells(V, H + 15) = Me.AdoBusca.Recordset("TotalDevengado")
'                   Case 5: objExcel.ActiveSheet.Cells(V, H + 17) = Me.AdoBusca.Recordset("TotalDevengado")
'                End Select
                objExcel.ActiveSheet.Cells(V, H + 7) = Format(TotalDevengado + SalarioVaca, "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 8) = Format(SueldoPeriodo, "##,##0.00")
'                objExcel.ActiveSheet.Cells(V, H + 21) = Format(SalarioVaca, "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 9) = Format(MontoInss + MontoInssVaca, "##,##0.00")
                objExcel.ActiveSheet.Cells(V, H + 10) = Semanas
               objExcel.ActiveSheet.Cells(V, H + 11) = Me.AdoBusca.Recordset("Departamento")
'                objExcel.ActiveSheet.Cells(V, H + 25) = ""
'                Select Case i
'                   Case 1: objExcel.ActiveSheet.Cells(V, H + 10) = Format(SemanaAjuste(i), "##,##0.00")
'                   Case 2: objExcel.ActiveSheet.Cells(V, H + 12) = Format(SemanaAjuste(i), "##,##0.00")
'                   Case 3: objExcel.ActiveSheet.Cells(V, H + 14) = Format(SemanaAjuste(i), "##,##0.00")
'                   Case 4: objExcel.ActiveSheet.Cells(V, H + 16) = Format(SemanaAjuste(i), "##,##0.00")
'                   Case 5: objExcel.ActiveSheet.Cells(V, H + 18) = Format(SemanaAjuste(i), "##,##0.00")
'                End Select
                objExcel.ActiveSheet.Cells(V, H + 12) = Me.AdoBusca.Recordset("NumeroInss") & ";" & Me.AdoBusca.Recordset("Nombre1") & ";" & Me.AdoBusca.Recordset("Apellido1") & ";" & "01" & ";" & " & Novedad & " & ";" & Format(Me.TxtFecha2.Value, "yyyy-mm-dd") & ";" & Format(TotalDevengado, "####0.00") & ";" & Format((30.4167 * TarifaHoraria * 8), "####0.00") & ";" & "0.00" & ";" & Semanas & ";" & "0"
              

              
              End If
             i = i + 1
         
            Me.AdoBusca.Recordset.MoveNext
 
         End If
   

     Loop
     
  


        objExcel.ActiveSheet.Columns("E").ColumnWidth = 18
        objExcel.ActiveSheet.Columns("F").ColumnWidth = 19
'        objExcel.ActiveSheet.Columns("G").ColumnWidth = 2
'        objExcel.ActiveSheet.Columns("X").ColumnWidth = 20
        objExcel.ActiveSheet.Columns("L").ColumnWidth = 140

 
 Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
  
  
  



 Case "Listado Maestro de Empleados"
  ArepListaMaestro.DataControl1.ConnectionString = ConexionReporte
  ArepListaMaestro.lblTitulo.Caption = Titulo
  ArepListaMaestro.LblSubtitulo.Caption = "LISTADO MAESTRO DE EMPLEADOS"
  ArepListaMaestro.ImgLogo.Picture = LoadPicture(RutaLogo)
  ArepListaMaestro.LblFecha.Caption = Format(Now, "Long Date")
  ArepListaMaestro.LblDesde.Caption = Me.DTFecha1.Value
  ArepListaMaestro.LblHasta.Caption = Me.DTFecha2.Value
  
  sql = "SELECT     TOP 100 PERCENT Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico," & vbLf
  sql = sql & "                    Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
  sql = sql & "                     Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
  sql = sql & "                    Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
  sql = sql & "                    DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
  sql = sql & "                    Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
  sql = sql & "                    DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
  sql = sql & "                    DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
  sql = sql & "                    DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
  sql = sql & "                    DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
  sql = sql & "                    Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
  sql = sql & "                    DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
  sql = sql & "                     DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia AS TotalDevengado," & vbLf
  sql = sql & "                    DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
  sql = sql & "                    (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
  sql = sql & "                     DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia)" & vbLf
  sql = sql & "                    - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
  sql = sql & "                    Empleado.TarifaHoraria, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, Empleado.Activo, Empleado.NumeroInss, Empleado.CodEmpleado1," & vbLf
  sql = sql & "                    Historico.FechaContrato , Empleado.Sexo" & vbLf
  sql = sql & "FROM         Nomina INNER JOIN" & vbLf
  sql = sql & "                    Grupo INNER JOIN" & vbLf
  sql = sql & "                    Cargo INNER JOIN" & vbLf
  sql = sql & "                    TipoNomina INNER JOIN" & vbLf
  sql = sql & "                    Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
  sql = sql & "                    DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
  sql = sql & "                    TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN" & vbLf
  sql = sql & "                    Historico ON Empleado.CodEmpleado = Historico.Codempleado" & vbLf
  sql = sql & "Where (Nomina.NumNomina = '" & Me.TxtNNomina.Text & "')" & vbLf
  sql = sql & "ORDER BY Nomina.NumNomina, DetalleNomina.CodEmpleado"
  
  ArepListaMaestro.DataControl1.Source = sql
  Exportar = True
  ArepListaMaestro.Show 1
  Exportar = False

 Case "Reporte Inss 2"
    Fecha1 = Year(Me.DTFecha1.Value) & "-" & Month(Me.DTFecha1.Value) & "-" & Day(Me.DTFecha1.Value)
    Fecha2 = Year(Me.DTFecha2.Value) & "-" & Month(Me.DTFecha2.Value) & "-" & Day(Me.DTFecha2.Value)
    Fecha1Reporte = DateSerial(Year(Me.DTFecha1.Value), Month(Me.DTFecha1.Value) - 1, Day(Me.DTFecha1.Value))
    Fecha2Reporte = DateSerial(Year(Me.DTFecha1.Value), Month(Me.DTFecha1.Value), 1 - 1)
    Fecha1Reporte = Year(Fecha1Reporte) & "-" & Month(Fecha1Reporte) & "-" & Day(Fecha1Reporte)
    Fecha2Reporte = Year(Fecha2Reporte) & "-" & Month(Fecha2Reporte) & "-" & Day(Fecha2Reporte)

    Mes1 = Month(Me.DTFecha1.Value)
    ConvertirMes (Mes1)
    ArepInss2.LblMes1.Caption = "Informe del mes de " & Convertir
    ArepInss2.LblPeriodo.Caption = Convertir & " / " & Year(Me.DTFecha1.Value)

    Mes2 = Month(Fecha1Reporte)
    ConvertirMes (Mes2)
    ArepInss2.LblMes2.Caption = "Informe de " & Convertir
    
    ArepInss2.AdoNomina.ConnectionString = ConexionReporte
    ArepInss2.lblTitulo.Caption = Titulo
    ArepInss2.LblSubtitulo.Caption = "REPORTE INSS EMPLEADOS"
    ArepInss2.ImgLogo.Picture = LoadPicture(RutaLogo)
 
    sql = "SELECT     TOP 100 PERCENT dbo.Empleado.Nombre1 + N' ' + dbo.Empleado.Nombre2 + N' ' + dbo.Empleado.Apellido1 + N' ' + dbo.Empleado.Apellido2 AS Nombres," & vbLf
    sql = sql & "                       dbo.DetalleNomina.CodEmpleado, dbo.DetalleNomina.MontoINSS, dbo.DetalleNomina.INSSPatronal," & vbLf
    sql = sql & "                dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
    sql = sql & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.IncetivoProduccion AS TotalDevengado," & vbLf
    sql = sql & "                       dbo.DetalleNomina.INATEC, dbo.Empleado.CodInss, dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.INSSPatronal AS TotalInss," & vbLf
    sql = sql & "                      dbo.Empleado.CodEmpleado1 , dbo.DetalleNomina.NumNomina, dbo.Nomina.FechaNomina, dbo.Cargo.Cargo" & vbLf
    sql = sql & "FROM         dbo.Nomina INNER JOIN" & vbLf
    sql = sql & "                      dbo.Grupo INNER JOIN" & vbLf
    sql = sql & "                      dbo.Cargo INNER JOIN" & vbLf
    sql = sql & "                      dbo.TipoNomina INNER JOIN" & vbLf
    sql = sql & "                      dbo.Empleado ON dbo.TipoNomina.CodTipoNomina = dbo.Empleado.CodTipoNomina ON dbo.Cargo.CodCargo = dbo.Empleado.CodCargo INNER JOIN" & vbLf
    sql = sql & "                      dbo.DetalleNomina ON dbo.Empleado.CodEmpleado = dbo.DetalleNomina.CodEmpleado ON dbo.Grupo.CodGrupo = dbo.Empleado.CodGrupo ON" & vbLf
    sql = sql & "                      dbo.TipoNomina.CodTipoNomina = dbo.Nomina.CodTipoNomina And dbo.Nomina.NumNomina = dbo.DetalleNomina.NumNomina" & vbLf
    sql = sql & "WHERE(Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102))AND(dbo.DetalleNomina.SalarioBasico <> 0)"
    sql = sql & "ORDER BY dbo.Empleado.CodEmpleado1, DetalleNomina.NumNomina"
     
     ArepInss2.AdoNomina.Source = sql
     Exportar = True
     ArepInss2.Show 1
     Exportar = False
  
  Case "Salario Basico Vrs Produccion"
   
    Dim Codigo1 As Integer, Codigo2 As Integer
    Exportar = True
    Fecha1 = Year(Me.TxtFecha1.Value) & "-" & Month(Me.TxtFecha1.Value) & "-" & Day(Me.TxtFecha1.Value)
    Fecha2 = Year(Me.TxtFecha2.Value) & "-" & Month(Me.TxtFecha2.Value) & "-" & Day(Me.TxtFecha2.Value)

     ArepBasicoProduccion.DataControl1.ConnectionString = ConexionReporte
     ArepBasicoProduccion.lblTitulo.Caption = Titulo
     ArepBasicoProduccion.LblSubtitulo.Caption = SubTitulo
     ArepBasicoProduccion.ImgLogo.Picture = LoadPicture(RutaLogo)
    If Me.DataCombo1.Text = "" Or Me.DataCombo2.Text = "" Then
     sql = "SELECT Empleado.CodEmpleado1, Empleado.CodEmpleado, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres,DetalleNomina.NumNomina, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HE, DetalleNomina.HorasExtras,DetalleNomina.OtrosIngresos , DetalleNomina.Comisiones, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, Nomina.FechaNomina, Nomina.Periodo,Nomina.FechaNominaINI FROM Empleado INNER JOIN" & vbLf
     sql = sql & "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN" & vbLf
     sql = sql & "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina" & vbLf
     sql = sql & "WHERE(Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102))"
     sql = sql & "ORDER BY Empleado.CodEmpleado,DetalleNomina.NumNomina"
    Else
      Codigo1 = Me.DataCombo1.Columns(0).Text
      Codigo2 = Me.DataCombo2.Columns(0).Text
      sql = "SELECT Empleado.CodEmpleado1, Empleado.CodEmpleado, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres,DetalleNomina.NumNomina, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HE, DetalleNomina.HorasExtras,DetalleNomina.OtrosIngresos , DetalleNomina.Comisiones, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, Nomina.FechaNomina, Nomina.Periodo,Nomina.FechaNominaINI FROM Empleado INNER JOIN" & vbLf
     sql = sql & "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN" & vbLf
     sql = sql & "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina" & vbLf
     sql = sql & "WHERE(Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) and  (Empleado.CodEmpleado BETWEEN " & Codigo1 & " AND " & Codigo2 & ")"
     sql = sql & "ORDER BY Empleado.CodEmpleado,DetalleNomina.NumNomina"
    End If
'     AND (dbo.DetalleNomina.SalarioBasico <> 0)
     ArepBasicoProduccion.DataControl1.Source = sql
     ArepBasicoProduccion.Show 1

  
  
  End Select
  Exportar = False
End Sub

Private Sub CmdRptOtros_Click()
j = Shell("c:\RamReport\generador de reportes.exe", vbNormalFocus)
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdVerreporte_Click()
Dim Mes As Integer, Anno As Integer, Mese As String, Ano As String
Dim Numero As Integer, sql As String, Fecha1 As String, Fecha2 As String, SalarioPieza As Double
Dim Mes1 As Integer, Mes2 As Integer, rs As New ADODB.Recordset, unidad As Double, Precio As Double, RsReportes As New ADODB.Recordset
Dim rpt As Object, CodigoEmpleado As Double, FechaIni As Date, FechaFin As Date
Dim fPreview As New FrmPreview, FechaInicioVaca As Date, FechaFinVaca As Date

NumFecha1 = Me.DTFecha1.Value
NumFecha2 = Me.DTFecha2.Value
Mes = Month(Me.Mes.Value)
Anno = Year(Me.Mes.Value)
Ano = Str(Anno)

Espacio = " "
Select Case CmbReportes.Text

Case "Reporte Carnet Empleados"
    Set rpt = New ArepCarnet
    
    If Me.TDBDepartamentoIni.Text = "" And Me.TDBDepartamentoFin.Text = "" Then
       SqlString = "SELECT  Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Cargo.Cargo, departamento.departamento , Empleado.Activo FROM  Empleado INNER JOIN Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento  Where (Empleado.Activo = 1)"
    Else
       SqlString = "SELECT  Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Cargo.Cargo, departamento.departamento , Empleado.Activo, Departamento.CodDepartamento FROM  Empleado INNER JOIN Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento  Where (Empleado.Activo = 1) AND (Departamento.CodDepartamento BETWEEN '" & Me.TDBDepartamentoIni.Text & "' AND '" & Me.TDBDepartamentoFin.Text & "')"
    End If
    
    rpt.DataControl1.ConnectionString = Conexion
    rpt.DataControl1.Source = SqlString
    fPreview.RunReport rpt
    fPreview.Show 1



Case "Listado de Empleados FHM"
    If Me.DBTipoNominas.Text = "" Then
        MsgBox ("Tenes que seleccionar el tipo de nomina")
        Exit Sub
    End If

    Set rpt = New ACListEmpleados
    SqlString = "SELECT     Empleado.CodEmpleado1 AS Codigo, Empleado.NumCedula, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombre,    Empleado.SueldoPeriodo * 2 AS SueldoMensual, Cargo.Cargo   FROM         Empleado INNER JOIN   Cargo ON Empleado.CodCargo = Cargo.CodCargo  WHERE     (Empleado.CodTipoNomina = N'" & DBTipoNominas.Columns(0) & "') and Empleado.Activo = 'True' ORDER BY Empleado.Nombre1"
    rpt.Label5.Caption = "Listado de Empleados" & " " & DBTipoNominas.Text & ""
    
    rpt.DataControl1.ConnectionString = Conexion
    rpt.DataControl1.Source = SqlString
    fPreview.RunReport rpt
    fPreview.Show 1

Case "Reporte Estimado Vacaciones"
  Dim SqlEmpleados As String

        For i = 1 To 6
             FechaInicioVaca = DateSerial(Year(Me.TxtFecha2.Value), Month(Me.TxtFecha2.Value) - i, 1)
            If i = 5 Then
                 FechaFinVaca = DateSerial(Year(Me.TxtFecha2.Value), Month(Me.TxtFecha2.Value) - i, 0)
            End If
        Next
               
'     SqlEmpleados = "SELECT  *, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres FROM Historico INNER JOIN Empleado ON Historico.Codempleado = Empleado.CodEmpleado WHERE (Empleado.Activo = 1) AND (Empleado.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') AND (Historico.FechaContratoVac BETWEEN CONVERT(DATETIME, '" & Format(FechaInicioVaca, "yyyy-mm-dd") & "',102) AND CONVERT(DATETIME, '" & Format(FechaFinVaca, "yyyy-mm-dd") & "', 102)) ORDER BY Historico.Codempleado"
     SqlEmpleados = "SELECT Historico.Id, Historico.Codempleado, Historico.FechaNacimiento, Historico.FechaContrato, Historico.FechaContratoVac, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres FROM Historico INNER JOIN Empleado ON Historico.Codempleado = Empleado.CodEmpleado WHERE     (Historico.FechaContratoVac BETWEEN CONVERT(DATETIME, '" & Format(FechaInicioVaca, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFinVaca, "yyyy-mm-dd") & "', 102)) ORDER BY Historico.Codempleado"
     MDIPrimero.DtaConsulta.RecordSource = SqlEmpleados
     MDIPrimero.DtaConsulta.Refresh


      

    Set rpt = New ArepEstimadoVacaciones
     rpt.DataControl1.ConnectionString = Conexion
     rpt.DataControl1.Source = SqlEmpleados
     fPreview.RunReport rpt
     fPreview.Show 1


Case "Reporte Horas Extra"
Set rpt = New ArepHorasExtras2
Fecha1 = Year(Me.TxtFecha1.Value) & "-" & Month(Me.TxtFecha1.Value) & "-" & Day(Me.TxtFecha1.Value)
Fecha2 = Year(Me.TxtFecha2.Value) & "-" & Month(Me.TxtFecha2.Value) & "-" & Day(Me.TxtFecha2.Value)
       rpt.DataControl1.ConnectionString = ConexionReporte
       sql = "SELECT  Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres,SUM(DetalleNomina.HE) AS HE, SUM(DetalleNomina.HorasExtras) AS HorasExtras FROM Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina WHERE (Nomina.FechaNominaINI BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Empleado.Nombre1, Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 Having (SUM(DetalleNomina.HE) <> 0) ORDER BY Empleado.Nombre1"
'       ArepHorasExtras2.DataControl1.Source = "SELECT  Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres,SUM(DetalleNomina.HE) AS HE, SUM(DetalleNomina.HorasExtras) AS HorasExtras FROM Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina WHERE (Nomina.FechaNominaINI BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) GROUP BY Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 Having (SUM(DetalleNomina.HE) <> 0) ORDER BY Empleado.CodEmpleado1"
       rpt.lblTitulo.Caption = Titulo
       rpt.LblSubtitulo.Caption = SubTitulo
       rpt.ImgLogo.Picture = LoadPicture(RutaLogo)
       rpt.LblDesde.Caption = "Desde " & Me.TxtFecha1.Value & " Hasta " & Me.TxtFecha2.Value
'       ArepHorasExtras2.Show 1
     
     rpt.DataControl1.ConnectionString = Conexion
     rpt.DataControl1.Source = sql
     fPreview.RunReport rpt
     fPreview.Show 1
     
Case "Reporte Registro Vacaciones"

Dim fs As Boolean
fs = False
'//////////////////////////////// Saco datos generales del empleado ///////////////////////
MDIPrimero.DtaConsulta.RecordSource = "SELECT     TOP (1) Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombre, Historico.FechaContratoVac, Empleado.CodEmpleado, DATEADD(month,    (YEAR(Historico.FechaContratoVac) - 1900) * 12 + MONTH(Historico.FechaContratoVac), - 1) AS UdMes  FROM         Empleado INNER JOIN   Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (Empleado.CodEmpleado1 = '" & DataCombo1.Text & "') and Empleado.Activo = 'True'"
MDIPrimero.DtaConsulta.Refresh

'///////// Inicializo parametros generales ////////////
Dim VacacionesAcumuladas, VacacionesSolicitadas, SaldoActual, tempVacacionSolicitada, tempVacacionAcumulada As Double
Dim NombreCompleto As String
NombreCompleto = MDIPrimero.DtaConsulta.Recordset("Nombre")
Dim TotalAcumuladas As Double, TotalSolicitadas As Double
Dim CodEmpleado1A As String
CodEmpleado1A = MDIPrimero.DtaConsulta.Recordset("CodEmpleado")

Dim Inicio As Date
Dim tempInicio As Date
Dim Fin As Date, tempFin As Date

Inicio = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
Fin = TxtFecha2.Value


'/////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////




tempInicio = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")

tempFin = MDIPrimero.DtaConsulta.Recordset("udMes")



'/////////////////////////////// Inicializo el reporte //////////////////////////////////
 rs.CursorLocation = adUseClient  '-------------RECORSET DESCONECTADOS -------------------------------
         rs.Open "DELETE FROM Reportes", Conexion

         
         MDIPrimero.AdoReportes.RecordSource = "SELECT * From Reportes"
         MDIPrimero.AdoReportes.Refresh
         SaldoActual = 0
         tempVacacionSolicitada = 0
         tempVacacionAcumulada = 0
         
         
         Dim Tipo As String
         MDIPrimero.DtaControles.Refresh
         Tipo = MDIPrimero.DtaControles.Recordset("DiasMes")
         
    Do While (Inicio < Fin)
     
    
                                                        
             MDIPrimero.AdoReportes.Recordset.AddNew
             MDIPrimero.AdoReportes.Recordset("Campo1") = "" & Format(tempInicio, "dd/MM/yyyy") & " - " & Format(tempFin, "dd/MM/yyyy") & ""
             MDIPrimero.AdoReportes.Recordset("Campo2") = Format(tempInicio, "MMMM")
          
 
          If Tipo = 30 Then
                If CDbl(Format(tempFin, "d")) > 30 Then
                tempFin = tempFin - 1
                End If
          End If
          

            
            Dim DiasMes As Double
            
            DiasMes = DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) + 1
            
            If Format(tempInicio, "MMMM") = "febrero" Or Format(tempInicio, "MMMM") = "Febrero" Or Format(tempInicio, "MMMM") = "FEBRERO" Then
               If Tipo = 30 Then
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 3) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 2) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
                
               'if tipo = 31
               Else
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 4) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 3) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
               End If
               
                
            Else
                If Tipo = 30 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < 30 Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                Else
                    If (DateDiff("d", tempInicio, tempFin) + 1) <= 30 Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
                    Else
                        tempVacacionAcumulada = 2.58
                    End If
                End If
            End If
          
             
             MDIPrimero.AdoReportes.Recordset("Num1") = Format(tempVacacionAcumulada, "##,##0.00")
             
             '////////     /////////     ////////        /////////       /////////       ////////        //////
             '////////// Calculo el total de dias y horas solicitadas en el rango de fechas recorrido /////////
             '/////////////////////////////////////////////////////////////////////////////////////////////////
           If DiasMes = 31 Then
                AdoAuxiliar.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM         SolicitudVacaciones WHERE not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'  and  (CodigoEmpleado = '" & Me.DataCombo1.Text & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(DateAdd("d", 1, tempFin), "dd/MM/yyyy") & " 23:59')"
           Else
                AdoAuxiliar.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM         SolicitudVacaciones WHERE not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado' and  (CodigoEmpleado = '" & Me.DataCombo1.Text & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(tempFin, "dd/MM/yyyy") & " 23:59')"
           End If
           
             AdoAuxiliar.Refresh
             If Not AdoAuxiliar.Recordset.EOF Then
                If Not IsNull(AdoAuxiliar.Recordset("VacacionesSolicitadas")) Then
                    tempVacacionSolicitada = AdoAuxiliar.Recordset("VacacionesSolicitadas")
                    MDIPrimero.AdoReportes.Recordset("Num2") = Format(tempVacacionSolicitada, "##,##0.00")
                Else
                    MDIPrimero.AdoReportes.Recordset("Num2") = 0
                    tempVacacionSolicitada = 0
                End If
             Else
                 tempVacacionSolicitada = 0
                 MDIPrimero.AdoReportes.Recordset("Num2") = 0
             End If
             
             SaldoActual = SaldoActual + (tempVacacionAcumulada - tempVacacionSolicitada)
             MDIPrimero.AdoReportes.Recordset("Num3") = Format(SaldoActual, "##,##0.00")
             MDIPrimero.AdoReportes.Recordset("Fecha1") = tempInicio
             If DiasMes = 31 Then
              MDIPrimero.AdoReportes.Recordset("Fecha2") = DateAdd("d", 1, tempFin)
             Else
               MDIPrimero.AdoReportes.Recordset("Fecha2") = tempFin
             End If
             
            TotalAcumuladas = TotalAcumuladas + tempVacacionAcumulada
             TotalSolicitadas = TotalSolicitadas + tempVacacionSolicitada
             
             MDIPrimero.AdoReportes.Recordset.Update
             

             If DateAdd("d", 2, tempFin) >= Fin Then
                 tempFin = DateAdd("m", -2, tempFin)
                 tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1)  'Inicio
                 tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin   '
                 'ponerle temp inicio para que los dias  no varien
                 Inicio = Fin
             Else
                tempFin = DateAdd("d", 2, tempFin)
                tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1)  'Inicio
                tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin   '
                'ponerle temp inicio para que los dias  no varien
                Inicio = tempInicio
             End If
             

            
           ' ////////////////
            
            
             If DateSerial(Year(tempFin), Month(tempFin) + 1, 0) >= Fin Then
                 tempFin = Fin
              
    
                                                        
             MDIPrimero.AdoReportes.Recordset.AddNew
             MDIPrimero.AdoReportes.Recordset("Campo1") = "" & Format(tempInicio, "dd/MM/yyyy") & " - " & Format(tempFin, "dd/MM/yyyy") & ""
             MDIPrimero.AdoReportes.Recordset("Campo2") = Format(tempInicio, "MMMM")

          If Tipo = 30 Then
                If CDbl(Format(tempFin, "d")) > 30 Then
                tempFin = tempFin - 1
                End If
          End If
            
            
            DiasMes = DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) + 1
            
            If Format(tempInicio, "MMMM") = "febrero" Or Format(tempInicio, "MMMM") = "Febrero" Or Format(tempInicio, "MMMM") = "FEBRERO" Then
                 
               If Tipo = 30 Then
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 3) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 2) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
                
               'if tipo = 31
               Else
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 4) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 3) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
               End If
            Else
                If Tipo = 30 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < 30 Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                Else
                    If (DateDiff("d", tempInicio, tempFin) + 1) <= 30 Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
                    Else
                        tempVacacionAcumulada = 2.58
                    End If
                End If
            End If
          
             
             MDIPrimero.AdoReportes.Recordset("Num1") = Format(tempVacacionAcumulada, "##,##0.00")
             
             '////////     /////////     ////////        /////////       /////////       ////////        //////
             '////////// Calculo el total de dias y horas solicitadas en el rango de fechas recorrido /////////
             '/////////////////////////////////////////////////////////////////////////////////////////////////
           If DiasMes = 31 Then
                AdoAuxiliar.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM SolicitudVacaciones WHERE  not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'  and    (CodigoEmpleado = '" & Me.DataCombo1.Text & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(DateAdd("d", 1, tempFin), "dd/MM/yyyy") & " 23:59')"
           Else
                AdoAuxiliar.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM SolicitudVacaciones WHERE  not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'  and    (CodigoEmpleado = '" & Me.DataCombo1.Text & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(tempFin, "dd/MM/yyyy") & " 23:59')"
           End If
           
             AdoAuxiliar.Refresh
             If Not AdoAuxiliar.Recordset.EOF Then
                If Not IsNull(AdoAuxiliar.Recordset("VacacionesSolicitadas")) Then
                    tempVacacionSolicitada = AdoAuxiliar.Recordset("VacacionesSolicitadas")
                    MDIPrimero.AdoReportes.Recordset("Num2") = Format(tempVacacionSolicitada, "##,##0.00")
                Else
                    MDIPrimero.AdoReportes.Recordset("Num2") = 0
                    tempVacacionSolicitada = 0
                End If
             Else
                 tempVacacionSolicitada = 0
                 MDIPrimero.AdoReportes.Recordset("Num2") = 0
             End If
             
             SaldoActual = SaldoActual + (tempVacacionAcumulada - tempVacacionSolicitada)
             MDIPrimero.AdoReportes.Recordset("Num3") = Format(SaldoActual, "##,##0.00")
             MDIPrimero.AdoReportes.Recordset("Fecha1") = tempInicio
             If DiasMes = 31 Then
               MDIPrimero.AdoReportes.Recordset("Fecha2") = DateAdd("d", 1, tempFin)
             Else
               MDIPrimero.AdoReportes.Recordset("Fecha2") = tempFin
             End If
             
             TotalAcumuladas = TotalAcumuladas + tempVacacionAcumulada
             TotalSolicitadas = TotalSolicitadas + tempVacacionSolicitada
             
             MDIPrimero.AdoReportes.Recordset.Update
             
             tempFin = DateAdd("d", 2, tempFin)
             
             tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1) ' Inicio
             tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin
                         'ponerle temp inicio para que los dias  no varien
            Inicio = tempInicio
                    Inicio = DateAdd("m", 1, Inicio)
            End If
            
            
            
    Loop
    
     Set rpt = New ArepRegistroVacaciones
     rpt.TxtNombre.Text = NombreCompleto
     rpt.txtCodigo.Text = DataCombo1.Text
     rpt.txtTotalAcumuladas.Text = Format(TotalAcumuladas, "##,##0.00")
     rpt.txtTotalSolicitadas.Text = Format(TotalSolicitadas, "##,##0.00")
     rpt.adoRegistroVacaciones.ConnectionString = Conexion
     rpt.adoRegistroVacaciones.Source = "Select * from Reportes"
     fPreview.RunReport rpt
     fPreview.Show 1
     
     

     
     
Case "Reporte Total Vacaciones"




       '/////////////////////////////// Inicializo el reporte //////////////////////////////////
         rs.CursorLocation = adUseClient  '-------------RECORSET DESCONECTADOS -------------------------------
         rs.Open "DELETE FROM Reportes", Conexion

         
         MDIPrimero.AdoReportes.RecordSource = "SELECT * From Reportes"
         MDIPrimero.AdoReportes.Refresh

If Me.DBTipoNominas.Text = "" Then
MDIPrimero.DtaConsulta.RecordSource = "SELECT      CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombre, Historico.FechaContratoVac, DATEADD(month,    (YEAR(Historico.FechaContratoVac) - 1900) * 12 + MONTH(Historico.FechaContratoVac), - 1) AS UdMes  FROM         Empleado INNER JOIN   Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE  (Empleado.Activo = 1) "

Else
MDIPrimero.DtaConsulta.RecordSource = "SELECT     CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombre, Historico.FechaContratoVac, DATEADD(month,    (YEAR(Historico.FechaContratoVac) - 1900) * 12 + MONTH(Historico.FechaContratoVac), - 1) AS UdMes  FROM         Empleado INNER JOIN   Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE  (Empleado.Activo = 1) and    (Empleado.CodTipoNomina = N'" & DBTipoNominas.Columns(0) & "') order by nombre"
End If


         'Dim Tipo As String
         MDIPrimero.DtaControles.Refresh
         Tipo = MDIPrimero.DtaControles.Recordset("DiasMes")

'//////////////////////////////// Saco datos generales de los empleados ///////////////////////

MDIPrimero.DtaConsulta.Refresh
MDIPrimero.DtaConsulta.Recordset.MoveFirst

    
    Barra.Max = MDIPrimero.DtaConsulta.Recordset.RecordCount
    DoEvents
    
    Do While Not MDIPrimero.DtaConsulta.Recordset.EOF
    
    
    
     Dim TotalVacacionesDisfrutadas, TotalDiasVacaciones, TotalDiasDisponibles As Double
    TotalVacacionesDisfrutadas = 0
    TotalDiasVacaciones = 0
    TotalDiasDisponibles = 0

    Barra.Value = Barra.Value + 1
    fs = False
'//////////////////////////////// Saco datos generales del empleado ///////////////////////

'///////// Inicializo parametros generales ////////////

VacacionesAcumuladas = 0
VacacionesSolicitadas = 0
SaldoActual = 0
tempVacacionSolicitada = 0
tempVacacionAcumulada = 0
NombreCompleto = MDIPrimero.DtaConsulta.Recordset("Nombre")
FrmReportes.Caption = NombreCompleto
Dim CodEmpleado1 As String
CodEmpleado1 = MDIPrimero.DtaConsulta.Recordset("CodEmpleado1")



Inicio = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
tempInicio = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
Fin = Me.TxtFecha2.Value
tempFin = MDIPrimero.DtaConsulta.Recordset("udMes")


         SaldoActual = 0

    Do While (Inicio < Fin)
             tempVacacionSolicitada = 0
             tempVacacionAcumulada = 0
         
           If Tipo = 30 Then
                If CDbl(Format(tempFin, "d")) > 30 Then
                tempFin = tempFin - 1
                End If
           End If
            
            
            
            DiasMes = DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) + 1
            
            If Format(tempInicio, "MMMM") = "febrero" Or Format(tempInicio, "MMMM") = "Febrero" Or Format(tempInicio, "MMMM") = "FEBRERO" Then
                
               If Tipo = 30 Then
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 3) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 2) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
                
               'if tipo = 31
               Else
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 4) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 3) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
               End If
            Else
            If Tipo = 30 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < 30 Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                Else
                    If (DateDiff("d", tempInicio, tempFin) + 1) <= 30 Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
                    Else
                        tempVacacionAcumulada = 2.58
                    End If
                End If
            End If
            
             TotalDiasVacaciones = TotalDiasVacaciones + tempVacacionAcumulada
             
             '////////     /////////     ////////        /////////       /////////       ////////        //////
             '////////// Calculo el total de dias y horas solicitadas en el rango de fechas recorrido /////////
             '/////////////////////////////////////////////////////////////////////////////////////////////////
           
           If DiasMes = 31 Then
                AdoAuxiliar.RecordSource = "select SUM(DiasDisfrutar)  AS VacacionesSolicitadas  FROM         SolicitudVacaciones WHERE not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'   and    (CodigoEmpleado = '" & CodEmpleado1 & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(DateAdd("d", 1, tempFin), "dd/MM/yyyy") & " 23:59')"
           Else
                AdoAuxiliar.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM         SolicitudVacaciones WHERE  not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'  and    (CodigoEmpleado = '" & CodEmpleado1 & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(tempFin, "dd/MM/yyyy") & " 23:59')"
           End If
           
             AdoAuxiliar.Refresh
             If Not AdoAuxiliar.Recordset.EOF Then
                If Not IsNull(AdoAuxiliar.Recordset("VacacionesSolicitadas")) Then
                    tempVacacionSolicitada = AdoAuxiliar.Recordset("VacacionesSolicitadas")
                Else
                    tempVacacionSolicitada = 0
                End If
             Else
                 tempVacacionSolicitada = 0
             End If
             
             TotalVacacionesDisfrutadas = TotalVacacionesDisfrutadas + tempVacacionSolicitada
             
             SaldoActual = SaldoActual + (tempVacacionAcumulada - tempVacacionSolicitada)
             
            
             
             If DateAdd("d", 2, tempFin) >= Fin Then
                 tempFin = DateAdd("m", -2, tempFin)
                 tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1)  'Inicio
                 tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin   '
                 'ponerle temp inicio para que los dias  no varien
                 Inicio = Fin
             Else
                tempFin = DateAdd("d", 2, tempFin)
                tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1)  'Inicio
                tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin   '
                'ponerle temp inicio para que los dias  no varien
                Inicio = tempInicio
             End If
            
           ' ////////////////
            
            
             If DateSerial(Year(tempFin), Month(tempFin) + 1, 0) >= Fin Then
                 tempFin = Fin
                tempVacacionSolicitada = 0
                tempVacacionAcumulada = 0
          
          If Tipo = 30 Then
                If CDbl(Format(tempFin, "d")) > 30 Then
                tempFin = tempFin - 1
                End If
          End If
            
            
            
            DiasMes = DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) + 1
            
            If Format(tempInicio, "MMMM") = "febrero" Or Format(tempInicio, "MMMM") = "Febrero" Or Format(tempInicio, "MMMM") = "FEBRERO" Then
                
               If Tipo = 30 Then
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 3) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 2) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
                
               'if tipo = 31
               Else
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 4) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 3) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
               End If
            Else
            If Tipo = 30 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < 30 Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                Else
                    If (DateDiff("d", tempInicio, tempFin) + 1) <= 30 Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
                    Else
                        tempVacacionAcumulada = 2.58
                    End If
                End If
            End If
          
             
           TotalDiasVacaciones = TotalDiasVacaciones + tempVacacionAcumulada
             
             '////////     /////////     ////////        /////////       /////////       ////////        //////
             '////////// Calculo el total de dias y horas solicitadas en el rango de fechas recorrido /////////
             '/////////////////////////////////////////////////////////////////////////////////////////////////
           If DiasMes = 31 Then
                AdoAuxiliar.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM SolicitudVacaciones WHERE  not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'   and     (CodigoEmpleado = '" & CodEmpleado1 & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(DateAdd("d", 1, tempFin), "dd/MM/yyyy") & " 23:59')"
           Else
                AdoAuxiliar.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM SolicitudVacaciones WHERE  not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'   and     (CodigoEmpleado = '" & CodEmpleado1 & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(tempFin, "dd/MM/yyyy") & " 23:59')"
           End If
           
             AdoAuxiliar.Refresh
             If Not AdoAuxiliar.Recordset.EOF Then
                If Not IsNull(AdoAuxiliar.Recordset("VacacionesSolicitadas")) Then
                    tempVacacionSolicitada = AdoAuxiliar.Recordset("VacacionesSolicitadas")
                Else
                    tempVacacionSolicitada = 0
                End If
             Else
                 tempVacacionSolicitada = 0
             End If
             
             TotalVacacionesDisfrutadas = TotalVacacionesDisfrutadas + tempVacacionSolicitada
             
             TotalDiasDisponibles = TotalDiasVacaciones - TotalVacacionesDisfrutadas
             
             SaldoActual = SaldoActual + (tempVacacionAcumulada - tempVacacionSolicitada)

             tempFin = DateAdd("d", 2, tempFin)
             
             tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1) ' Inicio
             tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin
                         'ponerle temp inicio para que los dias  no varien
            Inicio = tempInicio
            
             
                    Inicio = DateAdd("m", 1, Inicio)
            End If
            
            
            
    Loop
    

    
     MDIPrimero.AdoReportes.Recordset.AddNew
     MDIPrimero.AdoReportes.Recordset("Campo1") = CodEmpleado1
     MDIPrimero.AdoReportes.Recordset("Campo2") = NombreCompleto
     MDIPrimero.AdoReportes.Recordset("Num1") = TotalDiasVacaciones
     MDIPrimero.AdoReportes.Recordset("Num2") = TotalVacacionesDisfrutadas
     MDIPrimero.AdoReportes.Recordset("Num3") = TotalDiasDisponibles
     MDIPrimero.AdoReportes.Recordset.Update
                
     MDIPrimero.DtaConsulta.Recordset.MoveNext
     
     
    Loop
    
    
    
     Set rpt = New ArepTotalVacaciones
     
     rpt.LblDesde.Caption = "Nomina: " & Me.DBTipoNominas.Text
     rpt.AdoTotalVacaciones.ConnectionString = Conexion
     rpt.AdoTotalVacaciones.Source = "Select * from Reportes"
     fPreview.RunReport rpt
     fPreview.Show 1
    
    FrmReportes.Caption = "Reportes de Nomina"
    Barra.Value = 0
    


      
     
Case "Reporte Consolidado Vacaciones"









         '/////////////////////////////// Inicializo el reporte //////////////////////////////////
         rs.CursorLocation = adUseClient  '-------------RECORSET DESCONECTADOS -------------------------------
         rs.Open "DELETE FROM Reportes", Conexion

         
         MDIPrimero.AdoReportes.RecordSource = "SELECT * From Reportes"
         MDIPrimero.AdoReportes.Refresh

If Me.DBTipoNominas.Text = "" Then
 MDIPrimero.DtaConsulta.RecordSource = "SELECT     Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombre, Historico.FechaContratoVac, DATEADD(month, (YEAR(Historico.FechaContratoVac)   - 1900) * 12 + MONTH(Historico.FechaContratoVac), - 1) AS UdMes, Empleado.CodEmpleado1    FROM         Empleado INNER JOIN   Historico ON Empleado.CodEmpleado = Historico.Codempleado Where  (Empleado.Activo = 1) order by nombre"

Else
 MDIPrimero.DtaConsulta.RecordSource = "SELECT     Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombre, Historico.FechaContratoVac, DATEADD(month, (YEAR(Historico.FechaContratoVac)   - 1900) * 12 + MONTH(Historico.FechaContratoVac), - 1) AS UdMes, Empleado.CodEmpleado1   FROM         Empleado INNER JOIN    Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE  (Empleado.Activo = 1) and    (Empleado.CodTipoNomina = N'" & DBTipoNominas.Columns(0) & "') order by nombre"
End If

'//////////////////////////////// Saco datos generales de los empleados ///////////////////////

         'Dim Tipo As String
         MDIPrimero.DtaControles.Refresh
         Tipo = MDIPrimero.DtaControles.Recordset("DiasMes")

MDIPrimero.DtaConsulta.Refresh
MDIPrimero.DtaConsulta.Recordset.MoveFirst

    Barra.Max = MDIPrimero.DtaConsulta.Recordset.RecordCount
DoEvents
    Do While Not MDIPrimero.DtaConsulta.Recordset.EOF
      Dim CodigoEmpleadoS As String
        
        CodigoEmpleadoS = MDIPrimero.DtaConsulta.Recordset("CodEmpleado1")
        If CodigoEmpleadoS = "0007" Then
                CodigoEmpleadoS = "0007"
        End If
        
        NombreCompleto = MDIPrimero.DtaConsulta.Recordset("Nombre")
        Inicio = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
        tempInicio = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
        Fin = TxtFecha2.Value
        tempFin = MDIPrimero.DtaConsulta.Recordset("udMes")
        SaldoActual = 0
        
          FrmReportes.Caption = CodigoEmpleadoS & " " & NombreCompleto
          Barra.Value = Barra.Value + 1
          DoEvents
     
        Do While (Inicio < Fin)
            tempVacacionSolicitada = 0
            tempVacacionAcumulada = 0
            Dim Date1 As Date
            Date1 = tempInicio
            Dim Date2 As Date
            Date2 = tempFin
            
            
          If Tipo = 30 Then
                If CDbl(Format(tempFin, "d")) > 30 Then
                tempFin = tempFin - 1
                End If
          End If
            
            'Dim DiasMes As Double
            
            DiasMes = DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) + 1
            
            If Format(tempInicio, "MMMM") = "febrero" Or Format(tempInicio, "MMMM") = "Febrero" Or Format(tempInicio, "MMMM") = "FEBRERO" Then
                
               If Tipo = 30 Then
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 3) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 2) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
                
               'if tipo = 31
               Else
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 4) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 3) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
               End If
            Else
                If Tipo = 30 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < 30 Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                Else
                    If (DateDiff("d", tempInicio, tempFin) + 1) <= 30 Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
                    Else
                        tempVacacionAcumulada = 2.58
                    End If
                End If
            End If
            
          'tempVacacionAcumulada = CalcularDiasVaca(Date1, Date2)
          'tempVacacionAcumulada = tempVacacionAcumulada / 12
          
             'If (DateDiff("d", tempInicio, tempFin)) <=  Then
                 'tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
             'Else
                 'tempVacacionAcumulada = 2.5
             'End If
                    
       
             
             '////////     /////////     ////////        /////////       /////////       ////////        //////
             '////////// Calculo el total de dias y horas solicitadas en el rango de fechas recorrido /////////
             '////////    /////////     /////////       //////////      /////////       ////////////   ////////
           If DiasMes = 31 Then
                AdoAuxiliar.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM         SolicitudVacaciones WHERE   not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'  and    (CodigoEmpleado = '" & CodigoEmpleadoS & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(DateAdd("d", 1, tempFin), "dd/MM/yyyy") & " 23:59')"
           Else
                AdoAuxiliar.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM         SolicitudVacaciones WHERE   not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'  and   (CodigoEmpleado = '" & CodigoEmpleadoS & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(tempFin, "dd/MM/yyyy") & " 23:59')"
           End If
             AdoAuxiliar.Refresh
             If Not AdoAuxiliar.Recordset.EOF Then
                If Not IsNull(AdoAuxiliar.Recordset("VacacionesSolicitadas")) Then
                    tempVacacionSolicitada = AdoAuxiliar.Recordset("VacacionesSolicitadas")
                Else
                    tempVacacionSolicitada = 0
                End If
             Else
                 tempVacacionSolicitada = 0
             End If
             
             SaldoActual = SaldoActual + (tempVacacionAcumulada - tempVacacionSolicitada)
             
          If DateAdd("d", 2, tempFin) >= Fin Then
                 tempFin = DateAdd("m", -2, tempFin)
                 tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1)  'Inicio
                 tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin   '
                 'ponerle temp inicio para que los dias  no varien
                 Inicio = Fin
             Else
                tempFin = DateAdd("d", 2, tempFin)
                tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1)  'Inicio
                tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin   '
                'ponerle temp inicio para que los dias  no varien
                Inicio = tempInicio
             End If
            
             If DateSerial(Year(tempFin), Month(tempFin) + 1, 0) >= Fin Then
                 tempFin = Fin
                
                 tempVacacionSolicitada = 0
                 tempVacacionAcumulada = 0
                 Date1 = tempInicio
                 Date2 = tempFin
          
                   ' If (DateDiff("d", tempInicio, tempFin) + 1) <= 28 Then
                        'tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin)) / 12
                    'Else
                        'tempVacacionAcumulada = 2.5
                   ' End If
                   
                    'If (DateDiff("d", tempInicio, tempFin)) <= DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) Then
                       ' tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
                    'Else
                     '   tempVacacionAcumulada = 2.5
                    'End If
                    
            ' tempVacacionAcumulada = CalcularDiasVaca(Date1, Date2)
             'tempVacacionAcumulada = tempVacacionAcumulada / 12
             
          If Tipo = 30 Then
                If CDbl(Format(tempFin, "d")) > 30 Then
                tempFin = tempFin - 1
                End If
          End If
            
          
            
            DiasMes = DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) + 1
            
            If Format(tempInicio, "MMMM") = "febrero" Or Format(tempInicio, "MMMM") = "Febrero" Or Format(tempInicio, "MMMM") = "FEBRERO" Then
                 
               If Tipo = 30 Then
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 3) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 2) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
                
               'if tipo = 31
               Else
                If DiasMes = 28 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 4) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                ElseIf DiasMes = 29 Then
                      If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 3) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                End If
               End If
            Else
            If Tipo = 30 Then
                    If (DateDiff("d", tempInicio, tempFin) + 1) < 30 Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
                    Else
                        tempVacacionAcumulada = 2.5
                    End If
                Else
                    If (DateDiff("d", tempInicio, tempFin) + 1) <= 30 Then
                        tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
                    Else
                        tempVacacionAcumulada = 2.58
                    End If
                End If
            End If
             
             
                    '////////     /////////     ////////        /////////       /////////       ////////        //////
                    '////////// Calculo el total de dias y horas solicitadas en el rango de fechas recorrido /////////
                    '/////////////////////////////////////////////////////////////////////////////////////////////////
                    If DiasMes = 31 Then
                         AdoAuxiliar.RecordSource = "select SUM(DiasDisfrutar)  AS VacacionesSolicitadas  FROM         SolicitudVacaciones WHERE  not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'   and    (CodigoEmpleado = '" & CodigoEmpleadoS & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(DateAdd("d", 1, tempFin), "dd/MM/yyyy") & " 23:59')"
                    Else
                         AdoAuxiliar.RecordSource = "select SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM         SolicitudVacaciones WHERE   not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'   and   (CodigoEmpleado = '" & CodigoEmpleadoS & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(tempFin, "dd/MM/yyyy") & " 23:59')"
                    End If
                   AdoAuxiliar.Refresh
                    If Not AdoAuxiliar.Recordset.EOF Then
                        If Not IsNull(AdoAuxiliar.Recordset("VacacionesSolicitadas")) Then
                            tempVacacionSolicitada = AdoAuxiliar.Recordset("VacacionesSolicitadas")
                        Else
                            tempVacacionSolicitada = 0
                        End If
                    Else
                        tempVacacionSolicitada = 0
                    End If
             
                    SaldoActual = SaldoActual + (tempVacacionAcumulada - tempVacacionSolicitada)
                    Inicio = DateAdd("m", 1, Inicio)
               
            End If
            
               
        Loop
        
        If Me.Check1.Value = 1 Then
        
            If txtCantidad.Text = "" Or Not IsNumeric(txtCantidad.Text) Then
            txtCantidad.Text = 0
            End If
        If CodigoEmpleadoS = "108" Then
            CodigoEmpleadoS = "108"
        End If
        
            If SaldoActual > CDbl(Me.txtCantidad.Text) Then
                MDIPrimero.AdoReportes.Recordset.AddNew
                MDIPrimero.AdoReportes.Recordset("Campo1") = CodigoEmpleadoS
                MDIPrimero.AdoReportes.Recordset("Campo2") = NombreCompleto
                MDIPrimero.AdoReportes.Recordset("Fecha1") = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
                MDIPrimero.AdoReportes.Recordset("Num1") = SaldoActual
                MDIPrimero.AdoReportes.Recordset.Update
            End If
        Else
                MDIPrimero.AdoReportes.Recordset.AddNew
                MDIPrimero.AdoReportes.Recordset("Campo1") = CodigoEmpleadoS
                MDIPrimero.AdoReportes.Recordset("Campo2") = NombreCompleto
                MDIPrimero.AdoReportes.Recordset("Fecha1") = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
                MDIPrimero.AdoReportes.Recordset("Num1") = SaldoActual
                MDIPrimero.AdoReportes.Recordset.Update
        End If
        
        
        
        
        
     
     MDIPrimero.DtaConsulta.Recordset.MoveNext
    Loop



'///////// Inicializo parametros generales ////////////


    
     Set rpt = New ArepConsolidadoVacaciones
     rpt.Field3.Text = "Saldo Acumulado a: " & Format(Me.TxtFecha2.Value, "dd MMMM, yyyy")
     rpt.LblDesde.Caption = "Nomina: " & Me.DBTipoNominas.Text
     rpt.AdoConsolidadoVacaciones.ConnectionString = Conexion
     rpt.AdoConsolidadoVacaciones.Source = "Select * from Reportes"
     fPreview.RunReport rpt
     fPreview.Show 1
    
    FrmReportes.Caption = "Reportes de Nomina"
    Barra.Value = 0


    
    Case "Reporte x Provision"
''  Me.AdoBusca.RecordSource = "Select FechaNominaIni, FechaNomina from Nomina where CodTipoNomina = '" & Me.DBTipoNominas.Columns(0) & "' and activa = 'True'"
'' Me.AdoBusca.Refresh
 
     'Numero = val(Me.TxtNNomina.Text)
    Set rpt = New ArepProvicion

     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.lblTitulo.Caption = Titulo
     rpt.LblSubtitulo.Caption = SubTitulo
     rpt.LblTitulo3.Caption = "Reporte de Provisiones, Nomina " & Me.DBTipoNominas.Text & ", desde " & Format(Me.TxtFecha1.Value, "dd/MM/yyyy") & " hasta " & Format(Me.TxtFecha2.Value, "dd/MM/yyyy") & ""
     rpt.ImgLogo.Picture = LoadPicture(RutaLogo)


 

     rpt.DataControl1.Source = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Historico.FechaContrato, Departamento.Departamento, DetalleNomina.NumNomina, Nomina.FechaNominaINI, Nomina.FechaNomina, DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.SeptimoDia AS TotalDevengado, (DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.SeptimoDia) / 6.96 AS TotalDevengadoDiario, Empleado.CodEmpleado FROM  Empleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado " & _
                                         "INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina    WHERE     (Nomina.FechaNominaINI >= CONVERT(DATETIME, '" & Format(TxtFecha1.Value, "yyyy-MM-dd") & "', 102)) AND (Nomina.FechaNomina <= CONVERT(DATETIME,  '" & Format(Me.TxtFecha2.Value, "yyyy-MM-dd") & "', 102)) AND  (Empleado.CodTipoNomina =  '" & Me.DBTipoNominas.Columns(0) & "')"
                      'Where   (Nomina.FechaNominaINI >= '" & Format(TxtFecha1.Value, "yyyyMMdd") & "') AND (Nomina.FechaNomina <= '" & Format(Me.TxtFecha2.Value, "yyyyMdd") & "') and Empleado.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0) & "' ORDER BY Empleado.CodEmpleado1"


     fPreview.RunReport rpt
     fPreview.Show 1
    
    
    
Case "Reporte GRAL INGRESOS"

Fecha1 = Year(Me.TxtFecha1.Value) & "-" & Month(Me.TxtFecha1.Value) & "-" & Day(Me.TxtFecha1.Value)
Fecha2 = Year(Me.TxtFecha2.Value) & "-" & Month(Me.TxtFecha2.Value) & "-" & Day(Me.TxtFecha2.Value)

'      ArepMensualIR.DataControl1.ConnectionString = ConexionReporte
'     ArepMensualIR.LblTitulo.Caption = Titulo
'     ArepMensualIR.LblSubtitulo.Caption = "REPORTE IR EMPLEADOS"
'     ArepMensualIR.ImgLogo.Picture = LoadPicture(RutaLogo)
'     ArepMensualIR.LblFecha.Caption = Format(Now, "Long Date")

     
     
  If Me.Check1.Value = 0 Then
     sql = "SELECT     TOP 100 PERCENT MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, " & _
                      "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos " & _
                       "+ DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + dbo.DetalleNomina.Antiguedad) AS TotalDevengado, " & _
                       "Empleado.CodEmpleado1, MAX(DetalleNomina.NumNomina) AS NumNomina, SUM(DetalleNomina.MontoIR) AS MontoIR, " & _
                      "Nomina.CodTipoNomina AS CodTipoNomina,  MAX(Empleado.CodEmpleado) AS CodEmpleado, SUM(DetalleNomina.MontoINSS) AS MontoINSS " & _
            "FROM         Nomina INNER JOIN " & _
                      "Grupo INNER JOIN " & _
                      "Cargo INNER JOIN " & _
                      "TipoNomina INNER JOIN " & _
                      "Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN " & _
                      "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON " & _
                      "Nomina.NumNomina = DetalleNomina.NumNomina " & _
            "WHERE     (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
            "GROUP BY Empleado.CodEmpleado1, Nomina.CodTipoNomina " & _
            "HAVING      (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "')" & _
            "ORDER BY Empleado.CodEmpleado1, MAX(DetalleNomina.NumNomina) "
            
         rs.CursorLocation = adUseClient  '-------------RECORSET DESCONECTADOS -------------------------------
         rs.Open "DELETE FROM Reportes", Conexion

         
         MDIPrimero.AdoReportes.RecordSource = "SELECT * From Reportes"
         MDIPrimero.AdoReportes.Refresh
         
         
         MDIPrimero.DtaConsulta.RecordSource = sql
         MDIPrimero.DtaConsulta.Refresh
         MDIPrimero.DtaConsulta.Recordset.MoveLast
         
         
         RsReportes.Open sql, Conexion
         RsReportes.MoveFirst
         
         
         Me.Barra.Min = 0
         Me.Barra.Max = MDIPrimero.DtaConsulta.Recordset.RecordCount
         Me.Barra.Value = 0
         
         
         Do While Not RsReportes.EOF

           DoEvents
           CodigoEmpleado = RsReportes("CodEmpleado")
            FechaIni = MesIni(FrmReportes.Combo1.Text, FrmReportes.DBCAño.Text)
            FechaFin = MesIni(FrmReportes.Combo2.Text, FrmReportes.DBAño2.Text)
            FechaFin = DateSerial(Year(FechaFin), Month(FechaFin) + 1, 0)
           
            MDIPrimero.DtaConsulta.RecordSource = "SELECT SUM(DetalleNomVaca.Inss) AS Inss, SUM(DetalleNomVaca.Ir) AS Ir, SUM(DetalleNomVaca.TotalDevengado) AS TotalDevengado, DetalleNomVaca.CodEmpleado FROM DetalleNomVaca INNER JOIN  NomVaca ON DetalleNomVaca.NumNomVaca = NomVaca.NumNomVaca WHERE (NomVaca.FechaAplica BETWEEN CONVERT(DATETIME, '" & Format(FechaIni, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFin, "yyyy-mm-dd") & "', 102)) GROUP BY DetalleNomVaca.CodEmpleado Having (DetalleNomVaca.CodEmpleado = " & CodigoEmpleado & ")"
            MDIPrimero.DtaConsulta.Refresh

          If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
            MDIPrimero.AdoReportes.Recordset.AddNew
             MDIPrimero.AdoReportes.Recordset("Campo1") = RsReportes("Nombres")
             MDIPrimero.AdoReportes.Recordset("Num1") = RsReportes("TotalDevengado")
             MDIPrimero.AdoReportes.Recordset("Campo2") = RsReportes("CodEmpleado1")
             MDIPrimero.AdoReportes.Recordset("Num2") = RsReportes("NumNomina")
             MDIPrimero.AdoReportes.Recordset("Num3") = RsReportes("MontoIR")
             MDIPrimero.AdoReportes.Recordset("Campo3") = RsReportes("CodTipoNomina")
             MDIPrimero.AdoReportes.Recordset("Num4") = RsReportes("CodEmpleado")
             MDIPrimero.AdoReportes.Recordset("Num5") = RsReportes("MontoINSS")
             MDIPrimero.AdoReportes.Recordset("Num6") = MDIPrimero.DtaConsulta.Recordset("Ir")
            MDIPrimero.AdoReportes.Recordset.Update
           
          Else
            MDIPrimero.AdoReportes.Recordset.AddNew
             MDIPrimero.AdoReportes.Recordset("Campo1") = RsReportes("Nombres")
             MDIPrimero.AdoReportes.Recordset("Num1") = RsReportes("TotalDevengado")
             MDIPrimero.AdoReportes.Recordset("Campo2") = RsReportes("CodEmpleado1")
             MDIPrimero.AdoReportes.Recordset("Num2") = RsReportes("NumNomina")
             MDIPrimero.AdoReportes.Recordset("Num3") = RsReportes("MontoIR")
             MDIPrimero.AdoReportes.Recordset("Campo3") = RsReportes("CodTipoNomina")
             MDIPrimero.AdoReportes.Recordset("Num4") = RsReportes("CodEmpleado")
             MDIPrimero.AdoReportes.Recordset("Num5") = RsReportes("MontoINSS")
             MDIPrimero.AdoReportes.Recordset("Num6") = 0
          
          End If
         
         Me.Barra.Value = Me.Barra.Value + 1
         RsReportes.MoveNext
         Loop
            
       sql = "SELECT Campo1 AS Nombres, Num1 AS TotalDevengado, Campo2 AS CodEmpleado1, Num2 AS NumNomina, Num3 AS MontoIR, Campo3 AS CodTipoNomina, Num4 AS CodEmpleado, Num5 AS MontoInss, Num6 AS MontoIrVaca, Num3 + Num6 AS MontoTotalIR From Reportes Where (Num3 + Num6 <> 0) ORDER BY Campo2"
            
   Else
     sql = "SELECT     TOP 100 PERCENT MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, " & _
                      "SUM(DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos " & _
                       "+ DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + dbo.DetalleNomina.Antiguedad) AS TotalDevengado, " & _
                       "Empleado.CodEmpleado1, MAX(DetalleNomina.NumNomina) AS NumNomina, SUM(DetalleNomina.MontoIR) AS MontoIR, " & _
                      "Nomina.CodTipoNomina AS CodTipoNomina,  MAX(Empleado.CodEmpleado) AS CodEmpleado, SUM(DetalleNomina.MontoINSS) AS MontoINSS " & _
            "FROM         Nomina INNER JOIN " & _
                      "Grupo INNER JOIN " & _
                      "Cargo INNER JOIN " & _
                      "TipoNomina INNER JOIN " & _
                      "Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN " & _
                      "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON " & _
                      "Nomina.NumNomina = DetalleNomina.NumNomina " & _
            "WHERE     (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
            "GROUP BY Empleado.CodEmpleado1, Nomina.CodTipoNomina " & _
            "HAVING      (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "')" & _
            "ORDER BY Empleado.CodEmpleado1, MAX(DetalleNomina.NumNomina) "
  End If
     
'   ArepIRMensualDetallado.DataControl1.ConnectionString = Conexion
'   ArepIRMensualDetallado.DataControl1.Source = sql
'   ArepIRMensualDetallado.Show 1

   


     Set rpt = New ArepIRMensualDetallado
     rpt.DataControl1.ConnectionString = Conexion
     rpt.DataControl1.Source = sql
     fPreview.RunReport rpt
     fPreview.Show 1
     
Case "Reporte Proyeccion Vacaciones"


     Set rpt = New ArepProyectionVaca
     rpt.DataControl1.ConnectionString = Conexion
     rpt.DataControl1.Source = "SELECT  Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Historico.FechaContratoVac , departamento.departamento FROM  Historico INNER JOIN  Empleado ON Historico.Codempleado = Empleado.CodEmpleado INNER JOIN  Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento  " & _
                               "WHERE (Historico.FechaContratoVac BETWEEN CONVERT(DATETIME, '" & Format(Me.TxtFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.TxtFecha2.Value, "yyyy-mm-dd") & "', 102)) ORDER BY Departamento.Departamento, Historico.FechaContratoVac"
     fPreview.RunReport rpt

     fPreview.Show 1



    
    
Case "Reporte Dias Acumulados"
    Fecha1 = Year(Me.TxtFecha1.Value) & "-" & Month(Me.TxtFecha1.Value) & "-" & Day(Me.TxtFecha1.Value)
    Fecha2 = Year(Me.TxtFecha2.Value) & "-" & Month(Me.TxtFecha2.Value) & "-" & Day(Me.TxtFecha2.Value)


    ArepDevengado.DataControl1.ConnectionString = ConexionReporte
    ArepDevengado.lblTitulo.Caption = Titulo
    ArepDevengado.LblSubtitulo.Caption = "REPORTE DETALLADO DEDUCCIONES SEGUN NOMINA"
    ArepDevengado.LblFecha.Caption = "Impreso desde: " & Me.TxtFecha1.Value & " Hasta: " & Me.TxtFecha2.Value
    ArepDevengado.LblFechaHoy.Caption = Format(Now, "Long Date")
    
'    sql = "SELECT Empleado.CodEmpleado1 AS CodEmpleado1, Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2 AS Nombres, Nomina.FechaNomina, Empleado.Apellido1, Empleado.Apellido2,  SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos  + DetalleNomina.Incentivos + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) AS TotalIngresos,  MAX(DetalleNomina.NumNomina) AS NumNomina, Fecha_Planilla.mes AS Mes, Fecha_Planilla.año AS Año, SUM(DetalleNomina.MontoINSS) AS MontoInss, SUM(DetalleNomina.MontoIR) AS MontoIR, Empleado.NumCedula, Historico.FechaContrato, Empleado.CodEmpleado " & _
'          "FROM  Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN Fecha_Planilla ON DetalleNomina.NumNomina = Fecha_Planilla.NumNomina INNER JOIN  Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado  WHERE     (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) AND  (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') GROUP BY Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2, Fecha_Planilla.año, Empleado.CodEmpleado1, Fecha_Planilla.mes, Empleado.NumCedula, Empleado.Apellido1, Empleado.Apellido2, Historico.FechaContrato, Empleado.CodEmpleado , Nomina.FechaNomina " & _
'          "HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) <> 0) ORDER BY MAX(DetalleNomina.NumNomina), Empleado.CodEmpleado1"
     sql = "SELECT  Empleado.CodEmpleado1, Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2 AS Nombres, Empleado.Apellido1, Empleado.Apellido2 , Empleado.NumCedula, Empleado.CodEmpleado, Empleado.Activo, Empleado.NumeroInss, Historico.FechaContrato FROM  Empleado INNER JOIN  Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE  (Empleado.Activo = 1) AND (Empleado.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') ORDER BY Empleado.CodEmpleado1"
    ArepDevengado.ImgLogo.Picture = LoadPicture(RutaLogo)
    ArepDevengado.DataControl1.Source = sql
'    ArepDevengado.Show 1

     Set rpt = New ArepDevengado
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.DataControl1.Source = sql
     fPreview.RunReport rpt


     fPreview.Show 1

Case "Reporte INSS E IR MENSUAL"

    Fecha1 = Year(Me.TxtFecha1.Value) & "-" & Month(Me.TxtFecha1.Value) & "-" & Day(Me.TxtFecha1.Value)
    Fecha2 = Year(Me.TxtFecha2.Value) & "-" & Month(Me.TxtFecha2.Value) & "-" & Day(Me.TxtFecha2.Value)


    ArepInssIr.DataControl1.ConnectionString = ConexionReporte
    ArepInssIr.lblTitulo.Caption = Titulo
    ArepInssIr.LblSubtitulo.Caption = "REPORTE DETALLADO DEDUCCIONES SEGUN NOMINA"
    ArepInssIr.LblFecha.Caption = "Impreso desde: " & Me.TxtFecha1.Value & " Hasta: " & Me.TxtFecha2.Value
    ArepInssIr.LblFechaHoy.Caption = Format(Now, "Long Date")
    
          
'   SQl = "SELECT Empleado.CodEmpleado1 AS CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) AS TotalIngresos,MAX(DetalleNomina.NumNomina) AS NumNomina, Fecha_Planilla.mes AS Mes, Fecha_Planilla.año AS Año, SUM(DetalleNomina.MontoINSS)AS MontoInss, SUM(DetalleNomina.MontoIR) AS MontoIR, Empleado.NumCedula FROM  Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN Fecha_Planilla ON DetalleNomina.NumNomina = Fecha_Planilla.NumNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina  " & _
'         "WHERE (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) AND (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') GROUP BY Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2, Fecha_Planilla.año, Empleado.CodEmpleado1, Fecha_Planilla.Mes , Empleado.NumCedula HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos                        + DetalleNomina.Incentivos + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) <> 0) ORDER BY MAX(DetalleNomina.NumNomina), Empleado.CodEmpleado1"
        
    sql = "SELECT Empleado.CodEmpleado1 AS CodEmpleado1, Empleado.Nombre1 + N' ' + Empleado.Nombre2 AS Nombres, Empleado.Apellido1, Empleado.Apellido2, SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion + DetalleNomina.Antiguedad) AS TotalIngresos, MAX(DetalleNomina.NumNomina) AS NumNomina, Fecha_Planilla.mes AS Mes, Fecha_Planilla.año AS Año, SUM(DetalleNomina.MontoINSS) AS MontoInss, SUM(DetalleNomina.MontoIR) AS MontoIR, Empleado.NumCedula, Historico.FechaContrato FROM Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN  Fecha_Planilla ON DetalleNomina.NumNomina = Fecha_Planilla.NumNomina INNER JOIN  Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado " & _
          "WHERE (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) AND (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') GROUP BY Empleado.Nombre1 + N' ' + Empleado.Nombre2, Fecha_Planilla.año, Empleado.CodEmpleado1, Fecha_Planilla.mes, Empleado.NumCedula, Empleado.Apellido1 , Empleado.Apellido2, Historico.FechaContrato HAVING (SUM(DetalleNomina.SalarioBasico + DetalleNomina.Destajo + DetalleNomina.HorasExtras + DetalleNomina.Comisiones + DetalleNomina.OtrosIngresos + DetalleNomina.Incentivos + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion) <> 0) ORDER BY MAX(DetalleNomina.NumNomina), Empleado.CodEmpleado1"
        
    ArepInssIr.ImgLogo.Picture = LoadPicture(RutaLogo)
    ArepInssIr.DataControl1.Source = sql
    ArepInssIr.Show 1

Case "Reporte Detalle Deducciones"
'Dim rpt As Object
Set rpt = New ArepDetalleDeduccion
Fecha1 = Format(Me.TxtFecha1.Value, "yyyy/mm/dd")
Fecha2 = Format(Me.TxtFecha2.Value, "yyyy/mm/dd")





    'si empleado
If DataCombo1.Text <> "" Then
        
        
        If TDBCombo2.Text <> "" Then
            sql = "SELECT  TipoDeduccion.Deduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumNomina, Empleado.CodEmpleado1, " & _
                         "Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano " & _
                          "FROM  DetalleDeduccion INNER JOIN  Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion INNER JOIN  TipoDeduccion ON Deduccion.CodTipoDeduccion = TipoDeduccion.CodTipoDeduccion INNER JOIN  Empleado ON Deduccion.CodEmpleado = Empleado.CodEmpleado INNER JOIN Nomina ON DetalleDeduccion.NumNomina = Nomina.NumNomina " & _
                          "WHERE     (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) And (DetalleDeduccion.Valor <> 0 )  AND (Deduccion.CodTipoDeduccion = N'" & Me.TDBCombo2.Text & "') AND (Empleado.CodEmpleado1 = '" & Me.DataCombo1.Text & "') " & _
                          "ORDER BY TipoDeduccion.Deduccion, Empleado.Nombre1 "
        Else
               sql = "SELECT  TipoDeduccion.Deduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumNomina, Empleado.CodEmpleado1, " & _
                         "Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano " & _
                          "FROM  DetalleDeduccion INNER JOIN  Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion INNER JOIN  TipoDeduccion ON Deduccion.CodTipoDeduccion = TipoDeduccion.CodTipoDeduccion INNER JOIN  Empleado ON Deduccion.CodEmpleado = Empleado.CodEmpleado INNER JOIN Nomina ON DetalleDeduccion.NumNomina = Nomina.NumNomina " & _
                          "WHERE     (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) And (DetalleDeduccion.Valor <> 0 )  AND (Empleado.CodEmpleado1 = '" & Me.DataCombo1.Text & "') " & _
                          "ORDER BY TipoDeduccion.Deduccion, Empleado.Nombre1 "
        End If
        
    
    
Else
   
    If Me.DBTipoNominas.Text <> "" Then
    
        If TDBCombo2.Text <> "" Then
            sql = "SELECT  TipoDeduccion.Deduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumNomina, Empleado.CodEmpleado1, " & _
                         "Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano " & _
                          "FROM  DetalleDeduccion INNER JOIN  Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion INNER JOIN  TipoDeduccion ON Deduccion.CodTipoDeduccion = TipoDeduccion.CodTipoDeduccion INNER JOIN  Empleado ON Deduccion.CodEmpleado = Empleado.CodEmpleado INNER JOIN Nomina ON DetalleDeduccion.NumNomina = Nomina.NumNomina " & _
                          "WHERE     (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) And (DetalleDeduccion.Valor <> 0 )  AND (Deduccion.CodTipoDeduccion = N'" & Me.TDBCombo2.Text & "') AND (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "')  " & _
                          "ORDER BY TipoDeduccion.Deduccion, Empleado.Nombre1 "
        Else
            sql = "SELECT  TipoDeduccion.Deduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumNomina, Empleado.CodEmpleado1, " & _
                         "Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano " & _
                          "FROM  DetalleDeduccion INNER JOIN  Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion INNER JOIN  TipoDeduccion ON Deduccion.CodTipoDeduccion = TipoDeduccion.CodTipoDeduccion INNER JOIN  Empleado ON Deduccion.CodEmpleado = Empleado.CodEmpleado INNER JOIN Nomina ON DetalleDeduccion.NumNomina = Nomina.NumNomina " & _
                          "WHERE     (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) And (DetalleDeduccion.Valor <> 0 )  AND (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "')    " & _
                          "ORDER BY TipoDeduccion.Deduccion, Empleado.Nombre1 "
        End If
    
    Else
        
        If TDBCombo2.Text <> "" Then
            sql = "SELECT  TipoDeduccion.Deduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumNomina, Empleado.CodEmpleado1, " & _
                         "Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano " & _
                          "FROM  DetalleDeduccion INNER JOIN  Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion INNER JOIN  TipoDeduccion ON Deduccion.CodTipoDeduccion = TipoDeduccion.CodTipoDeduccion INNER JOIN  Empleado ON Deduccion.CodEmpleado = Empleado.CodEmpleado INNER JOIN Nomina ON DetalleDeduccion.NumNomina = Nomina.NumNomina " & _
                          "WHERE     (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) And (DetalleDeduccion.Valor <> 0 )  AND (Deduccion.CodTipoDeduccion = N'" & Me.TDBCombo2.Text & "') " & _
                          "ORDER BY TipoDeduccion.Deduccion, Empleado.Nombre1 "
        Else
            sql = "SELECT  TipoDeduccion.Deduccion, DetalleDeduccion.Valor, DetalleDeduccion.NumNomina, Empleado.CodEmpleado1, " & _
                         "Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Nomina.FechaNomina, Nomina.Mes, Nomina.Ano " & _
                          "FROM  DetalleDeduccion INNER JOIN  Deduccion ON DetalleDeduccion.NumDeduccion = Deduccion.NumDeduccion INNER JOIN  TipoDeduccion ON Deduccion.CodTipoDeduccion = TipoDeduccion.CodTipoDeduccion INNER JOIN  Empleado ON Deduccion.CodEmpleado = Empleado.CodEmpleado INNER JOIN Nomina ON DetalleDeduccion.NumNomina = Nomina.NumNomina " & _
                          "WHERE     (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) And (DetalleDeduccion.Valor <> 0 ) " & _
                          "ORDER BY TipoDeduccion.Deduccion, Empleado.Nombre1 "
        End If
        
    End If

End If





       
       
 rpt.DataControl1.ConnectionString = ConexionReporte
 rpt.lblTitulo.Caption = Titulo
 rpt.LblSubtitulo.Caption = "REPORTE DETALLADO DEDUCCIONES SEGUN NOMINA"
 rpt.LblDesde.Caption = "Impreso desde: " & Me.TxtFecha1.Value & " Hasta: " & Me.TxtFecha2.Value
 rpt.ImgLogo.Picture = LoadPicture(RutaLogo)
 rpt.DataControl1.Source = sql
' ArepDetalleDeduccion.Show 1
           fPreview.arv.ReportSource = rpt
           fPreview.Show 1


Case "Reporte Inss 2"
 Dim MesAnterior As String
Fecha1 = Year(Me.TxtFecha1.Value) & "-" & Month(Me.TxtFecha1.Value) & "-" & Day(Me.TxtFecha1.Value)
Fecha2 = Year(Me.TxtFecha2.Value) & "-" & Month(Me.TxtFecha2.Value) & "-" & Day(Me.TxtFecha2.Value)

'////////////////BUSCO LA FECHA DEL MES ANTERIOR///////////

Select Case Me.Combo1.Text
  Case "Enero"
     MesAnterior = "Diciembre"
  Case "Febrero"
     MesAnterior = "Enero"
  Case "Marzo"
     MesAnterior = "Febrero"
  Case "Abril"
     MesAnterior = "Marzo"
  Case "Mayo"
     MesAnterior = "Abril"
  Case "Junio"
     MesAnterior = "Mayo"
  Case "Julio"
      MesAnterior = "Junio"
  Case "Agosto"
     MesAnterior = "Julio"
  Case "Septiembre"
     MesAnterior = "Agosto"
  Case "Octubre"
     MesAnterior = "Septiembre"
  Case "Noviembre"
     MesAnterior = "Octubre"
  Case "Diciembre"
     MesAnterior = "Noviembre"

End Select

FMes (MesAnterior)
Mes1 = Format(Nmes, "0#")
FMes (MesAnterior)
Mes2 = Format(Nmes, "0#")
If MesAnterior = "Diciembre" Then
 Año1 = val(Me.DBCAño.Text) - 1
 Año2 = val(Me.DBCAño.Text) - 1
Else
 Año1 = val(Me.DBCAño.Text)
 Año2 = val(Me.DBCAño.Text)
End If
CodTipoNomina = Me.DBTipoNominas.Columns(0).Text


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Fecha1Reporte = Format(Me.AdoBusca.Recordset("Inicio"), "yyyy/mm/dd")
 End If
 
Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
  Fecha2Reporte = Format(Me.AdoBusca.Recordset("Final"), "yyyy/mm/dd")
 End If







Mes1 = Month(Fecha2)
ConvertirMes (Mes1)
ArepInss2.LblMes1.Caption = "Informe del mes de " & Convertir
ArepInss2.LblPeriodo.Caption = Convertir & " / " & Year(Me.DTFecha1.Value)

ArepInss2.LblMes2.Caption = "Informe de " & MesAnterior

    ArepInss2.AdoNomina.ConnectionString = ConexionReporte
    ArepInss2.lblTitulo.Caption = Titulo
    ArepInss2.LblSubtitulo.Caption = "REPORTE INSS EMPLEADOS"
    ArepInss2.ImgLogo.Picture = LoadPicture(RutaLogo)
 
sql = "SELECT     TOP 100 PERCENT dbo.Empleado.Nombre1 + N' ' + dbo.Empleado.Nombre2 + N' ' + dbo.Empleado.Apellido1 + N' ' + dbo.Empleado.Apellido2 AS Nombres," & vbLf
sql = sql & "                       dbo.DetalleNomina.CodEmpleado, dbo.DetalleNomina.MontoINSS, dbo.DetalleNomina.INSSPatronal," & vbLf
sql = sql & "                dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
sql = sql & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.IncetivoProduccion AS TotalDevengado," & vbLf
sql = sql & "                       dbo.DetalleNomina.INATEC, dbo.Empleado.NumeroInss, dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.INSSPatronal AS TotalInss," & vbLf
sql = sql & "                      dbo.Empleado.CodEmpleado1 , dbo.DetalleNomina.NumNomina, dbo.Nomina.FechaNomina, dbo.Cargo.Cargo" & vbLf
sql = sql & "FROM         dbo.Nomina INNER JOIN" & vbLf
sql = sql & "                      dbo.Grupo INNER JOIN" & vbLf
sql = sql & "                      dbo.Cargo INNER JOIN" & vbLf
sql = sql & "                      dbo.TipoNomina INNER JOIN" & vbLf
sql = sql & "                      dbo.Empleado ON dbo.TipoNomina.CodTipoNomina = dbo.Empleado.CodTipoNomina ON dbo.Cargo.CodCargo = dbo.Empleado.CodCargo INNER JOIN" & vbLf
sql = sql & "                      dbo.DetalleNomina ON dbo.Empleado.CodEmpleado = dbo.DetalleNomina.CodEmpleado ON dbo.Grupo.CodGrupo = dbo.Empleado.CodGrupo ON" & vbLf
sql = sql & "                      dbo.TipoNomina.CodTipoNomina = dbo.Nomina.CodTipoNomina And dbo.Nomina.NumNomina = dbo.DetalleNomina.NumNomina" & vbLf
sql = sql & "WHERE(Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) And Empleado.Activo = 'True'" & vbLf
sql = sql & "ORDER BY dbo.Empleado.CodEmpleado1, dbo.Nomina.FechaNomina"
     
     ArepInss2.AdoNomina.Source = sql
'     ArepInss2.Show 1
           fPreview.arv.ReportSource = ArepInss2
           fPreview.Show 1

Case "Numeros Disponibles"
  ArepNumerosDisponibles.DataControl1.ConnectionString = ConexionReporte
  ArepNumerosDisponibles.lblTitulo.Caption = Titulo
  ArepNumerosDisponibles.LblSubtitulo.Caption = "LISTADO DE NUMEROS DISPONIBLES"
  ArepNumerosDisponibles.ImgLogo.Picture = LoadPicture(RutaLogo)
  ArepNumerosDisponibles.LblTitulo3.Caption = Format(Now, "Long Date")

'  ArepNumerosDisponibles.Show 1
           fPreview.arv.ReportSource = ArepNumerosDisponibles
           fPreview.Show 1
 

Case "Listado Maestro de Empleados"

  Set rpt = New ArepListaMaestro
  rpt.DataControl1.ConnectionString = ConexionReporte
  rpt.lblTitulo.Caption = Titulo
  rpt.LblSubtitulo.Caption = "LISTADO MAESTRO DE EMPLEADOS"
  'ArepListaMaestro.ImgLogo.Picture = LoadPicture(RutaLogo)
  rpt.LblFecha.Caption = Format(Now, "Long Date")
  rpt.LblDesde.Caption = Format(Now, "Long Date")
  rpt.LblHasta.Caption = Format(Now, "Long Date")
  
  sql = "SELECT     TOP 100 PERCENT Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico," & vbLf
  sql = sql & "                    Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
  sql = sql & "                     Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
  sql = sql & "                    Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
  sql = sql & "                    DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
  sql = sql & "                    Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
  sql = sql & "                    DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
  sql = sql & "                    DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
  sql = sql & "                    DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
  sql = sql & "                    DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
  sql = sql & "                    Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
  sql = sql & "                    DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
  sql = sql & "                     DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia AS TotalDevengado," & vbLf
  sql = sql & "                    DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
  sql = sql & "                    (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
  sql = sql & "                     DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia)" & vbLf
  sql = sql & "                    - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
  sql = sql & "                    Empleado.TarifaHoraria, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, Empleado.Activo, Empleado.NumeroInss, Empleado.CodEmpleado1," & vbLf
  sql = sql & "                    Historico.FechaContrato , Empleado.Sexo" & vbLf
  sql = sql & "FROM         Nomina INNER JOIN" & vbLf
  sql = sql & "                    Grupo INNER JOIN" & vbLf
  sql = sql & "                    Cargo INNER JOIN" & vbLf
  sql = sql & "                    TipoNomina INNER JOIN" & vbLf
  sql = sql & "                    Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
  sql = sql & "                    DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
  sql = sql & "                    TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN" & vbLf
  sql = sql & "                    Historico ON Empleado.CodEmpleado = Historico.Codempleado" & vbLf
  sql = sql & "Where (Nomina.NumNomina = '" & Me.TxtNNomina.Text & "')" & vbLf
  sql = sql & "ORDER BY Nomina.NumNomina, Empleado.CodEmpleado1"
  
  rpt.DataControl1.Source = sql
'  ArepListaMaestro.Show 1
           fPreview.arv.ReportSource = rpt
           fPreview.Show 1


Case "Reporte Inss"
Set rpt = New ArepInssCotiza
Fecha1 = Year(Me.TxtFecha1.Value) & "-" & Month(Me.TxtFecha1.Value) & "-" & Day(Me.TxtFecha1.Value)
Fecha2 = Year(Me.TxtFecha2.Value) & "-" & Month(Me.TxtFecha2.Value) & "-" & Day(Me.TxtFecha2.Value)

    rpt.AdoNomina.ConnectionString = ConexionReporte
     rpt.lblTitulo.Caption = Titulo
     rpt.LblSubtitulo.Caption = "REPORTE INSS EMPLEADOS"
     rpt.ImgLogo.Picture = LoadPicture(RutaLogo)
     rpt.LblFecha.Caption = Format(Now, "Long Date")
     rpt.LblDesde.Caption = Me.DTFecha1.Value
     rpt.LblHasta.Caption = Me.DTFecha2.Value
sql = "SELECT     TOP 100 PERCENT dbo.Empleado.Nombre1 + N' ' + dbo.Empleado.Nombre2 + N' ' + dbo.Empleado.Apellido1 + N' ' + dbo.Empleado.Apellido2 AS Nombres," & vbLf
sql = sql & "                       dbo.DetalleNomina.CodEmpleado, dbo.DetalleNomina.MontoINSS, dbo.DetalleNomina.INSSPatronal," & vbLf
sql = sql & "                dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
sql = sql & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.IncetivoProduccion + dbo.DetalleNomina.Antiguedad AS TotalDevengado," & vbLf
sql = sql & "                       dbo.DetalleNomina.INATEC, dbo.Empleado.NumeroInss, dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.INSSPatronal AS TotalInss," & vbLf
sql = sql & "                      dbo.Empleado.CodEmpleado1 , dbo.DetalleNomina.NumNomina, dbo.Nomina.FechaNomina, dbo.Cargo.Cargo" & vbLf
sql = sql & "FROM         dbo.Nomina INNER JOIN" & vbLf
sql = sql & "                      dbo.Grupo INNER JOIN" & vbLf
sql = sql & "                      dbo.Cargo INNER JOIN" & vbLf
sql = sql & "                      dbo.TipoNomina INNER JOIN" & vbLf
sql = sql & "                      dbo.Empleado ON dbo.TipoNomina.CodTipoNomina = dbo.Empleado.CodTipoNomina ON dbo.Cargo.CodCargo = dbo.Empleado.CodCargo INNER JOIN" & vbLf
sql = sql & "                      dbo.DetalleNomina ON dbo.Empleado.CodEmpleado = dbo.DetalleNomina.CodEmpleado ON dbo.Grupo.CodGrupo = dbo.Empleado.CodGrupo ON" & vbLf
sql = sql & "                      dbo.TipoNomina.CodTipoNomina = dbo.Nomina.CodTipoNomina And dbo.Nomina.NumNomina = dbo.DetalleNomina.NumNomina" & vbLf
sql = sql & "WHERE(Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) and Empleado.Activo = 'False'"
sql = sql & "ORDER BY dbo.Empleado.CodEmpleado1, DetalleNomina.NumNomina"
'AND(dbo.DetalleNomina.SalarioBasico <> 0)
     rpt.AdoNomina.Source = sql
'     ArepInssCotiza.Show 1
           fPreview.arv.ReportSource = rpt
           fPreview.Show 1

     
Case "Reporte Detalle Ir"
Set rpt = New ArepIrDetalle
Fecha1 = Year(Me.TxtFecha1.Value) & "-" & Day(Me.TxtFecha1.Value) & "-" & Month(Me.TxtFecha1.Value)
Fecha2 = Year(Me.TxtFecha2.Value) & "-" & Day(Me.TxtFecha2.Value) & "-" & Month(Me.TxtFecha2.Value)

      rpt.AdoNomina.ConnectionString = ConexionReporte
     rpt.lblTitulo.Caption = Titulo
     rpt.LblSubtitulo.Caption = "REPORTE DETALLE Ir EMPLEADOS"
     rpt.ImgLogo.Picture = LoadPicture(RutaLogo)
     rpt.LblFecha.Caption = Format(Now, "Long Date")
     rpt.LblDesde.Caption = Me.TxtFecha1.Value
     rpt.LblHasta.Caption = Me.TxtFecha2.Value
sql = "SELECT     TOP 100 PERCENT dbo.Empleado.Nombre1 + N' ' + dbo.Empleado.Nombre2 + N' ' + dbo.Empleado.Apellido1 + N' ' + dbo.Empleado.Apellido2 AS Nombres," & vbLf
sql = sql & "                       dbo.DetalleNomina.CodEmpleado, dbo.DetalleNomina.MontoINSS, dbo.DetalleNomina.INSSPatronal," & vbLf
sql = sql & "                dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
sql = sql & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.IncetivoProduccion + dbo.DetalleNomina.Antiguedad AS TotalDevengado," & vbLf
sql = sql & "                       dbo.DetalleNomina.INATEC, dbo.Empleado.CodInss, dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.INSSPatronal AS TotalInss," & vbLf
sql = sql & "                      dbo.Empleado.CodEmpleado1 , dbo.DetalleNomina.NumNomina, dbo.Nomina.FechaNomina, dbo.Cargo.Cargo,DetalleNomina.MontoIr, Empleado.Codir" & vbLf
sql = sql & "FROM         dbo.Nomina INNER JOIN" & vbLf
sql = sql & "                      dbo.Grupo INNER JOIN" & vbLf
sql = sql & "                      dbo.Cargo INNER JOIN" & vbLf
sql = sql & "                      dbo.TipoNomina INNER JOIN" & vbLf
sql = sql & "                      dbo.Empleado ON dbo.TipoNomina.CodTipoNomina = dbo.Empleado.CodTipoNomina ON dbo.Cargo.CodCargo = dbo.Empleado.CodCargo INNER JOIN" & vbLf
sql = sql & "                      dbo.DetalleNomina ON dbo.Empleado.CodEmpleado = dbo.DetalleNomina.CodEmpleado ON dbo.Grupo.CodGrupo = dbo.Empleado.CodGrupo ON" & vbLf
sql = sql & "                      dbo.TipoNomina.CodTipoNomina = dbo.Nomina.CodTipoNomina And dbo.Nomina.NumNomina = dbo.DetalleNomina.NumNomina" & vbLf
sql = sql & "WHERE(Nomina.FechaNomina BETWEEN ('" & Fecha1 & "') AND ('" & Fecha2 & "'))AND(dbo.DetalleNomina.MontoIr <> 0)"
sql = sql & "ORDER BY dbo.Empleado.CodEmpleado1, DetalleNomina.NumNomina"
     
     rpt.AdoNomina.Source = sql
'     ArepIrDetalle.Show 1
           fPreview.arv.ReportSource = rpt
           fPreview.Show 1
     
Case "Reporte Ir"
Set rpt = New ArepIrDetalle

Fecha1 = Year(Me.TxtFecha1.Value) & "-" & Month(Me.TxtFecha1.Value) & "-" & Day(Me.TxtFecha1.Value)
Fecha2 = Year(Me.TxtFecha2.Value) & "-" & Month(Me.TxtFecha2.Value) & "-" & Day(Me.TxtFecha2.Value)

      rpt.AdoNomina.ConnectionString = ConexionReporte
     rpt.lblTitulo.Caption = Titulo
     rpt.LblSubtitulo.Caption = "REPORTE IR EMPLEADOS"
     rpt.ImgLogo.Picture = LoadPicture(RutaLogo)
     rpt.LblFecha.Caption = Format(Now, "Long Date")
     rpt.LblDesde.Caption = Me.DTFecha1.Value
     rpt.LblHasta.Caption = Me.DTFecha2.Value
sql = "SELECT     TOP 100 PERCENT dbo.Empleado.Nombre1 + N' ' + dbo.Empleado.Nombre2 + N' ' + dbo.Empleado.Apellido1 + N' ' + dbo.Empleado.Apellido2 AS Nombres," & vbLf
sql = sql & "                       dbo.DetalleNomina.CodEmpleado, dbo.DetalleNomina.MontoINSS, dbo.DetalleNomina.INSSPatronal," & vbLf
sql = sql & "                dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
sql = sql & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.IncetivoProduccion + dbo.DetalleNomina.Antiguedad AS TotalDevengado," & vbLf
sql = sql & "                       dbo.DetalleNomina.INATEC, dbo.Empleado.CodInss, dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.INSSPatronal AS TotalInss," & vbLf
sql = sql & "                      dbo.Empleado.CodEmpleado1 , dbo.DetalleNomina.NumNomina, dbo.Nomina.FechaNomina, dbo.Cargo.Cargo,DetalleNomina.MontoIr, Empleado.Codir" & vbLf
sql = sql & "FROM         dbo.Nomina INNER JOIN" & vbLf
sql = sql & "                      dbo.Grupo INNER JOIN" & vbLf
sql = sql & "                      dbo.Cargo INNER JOIN" & vbLf
sql = sql & "                      dbo.TipoNomina INNER JOIN" & vbLf
sql = sql & "                      dbo.Empleado ON dbo.TipoNomina.CodTipoNomina = dbo.Empleado.CodTipoNomina ON dbo.Cargo.CodCargo = dbo.Empleado.CodCargo INNER JOIN" & vbLf
sql = sql & "                      dbo.DetalleNomina ON dbo.Empleado.CodEmpleado = dbo.DetalleNomina.CodEmpleado ON dbo.Grupo.CodGrupo = dbo.Empleado.CodGrupo ON" & vbLf
sql = sql & "                      dbo.TipoNomina.CodTipoNomina = dbo.Nomina.CodTipoNomina And dbo.Nomina.NumNomina = dbo.DetalleNomina.NumNomina" & vbLf
sql = sql & "WHERE(Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102))AND(dbo.DetalleNomina.MontoIr <> 0)"
sql = sql & "ORDER BY dbo.Empleado.CodEmpleado1, DetalleNomina.NumNomina"
     
     rpt.AdoNomina.Source = sql
    ' rpt.Show 1
     
           fPreview.arv.ReportSource = rpt
           fPreview.Show 1

Case "Reporte IR MENSUAL"


Fecha1 = Year(Me.TxtFecha1.Value) & "-" & Month(Me.TxtFecha1.Value) & "-" & Day(Me.TxtFecha1.Value)
Fecha2 = Year(Me.TxtFecha2.Value) & "-" & Month(Me.TxtFecha2.Value) & "-" & Day(Me.TxtFecha2.Value)
Dim rpta As New ArepMensualIR
      rpta.DataControl1.ConnectionString = ConexionReporte
     rpta.lblTitulo.Caption = Titulo
     rpta.LblSubtitulo.Caption = "REPORTE IR EMPLEADOS"
'     If Dir(RutaLogo, vbDirectory) Then
       rpta.ImgLogo.Picture = LoadPicture(RutaLogo)
'     End If
     rpta.LblFecha.Caption = Format(Now, "Long Date")
'     ArepMensualIR.LblDesde.Caption = Me.DTFecha1.Value
'    ArepMensualIR.LblHasta.Caption = Me.DTFecha2.Value
     
    
  If Me.Check1.Value = 0 Then
   '+ DetalleNomina.Comisiones
   
     sql = "SELECT     TOP 100 PERCENT MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, " & _
                      "SUM(DetalleNomina.SalarioBasico  + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos " & _
                       "+ DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion+ DetalleNomina.BonoProduccion + dbo.DetalleNomina.Antiguedad) AS TotalDevengado, " & _
                       "Empleado.CodEmpleado1, MAX(DetalleNomina.NumNomina) AS NumNomina, SUM(DetalleNomina.MontoIR) AS MontoIR, " & _
                      "Nomina.CodTipoNomina AS CodTipoNomina " & _
            "FROM         Nomina INNER JOIN " & _
                      "Grupo INNER JOIN " & _
                      "Cargo INNER JOIN " & _
                      "TipoNomina INNER JOIN " & _
                      "Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN " & _
                      "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON " & _
                      "TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina " & _
            "WHERE     (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
            "GROUP BY Empleado.CodEmpleado1, Nomina.CodTipoNomina " & _
            "HAVING      (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') AND (SUM(DetalleNomina.MontoIR) <> 0) " & _
            "ORDER BY Empleado.CodEmpleado1, MAX(DetalleNomina.NumNomina) "
   Else
   
   '+ DetalleNomina.Comisiones
     sql = "SELECT     TOP 100 PERCENT MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, " & _
                      "SUM(DetalleNomina.SalarioBasico  + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos " & _
                       "+ DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion+ DetalleNomina.BonoProduccion + dbo.DetalleNomina.Antiguedad) AS TotalDevengado, " & _
                       "Empleado.CodEmpleado1, MAX(DetalleNomina.NumNomina) AS NumNomina, SUM(DetalleNomina.MontoIR) AS MontoIR, " & _
                      "Nomina.CodTipoNomina AS CodTipoNomina " & _
            "FROM         Nomina INNER JOIN " & _
                      "Grupo INNER JOIN " & _
                      "Cargo INNER JOIN " & _
                      "TipoNomina INNER JOIN " & _
                      "Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN " & _
                      "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON " & _
                      "TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina " & _
            "WHERE     (Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) " & _
            "GROUP BY Empleado.CodEmpleado1, Nomina.CodTipoNomina " & _
            "HAVING      (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "')" & _
            "ORDER BY Empleado.CodEmpleado1, MAX(DetalleNomina.NumNomina) "
  End If
     
     rpta.DataControl1.Source = sql
     'rpta.Show 1
          fPreview.arv.ReportSource = rpta
       fPreview.Show 1



Case "Reporte Detalle Inss"
Set rpt = New ArepInssCotizaDetalle
Fecha1 = Year(Me.TxtFecha1.Value) & "-" & Month(Me.TxtFecha1.Value) & "-" & Day(Me.TxtFecha1.Value)
Fecha2 = Year(Me.TxtFecha2.Value) & "-" & Month(Me.TxtFecha2.Value) & "-" & Day(Me.TxtFecha2.Value)

    rpt.AdoNomina.ConnectionString = ConexionReporte
     rpt.lblTitulo.Caption = Titulo
     rpt.LblSubtitulo.Caption = "REPORTE DETALLE INSS EMPLEADOS"
     rpt.ImgLogo.Picture = LoadPicture(RutaLogo)
     rpt.LblFecha.Caption = Format(Now, "Long Date")
     rpt.LblDesde.Caption = Me.DTFecha1.Value
     rpt.LblHasta.Caption = Me.DTFecha2.Value
sql = "SELECT     TOP 100 PERCENT dbo.Empleado.Nombre1 + N' ' + dbo.Empleado.Nombre2 + N' ' + dbo.Empleado.Apellido1 + N' ' + dbo.Empleado.Apellido2 AS Nombres," & vbLf
sql = sql & "                       dbo.DetalleNomina.CodEmpleado, dbo.DetalleNomina.MontoINSS, dbo.DetalleNomina.INSSPatronal," & vbLf
sql = sql & "                dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Comisiones + dbo.DetalleNomina.Incentivos + dbo.DetalleNomina.HorasExtras + dbo.DetalleNomina.OtrosIngresos" & vbLf
sql = sql & "                       + dbo.DetalleNomina.Destajo + dbo.DetalleNomina.VacacionesPagadas + dbo.DetalleNomina.SeptimoDia + dbo.DetalleNomina.IncetivoProduccion + dbo.DetalleNomina.Antiguedad AS TotalDevengado," & vbLf
sql = sql & "                       dbo.DetalleNomina.INATEC, dbo.Empleado.CodInss, dbo.DetalleNomina.MontoINSS + dbo.DetalleNomina.INSSPatronal AS TotalInss," & vbLf
sql = sql & "                      dbo.Empleado.CodEmpleado1 , dbo.DetalleNomina.NumNomina, dbo.Nomina.FechaNomina, dbo.Cargo.Cargo" & vbLf
sql = sql & "FROM         dbo.Nomina INNER JOIN" & vbLf
sql = sql & "                      dbo.Grupo INNER JOIN" & vbLf
sql = sql & "                      dbo.Cargo INNER JOIN" & vbLf
sql = sql & "                      dbo.TipoNomina INNER JOIN" & vbLf
sql = sql & "                      dbo.Empleado ON dbo.TipoNomina.CodTipoNomina = dbo.Empleado.CodTipoNomina ON dbo.Cargo.CodCargo = dbo.Empleado.CodCargo INNER JOIN" & vbLf
sql = sql & "                      dbo.DetalleNomina ON dbo.Empleado.CodEmpleado = dbo.DetalleNomina.CodEmpleado ON dbo.Grupo.CodGrupo = dbo.Empleado.CodGrupo ON" & vbLf
sql = sql & "                      dbo.TipoNomina.CodTipoNomina = dbo.Nomina.CodTipoNomina And dbo.Nomina.NumNomina = dbo.DetalleNomina.NumNomina" & vbLf
sql = sql & "WHERE(Nomina.FechaNomina BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102))" & vbLf
sql = sql & "ORDER BY dbo.Empleado.CodEmpleado1, dbo.Nomina.FechaNomina"
     
     rpt.AdoNomina.Source = sql
     'ArepInssCotizaDetalle.Show 1
           fPreview.arv.ReportSource = rpt
       fPreview.Show 1


Case "Detalle Deducciones"
Set rpt = New ArepDetalleDeduccion
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.lblTitulo.Caption = Titulo
     rpt.LblSubtitulo.Caption = SubTitulo
     rpt.ImgLogo.Picture = LoadPicture(RutaLogo)

     If DataCombo1.Text = "" Or Me.DataCombo2.Text = "" Then
        sql = " SELECT     Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, " & _
            "TipoDeduccion.Deduccion, DetalleDeduccion.Valor, Nomina.NumNomina, Nomina.CodTipoNomina, Nomina.FechaNominaINI, Nomina.FechaNomina, " & _
            "Nomina.Mes , Nomina.Ano, Nomina.Periodo, TipoNomina.Nomina " & _
            "FROM         Deduccion INNER JOIN " & _
            "DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion INNER JOIN " & _
            "TipoDeduccion ON Deduccion.CodTipoDeduccion = TipoDeduccion.CodTipoDeduccion INNER JOIN " & _
            "Empleado ON Deduccion.CodEmpleado = Empleado.CodEmpleado INNER JOIN " & _
            "Nomina ON Deduccion.NUmNomina = Nomina.NumNomina INNER JOIN " & _
            "TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina AND Nomina.CodTipoNomina = TipoNomina.CodTipoNomina " & _
            "WHERE     (DetalleDeduccion.Valor <> 0) AND (Nomina.FechaNominaINI BETWEEN '" & Format(Me.TxtFecha1, "yyyymmdd") & "' AND '" & Format(Me.TxtFecha2, "yyyymmdd") & "') AND (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "')" & _
            "ORDER BY Empleado.CodEmpleado1, Nomina.NumNomina "
      Else
        sql = " SELECT     Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, " & _
            "TipoDeduccion.Deduccion, DetalleDeduccion.Valor, Nomina.NumNomina, Nomina.CodTipoNomina, Nomina.FechaNominaINI, Nomina.FechaNomina, " & _
            "Nomina.Mes , Nomina.Ano, Nomina.Periodo, TipoNomina.Nomina " & _
            "FROM         Deduccion INNER JOIN " & _
            "DetalleDeduccion ON Deduccion.NumDeduccion = DetalleDeduccion.NumDeduccion INNER JOIN " & _
            "TipoDeduccion ON Deduccion.CodTipoDeduccion = TipoDeduccion.CodTipoDeduccion INNER JOIN " & _
            "Empleado ON Deduccion.CodEmpleado = Empleado.CodEmpleado INNER JOIN " & _
            "Nomina ON Deduccion.NUmNomina = Nomina.NumNomina INNER JOIN " & _
            "TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina AND Nomina.CodTipoNomina = TipoNomina.CodTipoNomina " & _
            "WHERE     (DetalleDeduccion.Valor <> 0) AND (Nomina.FechaNominaINI BETWEEN '" & Format(Me.TxtFecha1, "dd/MM/yyyy") & "' AND '" & Format(Me.TxtFecha2, "dd/MM/yyyy") & "') AND (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "')" & _
            "AND (Empleado.CodEmpleado BETWEEN '" & Me.DataCombo1.Columns(1).Text & "' AND '" & DataCombo2.Columns(1).Text & "') " & _
            "ORDER BY Empleado.CodEmpleado1, Nomina.NumNomina "
      
      
      End If
  
  
    rpt.DataControl1.Source = sql
'    ArepDetalleDeducciones.Show 1
           fPreview.arv.ReportSource = rpt
           fPreview.Show 1
 
Case "Resumen-Pago Mensual"
     
     
     Exportar = True
     Me.CommonDialog1.ShowSave
     Directorio = ""
     Directorio = Me.CommonDialog1.FileName + ".xls"
     
     Set rpt = New ArepResumen
     
     rpt.DataControl1.ConnectionString = ConexionReporte
     rpt.lblTitulo.Caption = Titulo
     rpt.LblSubtitulo.Caption = SubTitulo
     rpt.ImgLogo.Picture = LoadPicture(RutaLogo)
     
     If Me.DBCAño.Text = "" Then
       MsgBox "Seleecion un Año", vbCritical, "Sistema de Nominas"
       Exit Sub
     End If
     
'      FMes (Combo1.Text)
'      mes1 = Nmes
'      FMes (Combo2.Text)
'      Mes2 = Nmes
     
    If Combo1.Text = "" Or Me.Combo2.Text = "" Then

    sql = "SELECT  Empleado.CodEmpleado1, Empleado.CodEmpleado," & vbLf
    sql = sql & "Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Departamento.Departamento," & vbLf
    sql = sql & "Historico.FechaContrato, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.Incentivos, DetalleNomina.OtrosIngresos, DetalleNomina.SeptimoDia," & vbLf
    sql = sql & "DetalleNomina.SalarioBasico + DetalleNomina.OtrosIngresos + DetalleNomina.SeptimoDia + DetalleNomina.Destajo + DetalleNomina.Incentivos AS Sueldo," & vbLf
    sql = sql & "DetalleNomina.HorasExtras, DetalleNomina.MontoINSS, Nomina.FechaNominaINI, Nomina.FechaNomina," & vbLf
    sql = sql & "DetalleNomina.SalarioBasico + DetalleNomina.OtrosIngresos + DetalleNomina.SeptimoDia + DetalleNomina.Destajo + DetalleNomina.HorasExtras - DetalleNomina.MontoInss" & vbLf
    sql = sql & "AS Neto, Nomina.Mes, Nomina.Ano, Nomina.Periodo" & vbLf
    sql = sql & "FROM         Empleado INNER JOIN" & vbLf
    sql = sql & "Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN" & vbLf
    sql = sql & "Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN" & vbLf
    sql = sql & "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN" & vbLf
    sql = sql & "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina" & vbLf
    sql = sql & "WHERE     (DetalleNomina.SalarioBasico + DetalleNomina.Destajo <> 0) AND (Empleado.CodEmpleado Between '" & Me.DataCombo1.Columns(1).Text & "' AND '" & DataCombo2.Columns(1).Text & "')" & vbLf
    sql = sql & "ORDER BY Empleado.CodEmpleado, Nomina.Ano, Nomina.Mes,Nomina.Periodo"
      rpt.DataControl1.Source = sql

      
    ElseIf DataCombo1.Text = "" Or Me.DataCombo2.Text = "" Then
    
    sql = "SELECT  Empleado.CodEmpleado1, Empleado.CodEmpleado," & vbLf
    sql = sql & "Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Departamento.Departamento," & vbLf
    sql = sql & "Historico.FechaContrato, DetalleNomina.SalarioBasico, DetalleNomina.Incentivos, DetalleNomina.Destajo, DetalleNomina.OtrosIngresos, DetalleNomina.SeptimoDia," & vbLf
    sql = sql & "DetalleNomina.SalarioBasico + DetalleNomina.OtrosIngresos + DetalleNomina.SeptimoDia + DetalleNomina.Destajo  + DetalleNomina.Incentivos AS Sueldo," & vbLf
    sql = sql & "DetalleNomina.HorasExtras, DetalleNomina.MontoINSS, Nomina.FechaNominaINI, Nomina.FechaNomina," & vbLf
    sql = sql & "DetalleNomina.SalarioBasico + DetalleNomina.OtrosIngresos + DetalleNomina.SeptimoDia + DetalleNomina.Destajo + DetalleNomina.HorasExtras - DetalleNomina.MontoInss" & vbLf
    sql = sql & "AS Neto, Nomina.Mes, Nomina.Ano, Nomina.Periodo" & vbLf
    sql = sql & "FROM         Empleado INNER JOIN" & vbLf
    sql = sql & "Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN" & vbLf
    sql = sql & "Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN" & vbLf
    sql = sql & "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN" & vbLf
    sql = sql & "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina" & vbLf
    sql = sql & "WHERE (Nomina.FechaNomina BETWEEN '" & Format(Me.TxtFecha1, "yyyymmdd") & "' AND '" & Format(Me.TxtFecha2, "yyyymmdd") & "') AND (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') AND (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo <> 0)" & vbLf
    sql = sql & "ORDER BY Empleado.CodEmpleado, Nomina.Ano, Nomina.Mes, Nomina.Periodo"
    
    rpt.DataControl1.Source = sql
    ElseIf DataCombo1.Text <> "" And Me.DataCombo2.Text <> "" And Combo1.Text <> "" And Combo2.Text <> "" Then

    sql = "SELECT  Empleado.CodEmpleado1, Empleado.CodEmpleado," & vbLf
    sql = sql & "Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Departamento.Departamento," & vbLf
    sql = sql & "Historico.FechaContrato, DetalleNomina.SalarioBasico, DetalleNomina.Incentivos, DetalleNomina.Destajo, DetalleNomina.OtrosIngresos, DetalleNomina.SeptimoDia," & vbLf
    sql = sql & "DetalleNomina.SalarioBasico + DetalleNomina.OtrosIngresos + DetalleNomina.SeptimoDia + DetalleNomina.Destajo + DetalleNomina.Incentivos AS Sueldo," & vbLf
    sql = sql & "DetalleNomina.HorasExtras, DetalleNomina.MontoINSS, Nomina.FechaNominaINI, Nomina.FechaNomina," & vbLf
    sql = sql & "DetalleNomina.SalarioBasico + DetalleNomina.OtrosIngresos + DetalleNomina.SeptimoDia + DetalleNomina.Destajo + DetalleNomina.HorasExtras - DetalleNomina.MontoInss" & vbLf
    sql = sql & "AS Neto, Nomina.Mes, Nomina.Ano, Nomina.Periodo" & vbLf
    sql = sql & "FROM         Empleado INNER JOIN" & vbLf
    sql = sql & "Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN" & vbLf
    sql = sql & "Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN" & vbLf
    sql = sql & "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN" & vbLf
    sql = sql & "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina" & vbLf
    sql = sql & "WHERE (Nomina.FechaNomina BETWEEN '" & Format(Me.TxtFecha1, "dd/MM/yyyy") & "' AND '" & Format(Me.TxtFecha2, "dd/MM/yyyy") & "') AND (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') AND (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo <> 0) AND (Empleado.CodEmpleado1 Between '" & Me.DataCombo1.Columns(0).Text & "' AND '" & DataCombo2.Columns(0).Text & "')" & vbLf
    sql = sql & "ORDER BY Empleado.CodEmpleado, Nomina.Ano, Nomina.Mes, Nomina.Periodo"
      

      rpt.DataControl1.Refresh
      rpt.DataControl1.Source = sql

    End If
     
   'ArepResumen.Show 1
           fPreview.arv.ReportSource = rpt
           fPreview.Show 1

Case "Total-Pago Mensual":
    
   Exportar = True
   
   Me.CommonDialog1.ShowSave
   Directorio = ""
   Directorio = Me.CommonDialog1.FileName + ".xls"
    
   ArepTotalSemanaPago.DataControl1.ConnectionString = ConexionReporte
   ArepTotalSemanaPago.lblTitulo.Caption = Titulo
   ArepTotalSemanaPago.LblSubtitulo.Caption = SubTitulo
   ArepTotalSemanaPago.ImgLogo.Picture = LoadPicture(RutaLogo)
     
     If Me.DBCAño.Text = "" Then
       MsgBox "Seleccione un Año", vbCritical, "Sistema de Nominas"
       Exit Sub
     End If
     
'      FMes (Combo1.Text)
'      mes1 = Nmes
'      FMes (Combo2.Text)
'      Mes2 = Nmes
     
     If Combo1.Text = "" Or Me.Combo2.Text = "" Then

    sql = "SELECT  Empleado.CodEmpleado1, Empleado.CodEmpleado," & vbLf
    sql = sql & "Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Departamento.Departamento, " & vbLf
    sql = sql & "Empleado.Numeroinss, Empleado.TarifaHoraria, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Historico.FechaContrato, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.Incentivos,  DetalleNomina.OtrosIngresos, DetalleNomina.SeptimoDia," & vbLf
    sql = sql & "DetalleNomina.SalarioBasico + DetalleNomina.OtrosIngresos + DetalleNomina.SeptimoDia + DetalleNomina.Destajo + DetalleNomina.Incentivos AS Sueldo," & vbLf
    sql = sql & "DetalleNomina.HorasExtras, DetalleNomina.MontoINSS, Nomina.FechaNominaINI, Nomina.FechaNomina," & vbLf
    sql = sql & "DetalleNomina.SalarioBasico + DetalleNomina.OtrosIngresos + DetalleNomina.SeptimoDia + DetalleNomina.Destajo + DetalleNomina.HorasExtras - DetalleNomina.MontoInss" & vbLf
    sql = sql & "AS Neto, Nomina.Mes, Nomina.Ano, Nomina.Periodo" & vbLf
    sql = sql & "FROM         Empleado INNER JOIN" & vbLf
    sql = sql & "Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN" & vbLf
    sql = sql & "Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN" & vbLf
    sql = sql & "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN" & vbLf
    sql = sql & "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina" & vbLf
    sql = sql & "WHERE     (DetalleNomina.SalarioBasico + DetalleNomina.Destajo <> 0) AND (Empleado.CodEmpleado Between '" & Me.DataCombo1.Columns(1).Text & "' AND '" & DataCombo2.Columns(1).Text & "')" & vbLf
    sql = sql & "ORDER BY Empleado.CodEmpleado, Nomina.Ano, Nomina.Mes,Nomina.Periodo"
    ArepTotalSemanaPago.DataControl1.Source = sql

      
    ElseIf DataCombo1.Text = "" Or Me.DataCombo2.Text = "" Then
    
    sql = "SELECT  Empleado.CodEmpleado1, Empleado.CodEmpleado," & vbLf
    sql = sql & "Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Departamento.Departamento," & vbLf
    sql = sql & "Empleado.Numeroinss, Empleado.TarifaHoraria, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Historico.FechaContrato, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.Incentivos, DetalleNomina.OtrosIngresos, DetalleNomina.SeptimoDia," & vbLf
    sql = sql & "DetalleNomina.SalarioBasico + DetalleNomina.OtrosIngresos + DetalleNomina.SeptimoDia + DetalleNomina.Destajo + DetalleNomina.Incentivos AS Sueldo," & vbLf
    sql = sql & "DetalleNomina.HorasExtras, DetalleNomina.MontoINSS, Nomina.FechaNominaINI, Nomina.FechaNomina," & vbLf
    sql = sql & "DetalleNomina.SalarioBasico + DetalleNomina.OtrosIngresos + DetalleNomina.SeptimoDia + DetalleNomina.Destajo + DetalleNomina.HorasExtras - DetalleNomina.MontoInss" & vbLf
    sql = sql & "AS Neto, Nomina.Mes, Nomina.Ano, Nomina.Periodo" & vbLf
    sql = sql & "FROM         Empleado INNER JOIN" & vbLf
    sql = sql & "Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN" & vbLf
    sql = sql & "Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN" & vbLf
    sql = sql & "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN" & vbLf
    sql = sql & "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina" & vbLf
    sql = sql & "WHERE (Nomina.FechaNomina BETWEEN '" & Format(Me.TxtFecha1, "yyyymmdd") & "' AND '" & Format(Me.TxtFecha2, "yyyymmdd") & "') AND (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') AND (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo <> 0)" & vbLf
    sql = sql & "ORDER BY Empleado.CodEmpleado, Nomina.Ano, Nomina.Mes, Nomina.Periodo"
    
    ArepTotalSemanaPago.DataControl1.Source = sql
    ElseIf DataCombo1.Text <> "" And Me.DataCombo2.Text <> "" And Combo1.Text <> "" And Combo2.Text <> "" Then

    sql = "SELECT  Empleado.CodEmpleado1, Empleado.CodEmpleado," & vbLf
    sql = sql & "Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Departamento.Departamento," & vbLf
    sql = sql & "Empleado.Numeroinss, Empleado.TarifaHoraria, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Historico.FechaContrato, DetalleNomina.SalarioBasico, DetalleNomina.Incentivos, DetalleNomina.Destajo, DetalleNomina.OtrosIngresos, DetalleNomina.SeptimoDia," & vbLf
    sql = sql & "DetalleNomina.SalarioBasico + DetalleNomina.OtrosIngresos + DetalleNomina.SeptimoDia + DetalleNomina.Destajo + DetalleNomina.Incentivos AS Sueldo," & vbLf
    sql = sql & "DetalleNomina.HorasExtras, DetalleNomina.MontoINSS, Nomina.FechaNominaINI, Nomina.FechaNomina," & vbLf
    sql = sql & "DetalleNomina.SalarioBasico + DetalleNomina.OtrosIngresos + DetalleNomina.SeptimoDia + DetalleNomina.Destajo + DetalleNomina.HorasExtras - DetalleNomina.MontoInss" & vbLf
    sql = sql & "AS Neto, Nomina.Mes, Nomina.Ano, Nomina.Periodo" & vbLf
    sql = sql & "FROM         Empleado INNER JOIN" & vbLf
    sql = sql & "Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN" & vbLf
    sql = sql & "Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN" & vbLf
    sql = sql & "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado INNER JOIN" & vbLf
    sql = sql & "Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina" & vbLf
    sql = sql & "WHERE (Nomina.FechaNomina BETWEEN '" & Format(Me.TxtFecha1, "yyyymmdd") & "' AND '" & Format(Me.TxtFecha2, "yyyymmdd") & "') AND (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "') AND (dbo.DetalleNomina.SalarioBasico + dbo.DetalleNomina.Destajo <> 0) AND (Empleado.CodEmpleado Between '" & Me.DataCombo1.Columns(1).Text & "' AND '" & DataCombo2.Columns(1).Text & "')" & vbLf
    sql = sql & "ORDER BY Empleado.CodEmpleado, Nomina.Ano, Nomina.Mes, Nomina.Periodo"
      

      
    ArepTotalSemanaPago.DataControl1.Source = sql

    End If
    
'
    
    ArepTotalSemanaPago.lblNoNomina.Caption = Me.DBTipoNominas.Columns(0).Text
    ArepTotalSemanaPago.lblDescNomina.Caption = Me.DBTipoNominas.Columns(1).Text
    ArepTotalSemanaPago.lblFechaDesde.Caption = Me.TxtFecha1.Value
    ArepTotalSemanaPago.lblFechaHasta.Caption = Me.TxtFecha2.Value
     
 
'    ArepTotalSemanaPago.Show 1
           fPreview.arv.ReportSource = ArepTotalSemanaPago
           fPreview.Show 1






Case "Lista de Empleados Activos"
     ArepActivos.DataControl1.ConnectionString = ConexionReporte
     ArepActivos.lblTitulo.Caption = Titulo
     ArepActivos.LblSubtitulo.Caption = SubTitulo
     ArepActivos.ImgLogo.Picture = LoadPicture(RutaLogo)
'     ArepActivos.DataControl1.Source = "SELECT     Empleado.CodEmpleado1, Empleado.CodEmpleado,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Departamento.Departamento, TipoNomina.Nomina , Empleado.Activo, Empleado.CodTipoNomina FROM  Empleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina WHERE     (Empleado.Activo = 1) AND (Empleado.CodTipoNomina =  '" & Me.TDBCombo1.Columns(0).Text & "') ORDER BY Empleado.CodEmpleado1"
  
'           fPreview.arv.ReportSource = ArepActivos
'           fPreview.Show 1
     sql = "SELECT  Empleado.CodEmpleado1, Empleado.CodEmpleado,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Departamento.Departamento, TipoNomina.Nomina , Empleado.Activo, Empleado.CodTipoNomina FROM  Empleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina WHERE     (Empleado.Activo = 1) AND (Empleado.CodTipoNomina =  '" & Me.TDBCombo1.Columns(0).Text & "') ORDER BY Empleado.CodEmpleado1"
     Set rpt = New ArepActivos
     rpt.DataControl1.ConnectionString = Conexion
     rpt.DataControl1.Source = sql
     fPreview.RunReport rpt
     fPreview.Show 1



Case "Reporte INSS"
      ArepInssEmpleado.DataControl1.ConnectionString = ConexionReporte
      ArepInssEmpleado.lblTitulo.Caption = Titulo
      ArepInssEmpleado.LblSubtitulo.Caption = SubTitulo
      ArepInssEmpleado.ImgLogo.Picture = LoadPicture(RutaLogo)
      ArepInssEmpleado.DataControl1.Source = "SELECT Empleado.CodEmpleado1,[Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, PagosMensuales.Mes, PagosMensuales.Vacaciones, PagosMensuales.Anno, PagosMensuales.TotalIngresos, PagosMensuales.INSS, PagosMensuales.INSSPatronal, Cargo.Cargo, [PagosMensuales].[TotalIngresos]+[PagosMensuales].[Vacaciones] AS TIngresos FROM (Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN PagosMensuales ON Empleado.CodEmpleado = PagosMensuales.CodEmpleado  Where (((PagosMensuales.Mes) = " & Mes & ") And ((PagosMensuales.Anno) = " & Anno & ")) ORDER BY Empleado.CodEmpleado1"
      Mese = ConvertirMes(Mes)
      ArepInssEmpleado.LblFecha.Caption = Mese + "   del   " + Ano
      ArepInssEmpleado.LblFechaHoy = Format(Now, "dd/mm/yyyy")
'      ArepInssEmpleado.Show 1

           fPreview.arv.ReportSource = ArepInssEmpleado
           fPreview.Show 1
 Case "Reporte IR"
      ArepMensualIR.DataControl1.ConnectionString = ConexionReporte
      ArepMensualIR.lblTitulo.Caption = Titulo
      ArepMensualIR.LblSubtitulo.Caption = SubTitulo
      ArepMensualIR.ImgLogo.Picture = LoadPicture(RutaLogo)
      ArepMensualIR.DataControl1.Source = "SELECT Empleado.CodEmpleado1,[Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, PagosMensuales.Mes, PagosMensuales.Vacaciones, PagosMensuales.Anno, PagosMensuales.TotalIngresos, PagosMensuales.INSS, PagosMensuales.INSSPatronal, Cargo.Cargo, [PagosMensuales].[TotalIngresos]+[PagosMensuales].[Vacaciones] AS TIngresos, PagosMensuales.IR FROM (Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN PagosMensuales ON Empleado.CodEmpleado = PagosMensuales.CodEmpleado  Where (((PagosMensuales.Mes) = " & Mes & ") And ((PagosMensuales.Anno) = " & Anno & ")) ORDER BY Empleado.CodEmpleado1"
      Mese = ConvertirMes(Mes)
      ArepMensualIR.LblFecha.Caption = Mese + "   del   " + Ano
      ArepMensualIR.LblFechaHoy = Format(Now, "dd/mm/yyyy")
'      ArepMensualIR.Show 1

           fPreview.arv.ReportSource = ArepMensualIR
           fPreview.Show 1
      
      
Case "Listado de Reportes"
      MsgBox "Seleccione el Reporte que desea ver"
      Exit Sub
Case "Listado de Empleados"
        ACListEmpleados.DataControl1.ConnectionString = ConexionReporte
        'ACListEmpleados.DataControl1.Source
'        ACListEmpleados.Show 1

           fPreview.arv.ReportSource = ACListEmpleados
           fPreview.Show 1
           
Case "Listado de Cargos"
        ARListCargos.DataControl1.ConnectionString = ConexionReporte
'        ARListCargos.Show 1
           fPreview.arv.ReportSource = ARListCargos
           fPreview.Show 1
           
Case "Listado de Departamentos"
        ARListDepartamentos.DataControl1.ConnectionString = ConexionReporte
'        ARListDepartamentos.Show 1
           fPreview.arv.ReportSource = ARListDepartamentos
           fPreview.Show 1
Case "Listado de Tipos de Subsidios"
        ARListSubsidios.DataControl1.ConnectionString = ConexionReporte
'        ARListSubsidios.Show 1
           fPreview.arv.ReportSource = ARListSubsidios
           fPreview.Show 1
Case "Listado de Tipos de Incentivos"
      ARListIncentivos.DataControl1.ConnectionString = ConexionReporte
'      ARListIncentivos.Show 1
           fPreview.arv.ReportSource = ARListIncentivos
           fPreview.Show 1
Case "Listado de Tipos de Deducciones"
      ARlistDeducciones.DataControl1.ConnectionString = ConexionReporte
'      ARlistDeducciones.Show 1
           fPreview.arv.ReportSource = ARlistDeducciones
           fPreview.Show 1
End Select

End Sub

Private Sub DataCombo4_Click(Area As Integer)
Dim FechaIni As String, FechaFin As String

 Me.AdoBusca.RecordSource = "SELECT Periodo, año, CodTipoNomina, mes, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE  (Periodo = '" & Me.DBComboPeriodo.Text & "') AND (CodTipoNomina = '" & Me.TxtNumero.Text & "') AND (año = '" & Me.DBAño.Text & "')"
 Me.AdoBusca.Refresh
 InputBox "", "", Me.AdoBusca.RecordSource
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.MtxtFechaini = Me.AdoBusca.Recordset("Inicio")
   Me.MtxtFecha = Me.AdoBusca.Recordset("Final")
 End If
 
 FechaIni = Mid(Me.MtxtFechaini.Value, 7, 4) & "-" & Mid(Me.MtxtFechaini.Value, 4, 2) & "-" & Mid(Me.MtxtFechaini.Value, 1, 2)

FechaFin = Mid(Me.MtxtFecha.Value, 7, 4) & "-" & Mid(Me.MtxtFecha.Value, 4, 2) & "-" & Mid(Me.MtxtFecha.Value, 1, 2)
'& "/" & Mid(Me.MtxtFecha.Value, 4, 2) & "/" & Mid(Me.MtxtFecha.Value.Value, 1, 2)

 
 
 
 Me.AdoBusca.RecordSource = "SELECT NumNomina, CodTipoNomina, FechaNominaINI, FechaNomina From Nomina WHERE (FechaNominaINI = CONVERT(DATETIME, '" & FechaIni & " 00:00:00', 102)) AND (FechaNomina = CONVERT(DATETIME, '" & FechaFin & " 00:00:00', 102))"
 Me.AdoBusca.Refresh
 InputBox "", "", Me.AdoBusca.RecordSource
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.TxtNNomina.Text = Me.AdoBusca.Recordset("NumNomina")
 Else
   Me.TxtNNomina.Text = ""
 End If

End Sub

Private Sub Combo1_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String
If Not Me.TxtNumeros.Text = "" Then
    Me.AdoPeriodo.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & Me.TxtNumeros.Text & "')AND (año = '" & Me.DBCAño.Text & "')"
'    InputBox "", "", Me.AdoPeriodo.RecordSource
    Me.AdoPeriodo.Refresh
    Me.DBPeriodos.ListField = "Periodo"
End If

FMes (Combo1.Text)
Mes1 = Format(Nmes, "0#")
FMes (Combo2.Text)
Mes2 = Format(Nmes, "0#")
Año1 = val(Me.DBCAño.Text)
Año2 = val(Me.DBAño2.Text)
CodTipoNomina = Me.DBTipoNominas.Columns(0).Text


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.TxtFecha1.Value = Me.AdoBusca.Recordset("Inicio")
 End If
 
Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.TxtFecha2.Value = Me.AdoBusca.Recordset("Final")
 End If
End Sub

Private Sub Combo2_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String
If Not Me.TxtNumeros.Text = "" Then
    Me.AdoPeriodo.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & Me.TxtNumeros.Text & "')AND (año = '" & Me.DBCAño.Text & "')"
'    InputBox "", "", Me.AdoPeriodo.RecordSource
    Me.AdoPeriodo.Refresh
    Me.DBPeriodos.ListField = "Periodo"
End If

FMes (Combo1.Text)
Mes1 = Format(Nmes, "0#")
FMes (Combo2.Text)
Mes2 = Format(Nmes, "0#")
Año1 = val(Me.DBCAño.Text)
Año2 = val(Me.DBAño2.Text)
CodTipoNomina = Me.DBTipoNominas.Columns(0).Text


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.TxtFecha1.Value = Me.AdoBusca.Recordset("Inicio")
 End If
 
Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.TxtFecha2.Value = Me.AdoBusca.Recordset("Final")
 End If
End Sub

Private Sub Combo3_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String
If Not Me.TxtNumero.Text = "" Then
    Me.AdoPeriodo.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & Me.TxtNumeros.Text & "')AND (año = '" & Me.DBCAño.Text & "')"
'    InputBox "", "", Me.AdoPeriodo.RecordSource
    Me.AdoPeriodo.Refresh
    Me.DBPeriodos.ListField = "Periodo"
End If

FMes (Combo3.Text)
Mes1 = Format(Nmes, "0#")
FMes (Combo4.Text)
Mes2 = Format(Nmes, "0#")
Año1 = val(Me.DBAño.Text)
CodTipoNomina = Me.TDBCombo1.Columns(0).Text


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.MtxtFechaini.Value = Me.AdoBusca.Recordset("Inicio")
 End If
 
Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.MtxtFecha.Value = Me.AdoBusca.Recordset("Final")
 End If
End Sub

Private Sub Combo4_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String
If Not Me.TxtNumero.Text = "" Then
    Me.AdoPeriodo.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & Me.TxtNumeros.Text & "')AND (año = '" & Me.DBCAño.Text & "')"
'    InputBox "", "", Me.AdoPeriodo.RecordSource
    Me.AdoPeriodo.Refresh
    Me.DBPeriodos.ListField = "Periodo"
End If

FMes (Combo3.Text)
Mes1 = Format(Nmes, "0#")
FMes (Combo4.Text)
Mes2 = Format(Nmes, "0#")
Año1 = val(Me.DBAño.Text)

CodTipoNomina = Me.TDBCombo1.Columns(0).Text


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.MtxtFechaini.Value = Me.AdoBusca.Recordset("Inicio")
 End If
 
Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.MtxtFecha.Value = Me.AdoBusca.Recordset("Final")
 End If
End Sub

Private Sub Command1_Click()
Dim DiasMes As Double
Dim CodTipoNomina As String, AjusteINSS As Double, MontoInssBasico As Double, MontoInss As Double
Dim TarifaHorariaBasico As Double, TasaInss As Double, NumNomina As Double
Dim rs As New ADODB.Recordset

MDIPrimero.DtaControles.Refresh
DiasMes = MDIPrimero.DtaControles.Recordset("DiasMes")
TasaInss = 6.25


CodTipoNomina = Me.DBTipoNominas.Columns(0).Text

    MDIPrimero.DtaConsulta.RecordSource = "SELECT * From TipoNomina WHERE (CodTipoNomina = '" & CodTipoNomina & "')"
    MDIPrimero.DtaConsulta.Refresh
    If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
     If Not IsNull(MDIPrimero.DtaConsulta.Recordset("TarifaHoraria")) Then
      TarifaHorariaBasico = MDIPrimero.DtaConsulta.Recordset("TarifaHoraria")
     End If
    Else
      TarifaHorariaBasico = 0
    End If
    

 DoEvents
    
 sql = "SELECT TOP 100 PERCENT Empleado.Nombre1, Empleado.Apellido1, DetalleNomina.CodEmpleado AS CodEmpleado, DetalleNomina.MontoINSS AS MontoInss, " & _
       "DetalleNomina.INSSPatronal AS InssPatronal, " & _
       "DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + " & _
       "DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion AS TotalDevengado, " & _
       "DetalleNomina.INATEC AS MontoInatec, Empleado.NumeroInss AS NumeroInss, DetalleNomina.MontoINSS + DetalleNomina.INSSPatronal AS TotalInss, " & _
       "Empleado.CodEmpleado1 AS CodEmpleado1, DetalleNomina.NumNomina AS NumNomina, Nomina.FechaNomina AS Nomina, " & _
       "Cargo.Cargo AS Cargo, Nomina.Mes, Nomina.Ano, Nomina.Periodo, TipoNomina.Periodo AS PeriodoNomina,Empleado.NumCedula, DetalleNomina.AjusteINSS " & _
       "FROM         Nomina INNER JOIN " & _
       "Grupo INNER JOIN " & _
       "Cargo INNER JOIN " & _
       "TipoNomina INNER JOIN " & _
       "Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN  " & _
       "DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON " & _
       "TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina " & _
       "WHERE     (Nomina.FechaNomina BETWEEN '" & Format(Me.TxtFecha1.Value, "yyyymmdd") & "' And '" & Format(Me.TxtFecha2.Value, "yyyymmdd") & "') " & _
       "AND (Nomina.CodTipoNomina = '" & Me.DBTipoNominas.Columns(0).Text & "' ) " & _
       "ORDER BY Empleado.CodEmpleado1, Nomina.FechaNomina"
 Me.AdoBusca.RecordSource = sql
 AdoBusca.Refresh
 Me.AdoBusca.Recordset.MoveLast
 
 Barra.Min = 0
 Barra.Max = Me.AdoBusca.Recordset.RecordCount
 Barra.Value = 0
 
 Me.AdoBusca.Refresh
 Do While Not Me.AdoBusca.Recordset.EOF
   DoEvents
   NumNomina = Me.AdoBusca.Recordset("NumNomina")
   CodEmpleado = Me.AdoBusca.Recordset("CodEmpleado")
   CantSabados = SemanasPeriodos(Me.AdoBusca.Recordset("Ano"), Me.AdoBusca.Recordset("Mes"), CodTipoNomina)
   MontoInss = Me.AdoBusca.Recordset("MontoINSS")
    
    '//////////////Busco si la Nomina Existe para Editarla/////////////////
         AjusteINSS = 0
     MDIPrimero.DtaConsulta.RecordSource = "SELECT DetalleNomina.id, DetalleNomina.BonoProduccion ,DetalleNomina.IncetivoProduccion,DetalleNomina.NumNomina, DetalleNomina.CodEmpleado, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HE, DetalleNomina.DD, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.INATEC, DetalleNomina.Mes13, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.TotalSubsidio, DetalleNomina.VacacionesPagadas, DetalleNomina.DiasVacaciones,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.TarifaHoraria,DetalleNomina.produjo,DetalleNomina.AjusteINSS " & _
                                     " From DetalleNomina Where (((DetalleNomina.NumNomina) = " & NumNomina & ") And ((DetalleNomina.CodEmpleado) = '" & CodEmpleado & "'))"
     MDIPrimero.DtaConsulta.Refresh
         MontoInssBasico = ((TarifaHorariaBasico * 8 * DiasMes) * (TasaInss / 100) / CantSabados)
         If MontoInss > MontoInssBasico Then
           AjusteINSS = 0
         Else
           AjusteINSS = MontoInssBasico - MontoInss
         End If
         
     rs.Open "UPDATE DetalleNomina SET DetalleNomina.AjusteINSS = " & AjusteINSS & " WHERE (NumNomina = " & NumNomina & ") AND (CodEmpleado = '" & CodEmpleado & "') ", Conexion
         
     Me.AdoBusca.Recordset.MoveNext
     Barra.Value = Barra.Value + 1
  Loop
End Sub

Private Sub DBAño_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer


CodTipoNomina = Me.TDBCombo1.Columns(0).Text
If Not Me.TDBCombo1.Columns(0).Text = "" Then
    Me.AdoPeriodo.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & Me.TDBCombo1.Columns(0).Text & "')AND (año = '" & Me.DBAño.Text & "')"
'    InputBox "", "", Me.AdoPeriodo.RecordSource
    Me.AdoPeriodo.Refresh
    Me.DBComboPeriodo.ListField = "Periodo"
End If



End Sub

Private Sub DBAño2_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String
If Not Me.TxtNumeros.Text = "" Then
    Me.AdoPeriodo.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & Me.TxtNumeros.Text & "')AND (año = '" & Me.DBCAño.Text & "')"
'    InputBox "", "", Me.AdoPeriodo.RecordSource
    Me.AdoPeriodo.Refresh
    Me.DBPeriodos.ListField = "Periodo"
End If

FMes (Combo1.Text)
Mes1 = Format(Nmes, "0#")
FMes (Combo2.Text)
Mes2 = Format(Nmes, "0#")
Año1 = val(Me.DBCAño.Text)
Año2 = val(Me.DBAño2.Text)
CodTipoNomina = Me.DBTipoNominas.Columns(0).Text


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.TxtFecha1.Value = Me.AdoBusca.Recordset("Inicio")
 End If
 
Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.TxtFecha2.Value = Me.AdoBusca.Recordset("Final")
 End If
End Sub

Private Sub DBCAño_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String
If Not Me.TxtNumeros.Text = "" Then
    Me.AdoPeriodo.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & Me.TxtNumeros.Text & "')AND (año = '" & Me.DBCAño.Text & "')"
'    InputBox "", "", Me.AdoPeriodo.RecordSource
    Me.AdoPeriodo.Refresh
    Me.DBPeriodos.ListField = "Periodo"
End If

FMes (Combo1.Text)
Mes1 = Format(Nmes, "0#")
FMes (Combo2.Text)
Mes2 = Format(Nmes, "0#")
Año1 = val(Me.DBCAño.Text)
Año2 = val(Me.DBAño2.Text)
CodTipoNomina = Me.DBTipoNominas.Columns(0).Text


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.TxtFecha1.Value = Me.AdoBusca.Recordset("Inicio")
 End If
 
Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.TxtFecha2.Value = Me.AdoBusca.Recordset("Final")
 End If

End Sub

Private Sub DBComboPeriodo_Change()
Dim FechaIni As String, FechaFin As String
Dim NumNomina As Double

 Me.AdoBusca.RecordSource = "SELECT Periodo, año, CodTipoNomina, mes, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE  (Periodo = '" & Me.DBComboPeriodo.Text & "') AND (CodTipoNomina = '" & Me.TDBCombo1.Columns(0).Text & "') AND (año = '" & Me.DBAño.Text & "')"
 Me.AdoBusca.Refresh
 
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.MtxtFechaini = Me.AdoBusca.Recordset("Inicio")
   Me.MtxtFecha = Me.AdoBusca.Recordset("Final")
 End If
 
 FechaIni = Mid(Me.MtxtFechaini.Value, 7, 4) & "-" & Mid(Me.MtxtFechaini.Value, 4, 2) & "-" & Mid(Me.MtxtFechaini.Value, 1, 2)

FechaFin = Mid(Me.MtxtFecha.Value, 7, 4) & "-" & Mid(Me.MtxtFecha.Value, 4, 2) & "-" & Mid(Me.MtxtFecha.Value, 1, 2)
'& "/" & Mid(Me.MtxtFecha.Value, 4, 2) & "/" & Mid(Me.MtxtFecha.Value.Value, 1, 2)

 
 
 
 Me.AdoBusca.RecordSource = "SELECT NumNomina, CodTipoNomina, FechaNominaINI, FechaNomina From Nomina WHERE (FechaNominaINI = CONVERT(DATETIME, '" & FechaIni & " 00:00:00', 102)) AND (FechaNomina = CONVERT(DATETIME, '" & FechaFin & " 00:00:00', 102))"
 Me.AdoBusca.Refresh
' InputBox "", "", Me.AdoBusca.RecordSource
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.TxtNNomina.Text = Me.AdoBusca.Recordset("NumNomina")
   NumNomina = Me.AdoBusca.Recordset("NumNomina")
 Else
   Me.TxtNNomina.Text = ""
   NumNomina = -1
 End If
 
 
 Me.AdoEmpleadoActivo.RecordSource = "SELECT Empleado.CodEmpleado1, DetalleNomina.CodEmpleado, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres FROM  DetalleNomina INNER JOIN Empleado ON DetalleNomina.CodEmpleado = Empleado.CodEmpleado Where (DetalleNomina.NumNomina = " & NumNomina & " ) ORDER BY Empleado.CodEmpleado1"
 Me.AdoEmpleadoActivo.Refresh
 

End Sub

Private Sub DBPeriodos_Change()
Dim FechaIni As String, FechaFin As String

 Me.AdoBusca.RecordSource = "SELECT Periodo, año, CodTipoNomina, mes, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE  (Periodo = '" & Me.DBComboPeriodo.Text & "') AND (CodTipoNomina = '" & Me.TxtNumero.Text & "') AND (año = '" & Me.DBAño.Text & "')"
 Me.AdoBusca.Refresh
' InputBox "", "", Me.AdoBusca.RecordSource
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.TxtFecha1 = Me.AdoBusca.Recordset("Inicio")
   Me.TxtFecha2 = Me.AdoBusca.Recordset("Final")
 End If
 
 FechaIni = Mid(Me.MtxtFechaini.Value, 7, 4) & "-" & Mid(Me.MtxtFechaini.Value, 4, 2) & "-" & Mid(Me.MtxtFechaini.Value, 1, 2)

FechaFin = Mid(Me.MtxtFecha.Value, 7, 4) & "-" & Mid(Me.MtxtFecha.Value, 4, 2) & "-" & Mid(Me.MtxtFecha.Value, 1, 2)
'& "/" & Mid(Me.MtxtFecha.Value, 4, 2) & "/" & Mid(Me.MtxtFecha.Value.Value, 1, 2)

 
 
 
 Me.AdoBusca.RecordSource = "SELECT NumNomina, CodTipoNomina, FechaNominaINI, FechaNomina From Nomina WHERE (FechaNominaINI = CONVERT(DATETIME, '" & FechaIni & " 00:00:00', 102)) AND (FechaNomina = CONVERT(DATETIME, '" & FechaFin & " 00:00:00', 102))"
 Me.AdoBusca.Refresh
' InputBox "", "", Me.AdoBusca.RecordSource
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.TxtNumNominas.Text = Me.AdoBusca.Recordset("NumNomina")
 Else
   Me.TxtNumNominas.Text = ""
 End If
End Sub

Private Sub DBTipoNomina_Change()
Me.AdoBusca.RecordSource = "SELECT CodTipoNomina, Nomina, Periodo, UltFecha From TipoNomina WHERE  (Nomina = '" & Me.DbTipoNomina.Caption & "')"
Me.AdoBusca.Refresh
If Not Me.AdoBusca.Recordset.EOF Then
  Me.TxtNumero.Text = Me.AdoBusca.Recordset("CodTipoNomina")
End If

End Sub

Private Sub DBTipoNominas_ItemChange()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String
If Not Me.TxtNumeros.Text = "" Then
    Me.AdoPeriodo.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & Me.TxtNumeros.Text & "')AND (año = '" & Me.DBCAño.Text & "')"
'    InputBox "", "", Me.AdoPeriodo.RecordSource
    Me.AdoPeriodo.Refresh
    Me.DBPeriodos.ListField = "Periodo"
End If

FMes (Combo1.Text)
Mes1 = Format(Nmes, "0#")
FMes (Combo2.Text)
Mes2 = Format(Nmes, "0#")
Año1 = val(Me.DBCAño.Text)
Año2 = val(Me.DBAño2.Text)
CodTipoNomina = Me.DBTipoNominas.Columns(0).Text


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.TxtFecha1.Value = Me.AdoBusca.Recordset("Inicio")
 End If
 
Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.TxtFecha2.Value = Me.AdoBusca.Recordset("Final")
 End If
End Sub

'Private Sub DBTipoNominas_Click(Area As Integer)
'Me.AdoBusca.RecordSource = "SELECT CodTipoNomina, Nomina, Periodo, UltFecha From TipoNomina WHERE  (Nomina = '" & Me.DBTipoNominas.Text & "')"
'Me.AdoBusca.Refresh
'If Not Me.AdoBusca.Recordset.EOF Then
'  Me.TxtNumeros.Text = Me.AdoBusca.Recordset("CodTipoNomina")
'End If
'End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
Me.DTFecha1.Value = Format(Now, "DD/mm/yyyy")
Me.DTFecha2.Value = Format(Now, "DD/mm/yyyy")
Me.Mes.Value = Format(Now, "DD/mm/yyyy")

With Me.AdoAño
   .ConnectionString = Conexion
   .RecordSource = "SELECT DISTINCT Año From Año_Actual"
   .Refresh
 End With
 
  With Me.AdoSuspenciones
   .ConnectionString = Conexion
 End With
 
 With Me.AdoConsulta
   .ConnectionString = Conexion
 End With
 
  With Me.AdoAuxiliar
   .ConnectionString = Conexion
 End With
 
 With Me.AdoTarifa
   .ConnectionString = Conexion
 End With
 
  With Me.AdoReportes
   .ConnectionString = Conexion
 End With
 
  With Me.AdoNuevoIngreso
   .ConnectionString = Conexion
 End With
 
  With Me.AdoVacaciones
   .ConnectionString = Conexion
 End With
 
With Me.AdoBajas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
 End With
 

With Me.AdoDatosEmpresa
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
 End With

With Me.DtaNominas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
 End With

With Me.AdoBusca
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With


With Me.AdoPeriodo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With


With Me.AdoDeducciones
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT     CodTipoDeduccion, Deduccion, Tipo, CuentaContable  FROM         TipoDeduccion"
End With

Me.TDBCombo2.RowSource = Me.AdoDeducciones
Me.TDBCombo2.ListField = "CodTipoDeduccion"

Me.TDBCombo2.Columns(2).Visible = False
Me.TDBCombo2.Columns(3).Visible = False

With Me.AdoTipo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
    .RecordSource = "SELECT CodTipoNomina, Nomina, Periodo From TipoNomina"
   .Refresh
End With

With Me.AdoDepartamento
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
    .RecordSource = "SELECT * From Departamento"
   .Refresh
End With

With Me.AdoEmpleadoActivo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.adoEmpleado
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT CodEmpleado1,CodEmpleado, Nombre1 + ' '+ Nombre2 + ' ' + Apellido1 + ' ' + Apellido2 As Nombre From Empleado"
   .Refresh
End With


Me.DBAño2.ListField = "Año"
Me.DBCAño.ListField = "Año"
Me.DBAño.ListField = "Año"
'Me.BDAño2.ListField = "Año"
'Me.TDBCombo1.ListField = "CodEmpleado1"
'Me.TDBCombo1.Columns(1).Visible = False
'Me.TDBCombo1.Columns(0).Caption = "Codigo"
'Me.TDBCombo3.ListField = "CodEmpleado1"
'Me.TDBCombo3.Columns(1).Visible = False
'Me.TDBCombo3.Columns(0).Caption = "Codig
Me.DataCombo1.ListField = "CodEmpleado1"
Me.DataCombo2.ListField = "CodEmpleado1"
Me.DataCombo1.Columns(1).Visible = False
Me.DataCombo1.Columns(0).Caption = "Codigo"
Me.DataCombo2.Columns(1).Visible = False
Me.DataCombo2.Columns(0).Caption = "Codigo"


Me.DBTipoNominas.ListField = "Nomina"
'Me.DBTipoNomina.ListField = "Nomina"

With Me.DtaNomSubsidio
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

End Sub

Private Sub SmartButton1_Click()

End Sub

Private Sub TDBCombo1_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String
If Not Me.TxtNumero.Text = "" Then
    Me.AdoPeriodo.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & Me.TxtNumeros.Text & "')AND (año = '" & Me.DBCAño.Text & "')"
'    InputBox "", "", Me.AdoPeriodo.RecordSource
    Me.AdoPeriodo.Refresh
    Me.DBPeriodos.ListField = "Periodo"
End If

FMes (Combo3.Text)
Mes1 = Format(Nmes, "0#")
FMes (Combo4.Text)
Mes2 = Format(Nmes, "0#")
Año1 = val(Me.DBAño.Text)
CodTipoNomina = Me.TDBCombo1.Columns(0).Text


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.MtxtFechaini.Value = Me.AdoBusca.Recordset("Inicio")
 End If
 
Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.MtxtFecha.Value = Me.AdoBusca.Recordset("Final")
 End If
End Sub

Private Sub TDBCombo2_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String
If Not Me.TxtNumero.Text = "" Then
    Me.AdoPeriodo.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & Me.TxtNumeros.Text & "')AND (año = '" & Me.DBCAño.Text & "')"
'    InputBox "", "", Me.AdoPeriodo.RecordSource
    Me.AdoPeriodo.Refresh
    Me.DBPeriodos.ListField = "Periodo"
End If

FMes (Combo3.Text)
Mes1 = Format(Nmes, "0#")
FMes (Combo4.Text)
Mes2 = Format(Nmes, "0#")
Año1 = val(Me.DBAño.Text)

CodTipoNomina = Me.TDBCombo1.Columns(0).Text


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.MtxtFechaini.Value = Me.AdoBusca.Recordset("Inicio")
 End If
 
Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.MtxtFecha.Value = Me.AdoBusca.Recordset("Final")
 End If
End Sub

Private Sub TDBCombo3_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String
If Not Me.TxtNumero.Text = "" Then
    Me.AdoPeriodo.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & Me.TxtNumeros.Text & "')AND (año = '" & Me.DBCAño.Text & "')"
'    InputBox "", "", Me.AdoPeriodo.RecordSource
    Me.AdoPeriodo.Refresh
    Me.DBPeriodos.ListField = "Periodo"
End If

FMes (Combo3.Text)
Mes1 = Format(Nmes, "0#")
FMes (Combo4.Text)
Mes2 = Format(Nmes, "0#")
Año1 = val(Me.DBAño.Text)

CodTipoNomina = Me.TDBCombo1.Columns(0).Text


Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.MtxtFechaini.Value = Me.AdoBusca.Recordset("Inicio")
 End If
 
Me.AdoBusca.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
   Me.AdoBusca.Recordset.MoveLast
   Me.MtxtFecha.Value = Me.AdoBusca.Recordset("Final")
 End If
End Sub

