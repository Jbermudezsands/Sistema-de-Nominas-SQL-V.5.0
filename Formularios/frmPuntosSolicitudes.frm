VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPuntosSolicitudes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementos Salariales"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   12465
   Begin VB.Frame fracomplementos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   5040
      TabIndex        =   23
      Top             =   6600
      Width           =   7215
      Begin VB.TextBox txtSolicitados 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtSalario 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtValPorcentaje 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtValPts 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtCantPts 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtPorcentaje 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtPrecioPts 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1440
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosSolicitudes.frx":0000
         TabIndex        =   30
         Top             =   720
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosSolicitudes.frx":007A
         TabIndex        =   31
         Top             =   1080
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmPuntosSolicitudes.frx":00EC
         TabIndex        =   32
         Top             =   1440
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosSolicitudes.frx":0156
         TabIndex        =   33
         Top             =   1920
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "frmPuntosSolicitudes.frx":01BE
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "frmPuntosSolicitudes.frx":0227
         TabIndex        =   35
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   2880
         OleObjectBlob   =   "frmPuntosSolicitudes.frx":028A
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "frmPuntosSolicitudes.frx":02EF
         TabIndex        =   39
         Top             =   240
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   240
         X2              =   6960
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   240
         X2              =   6960
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2895
      Left            =   5040
      TabIndex        =   16
      Top             =   360
      Width           =   7215
      Begin TrueOleDBGrid80.TDBGrid tdbgEmpleados 
         Height          =   2295
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4048
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
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   5040
      OleObjectBlob   =   "frmPuntosSolicitudes.frx":0354
      TabIndex        =   21
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame fraSolicitud 
      Height          =   2895
      Left            =   5040
      TabIndex        =   18
      Top             =   3480
      Width           =   7215
      Begin VB.TextBox txtdirdoc 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   1680
         Visible         =   0   'False
         Width           =   4245
      End
      Begin VB.TextBox txtDocumento 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   4245
      End
      Begin VB.TextBox txtJustificacion 
         Height          =   645
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   4725
      End
      Begin VB.CommandButton cmddirdoc 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   330
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
         Left            =   6600
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
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
         Left            =   6600
         TabIndex        =   4
         Tag             =   "1"
         Top             =   240
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "frmPuntosSolicitudes.frx":03C4
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "frmPuntosSolicitudes.frx":043C
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpSolicitud 
         Height          =   300
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         Format          =   50003969
         CurrentDate     =   40729
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "frmPuntosSolicitudes.frx":04AC
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin TrueOleDBGrid80.TDBGrid tdbgPuntos 
         Height          =   2415
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
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
         Splits(0)._SavedRecordSelectors=   -1  'True
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
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   5040
      TabIndex        =   17
      Top             =   9120
      Width           =   7215
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   6000
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmPuntosSolicitudes.frx":052E
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ilnode 
      Left            =   3240
      Top             =   6720
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
            Picture         =   "frmPuntosSolicitudes.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPuntosSolicitudes.frx":0D5E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvActividades 
      Height          =   6135
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   10821
      _Version        =   393217
      Indentation     =   353
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ilnode"
      Appearance      =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView tvNominas 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4895
      _Version        =   393217
      Indentation     =   353
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ilnode"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmPuntosSolicitudes.frx":1B6E
      TabIndex        =   15
      Top             =   3360
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
      Height          =   255
      Left            =   5040
      OleObjectBlob   =   "frmPuntosSolicitudes.frx":1BD8
      TabIndex        =   37
      Top             =   3360
      Width           =   3375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   255
      Left            =   5040
      OleObjectBlob   =   "frmPuntosSolicitudes.frx":1C72
      TabIndex        =   38
      Top             =   6480
      Width           =   3375
   End
End
Attribute VB_Name = "frmPuntosSolicitudes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnx As New ADODB.Connection
Private rs As New ADODB.Recordset
Private rs3 As New ADODB.Recordset
Private rs4 As New ADODB.Recordset
Private sql As String
Private modal As Boolean

Private Sub cmdAgregarCta_Click()
On Error GoTo erradd
If Me.tdbgPuntos.Visible Then
    Me.tvActividades.Enabled = True
    Me.tvNominas.Enabled = False
    Me.tdbgEmpleados.Enabled = False
    Me.tdbgPuntos.Visible = False
    Me.SkinLabel10.Visible = True
    Me.SkinLabel13.Visible = True
    Me.SkinLabel14.Visible = True
    Me.dtpSolicitud.Visible = True
    Me.dtpSolicitud.Value = Now
    Me.txtJustificacion.Visible = True
    Me.txtJustificacion.Text = ""
    Me.txtDocumento.Visible = True
    Me.txtDocumento.Text = ""
    Me.txtdirdoc.Visible = True
    Me.txtdirdoc.Text = ""
    Me.cmddirdoc.Visible = True
    Me.cmdBorrarCta.Enabled = True
Else
    If tvActividades.SelectedItem.FullPath = tvActividades.SelectedItem.Text Then
        MsgBox "Seleccion la descripcion del punto a solicitar", vbInformation: Exit Sub
    ElseIf tvActividades.SelectedItem.Parent.Text = "ANTIGÜEDAD" Then
        MsgBox "Operación Cancelada, los puntos no pueden ser asignados por el usuario", vbInformation: Exit Sub
    ElseIf Trim(Me.txtJustificacion) = "" Then
        MsgBox "Justifique la solicitud de puntos", vbInformation
        Me.txtJustificacion.SetFocus
        Exit Sub
    End If
    
    sql = "INSERT INTO PUNTOSEMPLEADO(EMPLEADO, PUNTOS, APROBADO, FECHASOLICITUD, JUSTIFICACION, DOCUMENTO, DIRDOCUMENTO, ELIMINAR) VALUES(" & _
                val(Me.tdbgEmpleados.Columns(5).Value) & ", " & val(Mid(Me.tvActividades.SelectedItem.Key, 2)) & _
                ", 0, '" & Format$(Me.dtpSolicitud.Value, "yyyymmdd") & "', '" & Me.txtJustificacion.Text & "', '" & Me.txtDocumento.Text & "', '" & Me.txtdirdoc.Text & "',1)"
    cnx.Execute sql
    Call cmdBorrarCta_Click
    rs4.Find "ID = " & val(Mid(Me.tvActividades.SelectedItem.Key, 2))
End If
Exit Sub
erradd:
    If Err.Number = -2147217873 Then
        MsgBox "Operación cancelada, El punto ya ha sido solicitado", vbInformation
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub cmdBorrarCta_Click()
On Error GoTo errcta
If Me.tdbgPuntos.Visible Then
    If rs4.eof Then MsgBox "Operación Cancelada, no existen registros", vbInformation: Exit Sub
    If UCase(rs4!grupo) = "ANTIGÜEDAD" Then MsgBox "Operación Cancelada, el registro no puede ser eliminado", vbInformation: Exit Sub
    If MsgBox("¿Desea eliminar el registro?", vbYesNo) = vbYes Then
        sql = "DELETE FROM PUNTOSEMPLEADO WHERE empleado = " & rs4!CodEmpleado & " and puntos = " & rs4!Id
        cnx.Execute sql
    End If
Else
    Me.tvActividades.Enabled = False
    Me.tvNominas.Enabled = True
    Me.tdbgEmpleados.Enabled = True
    Me.tdbgPuntos.Visible = True
    Me.SkinLabel10.Visible = False
    Me.SkinLabel13.Visible = False
    Me.SkinLabel14.Visible = False
    Me.dtpSolicitud.Visible = False
    Me.txtJustificacion.Visible = False
    Me.txtDocumento.Visible = False
    Me.txtdirdoc.Visible = False
    Me.cmddirdoc.Visible = False
    Me.cmdBorrarCta.Enabled = False
End If

Call tdbgEmpleados_RowColChange(Me.tdbgEmpleados.Row, Me.tdbgEmpleados.col)

Exit Sub
errcta:
    MsgBox Err.Description
End Sub

Private Sub CmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo errgra
Dim IdEmpl As Integer
If val(Me.txtPorcentaje) < IIf(IsNull(rs3!porc), 0, rs3!porc) Then
    MsgBox "No puede definir un porcentaje menor al ya existente", vbInformation
    Me.txtPorcentaje.Text = IIf(IsNull(rs3!porc), 0, rs3!porc)
    Me.txtPorcentaje.SetFocus
    Exit Sub
End If
If MsgBox("Esta seguro de guardar los cambios", vbYesNo) = vbYes Then
    IdEmpl = rs3!CodEmpleado
    sql = "UPDATE [dbo].[Empleado] SET [SalPorcentaje] = " & val(Me.txtPorcentaje) & ", [SueldoPeriodo] = " & val(Me.txtTotal) & _
            " Where [CodEmpleado] = " & rs3!CodEmpleado
            
    cnx.Execute sql
    Call tvNominas_NodeClick(tvNominas.SelectedItem)
    rs3.Find "codempleado = " & IdEmpl
    MsgBox "Actualización completada", vbInformation

End If
Exit Sub
errgra:
    MsgBox Err.Description
End Sub

Private Sub Form_Activate()
On Error Resume Next
MDIPrimero.Skin1.ApplySkin hWnd
End Sub

Private Sub Form_Load()
On Error GoTo errload
Dim rs2 As New ADODB.Recordset

Me.Top = (MDIPrimero.ScaleHeight / 2) - (Me.Height / 2)
Me.Left = (MDIPrimero.ScaleWidth / 2) - (Me.Width / 2)

dtpSolicitud.Value = Now

If cnx.State = adStateClosed Then
'    sql = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PRUEBA;Data Source=WEBMASTER\SQL2005"
    cnx.ConnectionString = Conexion
    cnx.Open
End If

'NOMINAS
sql = "SELECT  [Nomina], TN.[CodTipoNomina], N.FECHANOMINAINI, N.FECHANOMINA From [dbo].[TipoNomina] TN inner join Nomina N ON TN.CODTIPONOMINA = N.CODTIPONOMINA Where TN.[Activa] = 1 AND N.ACTIVA = 1 "
With rs2
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

If Not rs2.eof Then
    rs2.MoveFirst
    Do While Not rs2.eof
        tvNominas.Nodes.Add , tvwLast, "A" & Trim(rs2!CodTipoNomina), Trim(rs2!Nomina) & ": " & Format$(rs2!fechanominaini, "dd/mm/yyyy") & " - " & Format$(rs2!FechaNomina, "dd/mm/yyyy"), 1, 2
        rs2.MoveNext
    Loop
End If

If tvNominas.Nodes.Count > 0 Then tvNominas.Nodes(1).Selected = True
    
'VERIFICO LA EXISTENCIA DE LOS PUNTOS POR ANTIGUEDAD
sql = "SELECT * From PUNTOSGRUPO WHERE GRUPO = upper('antigüedad')"
With rs2
    If .State = adStateOpen Then .Close
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

If rs2.BOF = rs2.eof And rs2.eof = True Then
    If MsgBox("Los puntos por antigüedad no existen. Sino los crea, la antigüedad no podrá ser calculada" & vbCr & _
                "¿Desea crearlos?", vbYesNo) = vbYes Then
        sql = "INSERT INTO PUNTOSGRUPO(GRUPO) VALUES('ANTIGÜEDAD')"
        cnx.Execute sql
        Dim X As Integer
        For X = 1 To 20
            sql = "INSERT INTO PUNTOS " & _
                "SELECT (SELECT ID FROM PUNTOSGRUPO WHERE GRUPO = upper('antigüedad')), 'SEMESTRE Nº " & X & IIf(X = 1, "', 3", "', 1")
            cnx.Execute sql
        Next
    End If
End If

'PUNTOS POR GRUPO E INDIVIDUALES
sql = "SELECT * FROM PUNTOSGRUPO ORDER BY GRUPO"
        
With rs2
    If .State = adStateOpen Then .Close
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

If Not rs2.eof Then
    rs2.MoveFirst
    Do While Not rs2.eof
        tvActividades.Nodes.Add , tvwLast, "A" & Trim(Str(rs2!Id)), Trim(rs2!grupo), 1, 2
        rs2.MoveNext
    Loop
End If
Set rs2 = Nothing

sql = "SELECT P.Id, G.GRUPO, P.DESCRIPCION, P.CANTPTS, G.ID AS IDG " & _
            "FROM PUNTOS P INNER JOIN PUNTOSGRUPO G ON P.GRUPO = G.ID " & _
            "ORDER BY P.DESCRIPCION"
With rs
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

If Not rs.eof Then
    rs.MoveFirst
    Do While Not rs.eof
        tvActividades.Nodes.Add "A" & Trim(Str(rs!idg)), tvwChild, "B" & Trim(Str(rs!Id)), Trim(rs!descripcion), 1, 2
        rs.MoveNext
    Loop
End If

If tvActividades.Nodes.Count > 0 Then tvActividades.Nodes(1).Selected = True
Call tvNominas_NodeClick(Me.tvNominas.SelectedItem)

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

Private Sub tdbgEmpleados_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo errsem

If tvActividades.Nodes.Count = 0 Then
    MsgBox "Agregue actividades para continuar con las operaciones", vbInformation: Exit Sub
ElseIf tvNominas.Nodes.Count = 0 Then
    MsgBox "Agregue o active las nominas para continuar con las operaciones", vbInformation: Exit Sub
End If

If rs3.BOF = rs3.eof And rs3.eof = False Then
    fraSolicitud.Enabled = True
    fracomplementos.Enabled = True
Else
    fraSolicitud.Enabled = False
    fracomplementos.Enabled = False
    Me.tdbgPuntos.ClearFields
    txtCantPts = ""
    Me.txtPorcentaje = ""
    Me.txtSolicitados = ""
    Me.txtValPorcentaje = ""
    Me.txtValPts = ""
    Me.txtTotal = ""
    Exit Sub
End If

sql = "SELECT  P.ID, G.GRUPO, P.DESCRIPCION, P.CANTPTS AS PUNTOS, EP.APROBADO, E.CODEMPLEADO, EP.JUSTIFICACION, " & _
        "EP.DOCUMENTO, EP.DIRDOCUMENTO, EP.FECHASOLICITUD, EP.FECHAAPROBADO " & _
        "FROM    EMPLEADO E INNER JOIN PUNTOSEMPLEADO EP ON E.CODEMPLEADO = EP.EMPLEADO " & _
        "INNER JOIN PUNTOS P ON EP.PUNTOS = P.ID " & _
        "INNER JOIN PUNTOSGRUPO G ON P.GRUPO = G.ID " & _
        "WHERE  E.CODEMPLEADO = " & rs3!CodEmpleado & " " & _
        "ORDER BY EP.APROBADO DESC, G.GRUPO, P.DESCRIPCION"

With rs4
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

Me.tdbgPuntos.DataSource = rs4

Dim X As Integer
For X = 5 To Me.tdbgPuntos.Columns.Count - 1
    Me.tdbgPuntos.Columns(X).Visible = False
Next
Me.tdbgPuntos.Columns(0).Width = 500
Me.tdbgPuntos.Columns(4).ValueItems.Presentation = dbgCheckBox

Me.txtCantPts = "0"
Me.txtSolicitados = "0"
If Not rs4.eof Then
    rs4.MoveFirst
    Do While Not rs4.eof
        If rs4!aprobado Then Me.txtCantPts.Text = "" & val(Me.txtCantPts.Text) + rs4!puntos
        If Not rs4!aprobado Then Me.txtSolicitados.Text = "" & val(Me.txtSolicitados.Text) + rs4!puntos
        rs4.MoveNext
    Loop
    rs4.MoveFirst
End If

Me.txtPorcentaje.Text = IIf(IsNull(rs3!porc), 0, rs3!porc)
Call txtPorcentaje_Change

Exit Sub
errsem:
    MsgBox Err.Description
End Sub

Private Sub tdbgPuntos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If rs4.State = adStateClosed Then Exit Sub
If rs4.BOF = rs4.eof And rs4.eof = True Then Exit Sub
If rs4!aprobado Then Me.cmdBorrarCta.Enabled = False
If Not rs4!aprobado Then Me.cmdBorrarCta.Enabled = True
End Sub

Private Sub tvNominas_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo errNode
Dim rs2 As New ADODB.Recordset

sql = "SELECT CodEmpleado1 AS Número, 'Empleado' = " & _
                "   case when Nombre2 is null then Nombre1 + ' ' + Apellido1 + ' ' + Apellido2 " & _
                "   else Nombre1 + ' ' + Nombre2 + ' ' + Apellido1 + ' ' + Apellido2 " & _
                "end, SalPorcentaje AS Porc, CantPts as Pts, SueldoPeriodo as Salario, E.CodEmpleado, h.fechacontrato, fechaantiguedad  " & _
        "FROM EMPLEADO E INNER JOIN TIPONOMINA N ON E.CODTIPONOMINA = N.CODTIPONOMINA inner join historico h on E.codempleado = h.codempleado " & _
        "WHERE N.CODTIPONOMINA = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "'"
        
With rs3
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

Me.tdbgEmpleados.DataSource = rs3
Me.tdbgEmpleados.Columns(0).Width = 1000
Me.tdbgEmpleados.Columns(1).Width = 3000
Me.tdbgEmpleados.Columns(2).Width = 500
Me.tdbgEmpleados.Columns(3).Width = 500
Me.tdbgEmpleados.Columns(4).Width = 900
Me.tdbgEmpleados.Columns(5).Visible = False
Me.tdbgEmpleados.Columns(6).Visible = False
Me.tdbgEmpleados.Columns(7).Visible = False

Dim fecha As Date
fecha = CDate(Mid(tvNominas.SelectedItem.Text, Len(tvNominas.SelectedItem.Text) - 22, 10))
If fecha > Me.dtpSolicitud.MaxDate Then Me.dtpSolicitud.MaxDate = fecha + 1

Me.dtpSolicitud.MinDate = CDate(Mid(tvNominas.SelectedItem.Text, Len(tvNominas.SelectedItem.Text) - 22, 10))
Me.dtpSolicitud.MaxDate = CDate(Right(tvNominas.SelectedItem.Text, 10))
Me.dtpSolicitud.Value = Me.dtpSolicitud.MinDate

'DATOS DEL SALARIO MINIMO Y EL VALOR X PUNTOS
sql = "SELECT [SalarioMinimo], [ValorPts] FROM [dbo].[DatosEmpresa] where numero = 1 "
With rs2
    If .State = adStateOpen Then .Close
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

Me.txtSalario = IIf(IsNull(rs2!salariominimo), 0, rs2!salariominimo)
Me.txtPrecioPts = IIf(IsNull(rs2!valorpts), 0, rs2!valorpts)

'CALCULO LA ANTIGUEDAD DE TODOS LOS EMPLEADOS DE LA NOMINA SELECCIONADA
Call CalcAntiguedad

Exit Sub
errNode:
    MsgBox Err.Description
End Sub

Private Sub CalcAntiguedad()
On Error GoTo errant
Dim rs2 As New ADODB.Recordset

sql = "select * " & _
        "from empleado e  inner join historico h on e.codempleado = h.codempleado " & _
        "WHERE datediff(m, e.fechaantiguedad, '" & Format$(dtpSolicitud.MaxDate, "yyyyMMdd") & "') / 6 > 0 "
With rs2
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

If rs2.BOF = rs2.eof And rs2.eof = True Then
    Set rs2 = Nothing
    Exit Sub
End If

Set rs2 = Nothing

sql = "INSERT INTO [dbo].[PuntosEmpleado]([Empleado], [Puntos], [Aprobado], [Justificacion], [FechaSolicitud], [Eliminar]) " & _
        "select e.codempleado, (select id from puntos where descripcion = 'Semestre Nº ' + cast(e.antiguedad + 1 as varchar)), " & _
        "0, 'PUNTOS POR ANTIGÜEDAD CORRESPONDIENTE AL ' + (select descripcion from puntos where descripcion = 'Semestre Nº ' + cast(e.antiguedad + 1 as varchar)), " & _
        "CAST('" & Format$(dtpSolicitud.MaxDate) & "' AS DATETIME), 1 " & _
        "from empleado e  inner join historico h on e.codempleado = h.codempleado " & _
        "WHERE datediff(m, e.fechaantiguedad, '" & Format$(dtpSolicitud.MaxDate) & "') / 6 > 0"

cnx.Execute sql

sql = "UPDATE EMPLEADO " & _
        "SET FECHAANTIGUEDAD = DATEADD(m, 6, FECHAANTIGUEDAD), ANTIGUEDAD = ANTIGUEDAD + 1 " & _
        "WHERE datediff(m, fechaantiguedad, '" & Format$(dtpSolicitud.MaxDate) & "') / 6 > 0"

cnx.Execute sql

Call CalcAntiguedad

Exit Sub
errant:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub txtPorcentaje_Change()
On Error Resume Next
Me.txtValPorcentaje.Text = "" & Round(val(Me.txtSalario.Text) * val(Me.txtPorcentaje.Text) / 100, 2)
Me.txtValPts.Text = "" & val(Me.txtPrecioPts.Text) * val(Me.txtCantPts.Text)
Me.txtTotal = "" & val(Me.txtSalario.Text) + val(Me.txtValPorcentaje.Text) + val(Me.txtValPts.Text)
End Sub
