VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmProduccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de Produccion"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   15000
   Icon            =   "FrmProduccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   514
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   12600
      TabIndex        =   27
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   2160
      TabIndex        =   26
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "Anterior"
      Height          =   375
      Left            =   360
      TabIndex        =   25
      Top             =   6960
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   120
      ScaleHeight     =   4395
      ScaleWidth      =   14715
      TabIndex        =   13
      Top             =   2400
      Width           =   14775
      Begin TabDlg.SSTab SSTab1 
         Height          =   4335
         Left            =   -240
         TabIndex        =   14
         Top             =   0
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   7646
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Registro Unidades Producidas"
         TabPicture(0)   =   "FrmProduccion.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Registro Horas Trabajadas"
         TabPicture(1)   =   "FrmProduccion.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TDBGridHoras"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame2 
            Height          =   3735
            Left            =   360
            TabIndex        =   15
            Top             =   480
            Width           =   14415
            Begin VB.TextBox Text7 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   10440
               TabIndex        =   22
               Text            =   "1"
               Top             =   120
               Width           =   375
            End
            Begin VB.TextBox Text6 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   9600
               TabIndex        =   21
               Text            =   "1"
               Top             =   120
               Width           =   375
            End
            Begin VB.TextBox Text5 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   8760
               TabIndex        =   20
               Text            =   "1"
               Top             =   120
               Width           =   375
            End
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   7920
               TabIndex        =   19
               Text            =   "1"
               Top             =   120
               Width           =   375
            End
            Begin VB.TextBox Text3 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   6960
               TabIndex        =   18
               Text            =   "1"
               Top             =   120
               Width           =   375
            End
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   6000
               TabIndex        =   17
               Text            =   "1"
               Top             =   120
               Width           =   375
            End
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   5160
               TabIndex        =   16
               Text            =   "1"
               Top             =   120
               Width           =   375
            End
            Begin TrueOleDBGrid70.TDBGrid TDBGProduccion 
               Bindings        =   "FrmProduccion.frx":0342
               Height          =   2655
               Left            =   120
               TabIndex        =   23
               Top             =   480
               Width           =   14175
               _ExtentX        =   25003
               _ExtentY        =   4683
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "CodProc"
               Columns(0).DataField=   "CodProceso"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "CodRef"
               Columns(1).DataField=   "CodReferencia1"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Ref"
               Columns(2).DataField=   "Ref"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "Linea"
               Columns(3).DataField=   "Linea"
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "Precio"
               Columns(4).DataField=   "Precio"
               Columns(4).NumberFormat=   "##,##0.00"
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(5)._VlistStyle=   0
               Columns(5)._MaxComboItems=   5
               Columns(5).Caption=   "Unidad"
               Columns(5).DataField=   "Unidad"
               Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(6)._VlistStyle=   0
               Columns(6)._MaxComboItems=   5
               Columns(6).Caption=   "Lunes"
               Columns(6).DataField=   "Lunes"
               Columns(6).DefaultValue=   "0"
               Columns(6).DefaultValue.vt=   8
               Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(7)._VlistStyle=   0
               Columns(7)._MaxComboItems=   5
               Columns(7).Caption=   "Martes"
               Columns(7).DataField=   "Martes"
               Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(8)._VlistStyle=   0
               Columns(8)._MaxComboItems=   5
               Columns(8).Caption=   "Miercoles"
               Columns(8).DataField=   "Miercoles"
               Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(9)._VlistStyle=   0
               Columns(9)._MaxComboItems=   5
               Columns(9).Caption=   "Jueves"
               Columns(9).DataField=   "Jueves"
               Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(10)._VlistStyle=   0
               Columns(10)._MaxComboItems=   5
               Columns(10).Caption=   "Viernes"
               Columns(10).DataField=   "Viernes"
               Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(11)._VlistStyle=   0
               Columns(11)._MaxComboItems=   5
               Columns(11).Caption=   "Sabado"
               Columns(11).DataField=   "Sabado"
               Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(12)._VlistStyle=   0
               Columns(12)._MaxComboItems=   5
               Columns(12).Caption=   "Domingo"
               Columns(12).DataField=   "Domingo"
               Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(13)._VlistStyle=   0
               Columns(13)._MaxComboItems=   5
               Columns(13).Caption=   "30%"
               Columns(13).DataField=   "Incentivo"
               Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(14)._VlistStyle=   0
               Columns(14)._MaxComboItems=   5
               Columns(14).Caption=   "TotalUnidades"
               Columns(14).DataField=   "TotalUnidades"
               Columns(14).NumberFormat=   "##,##0.00"
               Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(15)._VlistStyle=   0
               Columns(15)._MaxComboItems=   5
               Columns(15).Caption=   "SalarioPieza"
               Columns(15).DataField=   "SalarioPieza"
               Columns(15).NumberFormat=   "##,##0.00"
               Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(16)._VlistStyle=   0
               Columns(16)._MaxComboItems=   5
               Columns(16).Caption=   "CodEmpleado"
               Columns(16).DataField=   "CodEmpleado"
               Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(17)._VlistStyle=   0
               Columns(17)._MaxComboItems=   5
               Columns(17).Caption=   "NumNomina"
               Columns(17).DataField=   "NumNomina"
               Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(18)._VlistStyle=   0
               Columns(18)._MaxComboItems=   5
               Columns(18).Caption=   "Pagado"
               Columns(18).DataField=   "Pagado"
               Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(19)._VlistStyle=   0
               Columns(19)._MaxComboItems=   5
               Columns(19).Caption=   "CodReferencia"
               Columns(19).DataField=   "CodReferencia"
               Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   20
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectorWidth=   688
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   14215660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=20"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=1508"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1429"
               Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=1"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=1508"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1429"
               Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=1"
               Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(11)=   "Column(2).Width=1402"
               Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1323"
               Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8193"
               Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(16)=   "Column(3).Width=1270"
               Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1191"
               Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=1"
               Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(21)=   "Column(4).Width=1402"
               Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1323"
               Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8193"
               Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
               Splits(0)._ColumnProps(26)=   "Column(5).Width=1508"
               Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
               Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1429"
               Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8193"
               Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
               Splits(0)._ColumnProps(31)=   "Column(6).Width=1402"
               Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
               Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1323"
               Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=1"
               Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
               Splits(0)._ColumnProps(36)=   "Column(7).Width=1508"
               Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
               Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1429"
               Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=1"
               Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
               Splits(0)._ColumnProps(41)=   "Column(8).Width=1508"
               Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
               Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1429"
               Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=1"
               Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
               Splits(0)._ColumnProps(46)=   "Column(9).Width=1508"
               Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
               Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=1429"
               Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=1"
               Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
               Splits(0)._ColumnProps(51)=   "Column(10).Width=1508"
               Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
               Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=1429"
               Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=1"
               Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
               Splits(0)._ColumnProps(56)=   "Column(11).Width=1508"
               Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
               Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=1429"
               Splits(0)._ColumnProps(59)=   "Column(11)._ColStyle=1"
               Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
               Splits(0)._ColumnProps(61)=   "Column(12).Width=1508"
               Splits(0)._ColumnProps(62)=   "Column(12).DividerColor=0"
               Splits(0)._ColumnProps(63)=   "Column(12)._WidthInPix=1429"
               Splits(0)._ColumnProps(64)=   "Column(12)._ColStyle=1"
               Splits(0)._ColumnProps(65)=   "Column(12).Order=13"
               Splits(0)._ColumnProps(66)=   "Column(13).Width=1508"
               Splits(0)._ColumnProps(67)=   "Column(13).DividerColor=0"
               Splits(0)._ColumnProps(68)=   "Column(13)._WidthInPix=1429"
               Splits(0)._ColumnProps(69)=   "Column(13)._ColStyle=1"
               Splits(0)._ColumnProps(70)=   "Column(13).Order=14"
               Splits(0)._ColumnProps(71)=   "Column(14).Width=1773"
               Splits(0)._ColumnProps(72)=   "Column(14).DividerColor=0"
               Splits(0)._ColumnProps(73)=   "Column(14)._WidthInPix=1693"
               Splits(0)._ColumnProps(74)=   "Column(14)._ColStyle=8193"
               Splits(0)._ColumnProps(75)=   "Column(14).Order=15"
               Splits(0)._ColumnProps(76)=   "Column(15).Width=1773"
               Splits(0)._ColumnProps(77)=   "Column(15).DividerColor=0"
               Splits(0)._ColumnProps(78)=   "Column(15)._WidthInPix=1693"
               Splits(0)._ColumnProps(79)=   "Column(15)._ColStyle=8193"
               Splits(0)._ColumnProps(80)=   "Column(15).Order=16"
               Splits(0)._ColumnProps(81)=   "Column(16).Width=2725"
               Splits(0)._ColumnProps(82)=   "Column(16).DividerColor=0"
               Splits(0)._ColumnProps(83)=   "Column(16)._WidthInPix=2646"
               Splits(0)._ColumnProps(84)=   "Column(16).Visible=0"
               Splits(0)._ColumnProps(85)=   "Column(16).Order=17"
               Splits(0)._ColumnProps(86)=   "Column(17).Width=2725"
               Splits(0)._ColumnProps(87)=   "Column(17).DividerColor=0"
               Splits(0)._ColumnProps(88)=   "Column(17)._WidthInPix=2646"
               Splits(0)._ColumnProps(89)=   "Column(17).Visible=0"
               Splits(0)._ColumnProps(90)=   "Column(17).Order=18"
               Splits(0)._ColumnProps(91)=   "Column(18).Width=2725"
               Splits(0)._ColumnProps(92)=   "Column(18).DividerColor=0"
               Splits(0)._ColumnProps(93)=   "Column(18)._WidthInPix=2646"
               Splits(0)._ColumnProps(94)=   "Column(18).Visible=0"
               Splits(0)._ColumnProps(95)=   "Column(18).Order=19"
               Splits(0)._ColumnProps(96)=   "Column(19).Width=2725"
               Splits(0)._ColumnProps(97)=   "Column(19).DividerColor=0"
               Splits(0)._ColumnProps(98)=   "Column(19)._WidthInPix=2646"
               Splits(0)._ColumnProps(99)=   "Column(19).Visible=0"
               Splits(0)._ColumnProps(100)=   "Column(19).Order=20"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               AllowDelete     =   -1  'True
               AllowAddNew     =   -1  'True
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               MultipleLines   =   0
               CellTipsWidth   =   0
               DeadAreaBackColor=   -2147483637
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
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HC0C0FF&,.bold=0"
               _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=82,.parent=13,.alignment=2,.locked=0"
               _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=79,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=80,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=81,.parent=17"
               _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=78,.parent=13,.alignment=2,.locked=-1"
               _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=75,.parent=14"
               _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=76,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=77,.parent=17"
               _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=2"
               _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
               _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
               _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
               _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=74,.parent=13,.alignment=2,.locked=-1"
               _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=14"
               _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=15"
               _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=17"
               _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=70,.parent=13,.alignment=2,.locked=-1"
               _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
               _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
               _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
               _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=2"
               _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
               _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
               _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
               _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=2"
               _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
               _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
               _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
               _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=58,.parent=13,.alignment=2"
               _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
               _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
               _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
               _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13,.alignment=2"
               _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
               _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
               _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
               _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=50,.parent=13,.alignment=2"
               _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=47,.parent=14"
               _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=48,.parent=15"
               _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=49,.parent=17"
               _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=46,.parent=13,.alignment=2"
               _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=43,.parent=14"
               _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=44,.parent=15"
               _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=45,.parent=17"
               _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=114,.parent=13,.alignment=2"
               _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=111,.parent=14,.bold=0,.fontsize=825"
               _StyleDefs(86)  =   ":id=111,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(87)  =   ":id=111,.fontname=MS Sans Serif"
               _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=112,.parent=15"
               _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=113,.parent=17"
               _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=86,.parent=13,.alignment=2"
               _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=14"
               _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=15"
               _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=17"
               _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=110,.parent=13,.alignment=2,.locked=-1"
               _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=107,.parent=14"
               _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=108,.parent=15"
               _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=109,.parent=17"
               _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=106,.parent=13,.alignment=2,.locked=-1"
               _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=103,.parent=14"
               _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=104,.parent=15"
               _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=105,.parent=17"
               _StyleDefs(102) =   "Splits(0).Columns(16).Style:id=102,.parent=13"
               _StyleDefs(103) =   "Splits(0).Columns(16).HeadingStyle:id=99,.parent=14"
               _StyleDefs(104) =   "Splits(0).Columns(16).FooterStyle:id=100,.parent=15"
               _StyleDefs(105) =   "Splits(0).Columns(16).EditorStyle:id=101,.parent=17"
               _StyleDefs(106) =   "Splits(0).Columns(17).Style:id=98,.parent=13"
               _StyleDefs(107) =   "Splits(0).Columns(17).HeadingStyle:id=95,.parent=14"
               _StyleDefs(108) =   "Splits(0).Columns(17).FooterStyle:id=96,.parent=15"
               _StyleDefs(109) =   "Splits(0).Columns(17).EditorStyle:id=97,.parent=17"
               _StyleDefs(110) =   "Splits(0).Columns(18).Style:id=94,.parent=13"
               _StyleDefs(111) =   "Splits(0).Columns(18).HeadingStyle:id=91,.parent=14"
               _StyleDefs(112) =   "Splits(0).Columns(18).FooterStyle:id=92,.parent=15"
               _StyleDefs(113) =   "Splits(0).Columns(18).EditorStyle:id=93,.parent=17"
               _StyleDefs(114) =   "Splits(0).Columns(19).Style:id=90,.parent=13"
               _StyleDefs(115) =   "Splits(0).Columns(19).HeadingStyle:id=87,.parent=14"
               _StyleDefs(116) =   "Splits(0).Columns(19).FooterStyle:id=88,.parent=15"
               _StyleDefs(117) =   "Splits(0).Columns(19).EditorStyle:id=89,.parent=17"
               _StyleDefs(118) =   "Named:id=33:Normal"
               _StyleDefs(119) =   ":id=33,.parent=0"
               _StyleDefs(120) =   "Named:id=34:Heading"
               _StyleDefs(121) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(122) =   ":id=34,.wraptext=-1"
               _StyleDefs(123) =   "Named:id=35:Footing"
               _StyleDefs(124) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(125) =   "Named:id=36:Selected"
               _StyleDefs(126) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(127) =   "Named:id=37:Caption"
               _StyleDefs(128) =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(129) =   "Named:id=38:HighlightRow"
               _StyleDefs(130) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(131) =   "Named:id=39:EvenRow"
               _StyleDefs(132) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(133) =   "Named:id=40:OddRow"
               _StyleDefs(134) =   ":id=40,.parent=33"
               _StyleDefs(135) =   "Named:id=41:RecordSelector"
               _StyleDefs(136) =   ":id=41,.parent=34"
               _StyleDefs(137) =   "Named:id=42:FilterBar"
               _StyleDefs(138) =   ":id=42,.parent=33"
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGridHoras 
            Bindings        =   "FrmProduccion.frx":0365
            Height          =   3255
            Left            =   -74400
            TabIndex        =   24
            Top             =   600
            Width           =   13215
            _ExtentX        =   23310
            _ExtentY        =   5741
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "CodEmpleado"
            Columns(0).DataField=   "CodEmpleado"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "NumNomina"
            Columns(1).DataField=   "NumNomina"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "NumLinea"
            Columns(2).DataField=   "NumLinea"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Lunes"
            Columns(3).DataField=   "Lunes"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Martes"
            Columns(4).DataField=   "Martes"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Miercoles"
            Columns(5).DataField=   "Miercoles"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Jueves"
            Columns(6).DataField=   "Jueves"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Viernes"
            Columns(7).DataField=   "Viernes"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Sabado"
            Columns(8).DataField=   "Sabado"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Domingo"
            Columns(9).DataField=   "Domingo"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "TotalHoras"
            Columns(10).DataField=   "TotalHoras"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "SalarioHora"
            Columns(11).DataField=   "SalarioHora"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "TotalSalarioHora"
            Columns(12).DataField=   "TotalSalarioHora"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "Pagado"
            Columns(13).DataField=   "Pagado"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   14
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=14"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2117"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2037"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=1"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=2117"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2037"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=1"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=2117"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2037"
            Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=1"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=2117"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2037"
            Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=1"
            Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(36)=   "Column(7).Width=2117"
            Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2037"
            Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=1"
            Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(41)=   "Column(8).Width=2117"
            Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2037"
            Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=1"
            Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(46)=   "Column(9).Width=2117"
            Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2037"
            Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=1"
            Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(51)=   "Column(10).Width=2117"
            Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=2037"
            Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=8193"
            Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(56)=   "Column(11).Width=2646"
            Splits(0)._ColumnProps(57)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(11)._WidthInPix=2566"
            Splits(0)._ColumnProps(59)=   "Column(11)._ColStyle=2"
            Splits(0)._ColumnProps(60)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(61)=   "Column(12).Width=2646"
            Splits(0)._ColumnProps(62)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(63)=   "Column(12)._WidthInPix=2566"
            Splits(0)._ColumnProps(64)=   "Column(12)._ColStyle=8194"
            Splits(0)._ColumnProps(65)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(66)=   "Column(13).Width=2725"
            Splits(0)._ColumnProps(67)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(68)=   "Column(13)._WidthInPix=2646"
            Splits(0)._ColumnProps(69)=   "Column(13).Visible=0"
            Splits(0)._ColumnProps(70)=   "Column(13).Order=14"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowDelete     =   -1  'True
            AllowAddNew     =   -1  'True
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   -2147483637
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
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HC0C0FF&,.bold=0"
            _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=98,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=95,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=96,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=97,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=94,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=91,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=92,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=93,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=90,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=87,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=88,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=89,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=86,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=83,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=84,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=85,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=82,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=79,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=80,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=81,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=78,.parent=13,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=75,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=76,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=77,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=74,.parent=13,.alignment=2"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=62,.parent=13,.alignment=2,.locked=-1"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=59,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=60,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=61,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=55,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=56,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=57,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=54,.parent=13,.alignment=1,.locked=-1"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=51,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=52,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=53,.parent=17"
            _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=50,.parent=13"
            _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=47,.parent=14"
            _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=48,.parent=15"
            _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=49,.parent=17"
            _StyleDefs(92)  =   "Named:id=33:Normal"
            _StyleDefs(93)  =   ":id=33,.parent=0"
            _StyleDefs(94)  =   "Named:id=34:Heading"
            _StyleDefs(95)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(96)  =   ":id=34,.wraptext=-1"
            _StyleDefs(97)  =   "Named:id=35:Footing"
            _StyleDefs(98)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(99)  =   "Named:id=36:Selected"
            _StyleDefs(100) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(101) =   "Named:id=37:Caption"
            _StyleDefs(102) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(103) =   "Named:id=38:HighlightRow"
            _StyleDefs(104) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(105) =   "Named:id=39:EvenRow"
            _StyleDefs(106) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(107) =   "Named:id=40:OddRow"
            _StyleDefs(108) =   ":id=40,.parent=33"
            _StyleDefs(109) =   "Named:id=41:RecordSelector"
            _StyleDefs(110) =   ":id=41,.parent=34"
            _StyleDefs(111) =   "Named:id=42:FilterBar"
            _StyleDefs(112) =   ":id=42,.parent=33"
         End
      End
   End
   Begin MSAdodcLib.Adodc DtaDetalleHora 
      Height          =   330
      Left            =   7440
      Top             =   8160
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
      Caption         =   "DtaDetalleHora"
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
      Left            =   960
      Top             =   8520
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
   Begin MSAdodcLib.Adodc DtaEmpleados 
      Height          =   330
      Left            =   4080
      Top             =   8520
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
   Begin MSAdodcLib.Adodc DtaDetalleProduccion 
      Height          =   330
      Left            =   7440
      Top             =   8520
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
      Caption         =   "DtaDetalleProduccion"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   14655
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   13560
         OleObjectBlob   =   "FrmProduccion.frx":0382
         TabIndex        =   35
         Top             =   480
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   12120
         OleObjectBlob   =   "FrmProduccion.frx":03E6
         TabIndex        =   34
         Top             =   480
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   10680
         OleObjectBlob   =   "FrmProduccion.frx":044C
         TabIndex        =   33
         Top             =   480
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   7440
         OleObjectBlob   =   "FrmProduccion.frx":04BA
         TabIndex        =   32
         Top             =   840
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmProduccion.frx":0526
         TabIndex        =   31
         Top             =   840
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   3000
         OleObjectBlob   =   "FrmProduccion.frx":0592
         TabIndex        =   30
         Top             =   840
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   2640
         OleObjectBlob   =   "FrmProduccion.frx":0602
         TabIndex        =   29
         Top             =   240
         Width           =   6135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   200
         Left            =   240
         OleObjectBlob   =   "FrmProduccion.frx":0686
         TabIndex        =   28
         Top             =   240
         Width           =   1905
      End
      Begin VB.TextBox TxtCodEmpleado 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TxtPeriodo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8040
         TabIndex        =   10
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox TxtTipoNomina 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5520
         TabIndex        =   9
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox TxtCodTipoNomina 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   8
         Top             =   840
         Width           =   735
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
         Left            =   2160
         Picture         =   "FrmProduccion.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TxtDesde 
         Enabled         =   0   'False
         Height          =   285
         Left            =   10800
         TabIndex        =   6
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox TxtMes 
         Enabled         =   0   'False
         Height          =   285
         Left            =   13920
         TabIndex        =   5
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TxtAño 
         Enabled         =   0   'False
         Height          =   285
         Left            =   12720
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox TxtNperiodo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   11400
         TabIndex        =   3
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox TxtNombres 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   2
         Top             =   480
         Width           =   7935
      End
      Begin MSDataListLib.DataCombo DBCodEmpleado 
         Bindings        =   "FrmProduccion.frx":0850
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "CodEmpleado1"
         Text            =   ""
      End
   End
   Begin VB.PictureBox xp_canvas1 
      Height          =   4335
      Left            =   2520
      ScaleHeight     =   4275
      ScaleWidth      =   14115
      TabIndex        =   11
      Top             =   9720
      Width           =   14175
   End
   Begin Threed.SSCommand CmdAcercade 
      Height          =   555
      Left            =   360
      TabIndex        =   36
      Top             =   120
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   979
      _Version        =   196610
      Font3D          =   2
      MarqueeStyle    =   4
      ForeColor       =   8388608
      MarqueeDelay    =   5
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "MOVIMIENTOS DE PRODUCCION"
      ButtonStyle     =   4
      AutoRepeat      =   -1  'True
   End
End
Attribute VB_Name = "FrmProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub CmdAnterior_Click()
 Me.DtaEmpleados.Recordset.MovePrevious
 If Me.DtaEmpleados.Recordset.BOF Then
  MsgBox "Este es el Primer Registro", vbExclamation, "Sistema de Nominas"
  Me.DtaEmpleados.Recordset.MoveNext
 Else
  Me.DBCodEmpleado.Text = Me.DtaEmpleados.Recordset("CodEmpleado1")
 End If
End Sub

Private Sub CmdBuscarEmpleado_Click()
Quien = "Produccion"
FrmBuscaEmpleado.Show 1
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSiguiente_Click()
 Me.DtaEmpleados.Recordset.MoveNext
 If Me.DtaEmpleados.Recordset.EOF Then
  MsgBox "Este es el Ultimo Registro", vbExclamation, "Sistema de Nominas"
  Me.DtaEmpleados.Recordset.MovePrevious
 Else
  Me.DBCodEmpleado.Text = Me.DtaEmpleados.Recordset("CodEmpleado1")
 End If
End Sub

Private Sub DBCodEmpleado_Change()
 Dim Lunes As Double, Martes As Double, Miercoles As Double, Jueves As Double, Viernes As Double
 Dim Sabado As Double, Domingo As Double, Incentivo As Double
 Dim NumNomina As Integer
 

 Me.DtaConsulta.RecordSource = "SELECT CodEmpleado, CodEmpleado1, Activo, Liquidado From Empleado WHERE (Activo = 1) AND (CodEmpleado1 = '" & Me.DBCodEmpleado.Text & "')"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
   Me.TxtCodEmpleado.Text = Me.DtaConsulta.Recordset("CodEmpleado")
 End If
 Me.DtaConsulta.RecordSource = "SELECT Empleado.CodEmpleado,Empleado.CodEmpleado1, Empleado.Nombre1 + '  ' + Empleado.Nombre2 + '  ' + Empleado.Apellido1 + '  ' + Empleado.Apellido2 AS Nombres,Empleado.CodTipoNomina , TipoNomina.Nomina, TipoNomina.Periodo FROM Empleado INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina WHERE (Empleado.CodEmpleado = " & val(Me.TxtCodEmpleado) & ")"
 Me.DtaConsulta.Refresh
 If Not Me.DtaConsulta.Recordset.EOF Then
   Me.TxtNombres.Text = Me.DtaConsulta.Recordset("Nombres")
   Me.TxtCodTipoNomina.Text = Me.DtaConsulta.Recordset("CodTipoNomina")
   Me.txtPeriodo.Text = Me.DtaConsulta.Recordset("Periodo")
   Me.TxtTipoNomina.Text = Me.DtaConsulta.Recordset("Nomina")
   Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Actual, Inicio, Final From Fecha_Planilla WHERE (Actual = 1) AND (CodTipoNomina = '" & Me.TxtCodTipoNomina.Text & "')"
    Me.CmdAcercade.Caption = "Empleado: " & Me.DBCodEmpleado & " " & Me.DtaConsulta.Recordset("Nombres")
'   InputBox "", "", Me.DtaConsulta.RecordSource
   Me.DtaConsulta.Refresh
   If Not DtaConsulta.Recordset.EOF Then
     Me.TxtAño.Text = Me.DtaConsulta.Recordset("año")
     Me.txtMes.Text = Me.DtaConsulta.Recordset("mes")
     Me.TxtNperiodo.Text = Me.DtaConsulta.Recordset("Periodo")
     Me.txtDesde.Text = "Desde " & Me.DtaConsulta.Recordset("Inicio") & "  Al  " & Me.DtaConsulta.Recordset("Final")
   End If
   Me.DtaConsulta.RecordSource = "SELECT Nomina.*, TipoNomina.TipoPago FROM TipoNomina INNER JOIN Nomina ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina Where (((Nomina.Activa) = 1)) And Nomina.CodTipoNomina = '" & Me.TxtCodTipoNomina.Text & "' "
   'Me.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, Periodo, Activa From TipoNomina WHERE     (CodTipoNomina = '" & Me.TxtCodTipoNomina.Text & "') AND (Activa = 1) "
   Me.DtaConsulta.Refresh
   If DtaConsulta.Recordset.EOF Then
     MsgBox "No Exste Nomina Activa para Este Empleado", vbCritical, "Sistema de Nominas"
     Me.DBCodEmpleado.Text = ""
     Me.TxtCodEmpleado.Text = ""
     Exit Sub
   Else
     NumNomina = Me.DtaConsulta.Recordset("NumNomina")
   End If
   
  '/////////////////////////////////////////////////////////////////////////////////////////////////
  '////////INICIO CON LAS UNIDADES PRODUCIDAS///////////////////////////////////////////////
  '///////////////////////////////////////////////////////////////////////////////////////////
   

'Me.DtaDetalleProduccion.RecordSource = "SELECT DetalleProduccion.CodEmpleado, DetalleProduccion.NumNomina, DetalleProduccion.CodReferencia, Referencia.CodReferencia1, DetalleProduccion.CodProceso, DetalleProduccion.Ref, DetalleProduccion.Linea, DetalleProduccion.Lunes, DetalleProduccion.Martes, DetalleProduccion.Miercoles, DetalleProduccion.Jueves, DetalleProduccion.Viernes, DetalleProduccion.Sabado, DetalleProduccion.Domingo,DetalleProduccion.Incentivo, DetalleProduccion.TotalUnidades, DetalleProduccion.SalarioPieza, DetalleProduccion.Precio, DetalleProduccion.Unidad,DetalleProduccion.Pagado , Nomina.Activa " & _
'                                       "FROM  DetalleProduccion INNER JOIN Nomina ON DetalleProduccion.NumNomina = Nomina.NumNomina INNER JOIN Referencia ON DetalleProduccion.CodReferencia = Referencia.CodReferencia " & _
'                                       "WHERE  (DetalleProduccion.CodEmpleado = '" & Me.TxtCodEmpleado & "') And (DetalleProduccion.NumNomina = " & NumNomina & ") And (Nomina.Activa = 1)"
       Me.DtaDetalleProduccion.RecordSource = "SELECT DetalleProduccion.CodEmpleado, DetalleProduccion.NumNomina, DetalleProduccion.CodReferencia, DetalleProduccion.CodReferencia1, DetalleProduccion.CodProceso,DetalleProduccion.Ref, DetalleProduccion.Linea,DetalleProduccion.Lunes, DetalleProduccion.Martes, DetalleProduccion.Miercoles, DetalleProduccion.Jueves, DetalleProduccion.Viernes, DetalleProduccion.Sabado, DetalleProduccion.Domingo, DetalleProduccion.Incentivo,DetalleProduccion.TotalUnidades, DetalleProduccion.SalarioPieza,DetalleProduccion.Precio , DetalleProduccion.unidad, DetalleProduccion.Pagado, Nomina.Activa FROM DetalleProduccion INNER JOIN Nomina ON DetalleProduccion.NumNomina = Nomina.NumNomina Where (DetalleProduccion.CodEmpleado = '" & Me.TxtCodEmpleado & "') And (DetalleProduccion.NumNomina = " & NumNomina & ") And (Nomina.Activa = 1)"

        Me.DtaDetalleProduccion.Refresh
Me.TDBGProduccion.Columns(0).Button = True


 If Not Me.TDBGProduccion.Columns(4).Text = "" Then
  Precio = Me.TDBGProduccion.Columns(4).Text
 Else
  Precio = 0
 End If
 
 If Not Me.TDBGProduccion.Columns(5).Text = "" Then
 unidad = Me.TDBGProduccion.Columns(5).Text
 Else
  unidad = 0
 End If
 
 If Not Me.TDBGProduccion.Columns(6).Text = "" Then
  Lunes = Me.TDBGProduccion.Columns(6).Text
 Else
  Lunes = 0
 End If
 If Not Me.TDBGProduccion.Columns(7).Text = "" Then
   Martes = val(Me.TDBGProduccion.Columns(7).Text)
 Else
   Martes = 0
 End If
 If Not Me.TDBGProduccion.Columns(8).Text = "" Then
  Miercoles = val(Me.TDBGProduccion.Columns(8).Text)
 Else
  Miercoles = 0
 End If
 
 If Not Me.TDBGProduccion.Columns(9).Text = "" Then
 Jueves = val(Me.TDBGProduccion.Columns(9).Text)
 Else
  Jueves = 0
 End If
 If Not Me.TDBGProduccion.Columns(10).Text = "" Then
  Viernes = val(Me.TDBGProduccion.Columns(10).Text)
 Else
  Viernes = 0
 End If
 If Not Me.TDBGProduccion.Columns(11).Text = "" Then
  Sabado = val(Me.TDBGProduccion.Columns(11).Text)
 Else
  Sabado = 0
 End If
 If Not Me.TDBGProduccion.Columns(12).Text = "" Then
  Domingo = val(Me.TDBGProduccion.Columns(12).Text)
 Else
  Domingo = 0
 End If
 
If Not Me.TDBGProduccion.Columns(13).Text = "" Then
  Incentivo = val(Me.TDBGProduccion.Columns(13).Text)
 Else
  Incentivo = 0
 End If
 
 TotalUnidad = Lunes + Martes + Miercoles + Jueves + Viernes + Sabado + Domingo + Incentivo
 If Not unidad = 0 Then
  SalPieza = (TotalUnidad / unidad) * Precio
 Else
  SalPieza = 0
 End If
 
 If Not TotalUnidad = 0 Then
    Me.TDBGProduccion.Columns(15).Text = SalPieza
    Me.TDBGProduccion.Columns(14).Text = TotalUnidad
 End If


   If Not Me.DtaDetalleProduccion.Recordset.EOF Then
    Me.TDBGProduccion.Columns(16).Text = Me.TxtCodEmpleado.Text
    Me.TDBGProduccion.Columns(17).Text = NumNomina
   End If
   Me.Frame2.Enabled = True
 
 '//////////////////////////////////////////////////////////////////////////////////////////////
 '/////////////////INICIO CON LOS DATOS POR HORA DE PRODUCCION/////////////////////////////////
 '/////////////////////////////////////////////////////////////////////////////////////////////
 
 
 Me.DtaDetalleHora.RecordSource = "SELECT CodEmpleado, NumNomina, NumLinea, Lunes, Martes, Miercoles, Jueves, Viernes, Sabado, Domingo, TotalHoras, SalarioHora, TotalSalarioHora,Pagado From dbo.DetalleHorasProduccion WHERE (Pagado = 0) AND (CodEmpleado = '" & Me.TxtCodEmpleado & "') AND (NumNomina = " & NumNomina & ")"
' InputBox "", "", Me.DtaDetalleHora.RecordSource
 Me.DtaDetalleHora.Refresh
' Me.TDBGridHoras.Columns(0).Visible = False
' Me.TDBGridHoras.Columns(1).Visible = False
' Me.TDBGridHoras.Columns(2).Visible = False
' Me.TDBGridHoras.Columns(3).Width = 1200
' Me.TDBGridHoras.Columns(4).Width = 1200
' Me.TDBGridHoras.Columns(5).Width = 1200
' Me.TDBGridHoras.Columns(6).Width = 1200
' Me.TDBGridHoras.Columns(7).Width = 1200
' Me.TDBGridHoras.Columns(8).Width = 1200
' Me.TDBGridHoras.Columns(9).Width = 1200
' Me.TDBGridHoras.Columns(10).Width = 1200
' Me.TDBGridHoras.Columns(11).Width = 1500
' Me.TDBGridHoras.Columns(11).Locked = True
' Me.TDBGridHoras.Columns(11).Locked = True
' Me.TDBGridHoras.Columns(12).Width = 1500
' Me.TDBGridHoras.Columns(13).Visible = False
  
  If Not Me.TDBGridHoras.Columns(3).Text = "" Then
   Lunes = Me.TDBGridHoras.Columns(3).Text
  Else
    Lunes = 0
  End If
  If Not Me.TDBGridHoras.Columns(4).Text = "" Then
   Martes = Me.TDBGridHoras.Columns(4).Text
  Else
   Martes = 0
  End If
  If Not Me.TDBGridHoras.Columns(5).Text = "" Then
   Miercoles = Me.TDBGridHoras.Columns(5).Text
  Else
   Miercoles = 0
  End If
  If Not Me.TDBGridHoras.Columns(6).Text = "" Then
   Jueves = Me.TDBGridHoras.Columns(6).Text
  Else
   Jueves = 0
  End If
  If Not Me.TDBGridHoras.Columns(7).Text = "" Then
   Viernes = Me.TDBGridHoras.Columns(7).Text
  Else
   Viernes = 0
  End If
  If Not Me.TDBGridHoras.Columns(8).Text = "" Then
   Sabado = Me.TDBGridHoras.Columns(8).Text
  Else
   Sabado = 0
  End If
  If Not Me.TDBGridHoras.Columns(9).Text = "" Then
   Domingo = Me.TDBGridHoras.Columns(9).Text
  Else
    Domingo = 0
  End If
   
   TotalHoras = Lunes + Martes + Miercoles + Jueves + Viernes + Sabado + Domingo
   If Not TotalHoras = 0 Then
   TotalHoras = Lunes + Martes + Miercoles + Jueves + Viernes + Sabado + Domingo
   Me.TDBGridHoras.Columns(10).Text = TotalHoras
   
'   Me.DtaConsulta.RecordSource = "SELECT CodEmpleado, TarifaHoraria From dbo.Empleado WHERE (CodEmpleado = " & Me.TxtCodEmpleado.Text & " )"
'   Me.DtaConsulta.Refresh
'   If Not DtaConsulta.Recordset.EOF Then
'     TarifaHoraria = Me.DtaConsulta.Recordset("TarifaHoraria")
'   End If
   TarifaHoraria = BuscaTarifaHoraria(Me.TxtCodEmpleado.Text)
   Me.TDBGridHoras.Columns(11).Text = Format(TarifaHoraria, "##,##0.000")
   Me.TDBGridHoras.Columns(12).Text = Format(TarifaHoraria * TotalHoras, "##,##0.00")
   End If
 
 Else
 
 Me.DtaDetalleHora.RecordSource = "SELECT CodEmpleado, NumNomina, NumLinea, Lunes, Martes, Miercoles, Jueves, Viernes, Sabado, Domingo, TotalHoras, SalarioHora, TotalSalarioHora,Pagado From dbo.DetalleHorasProduccion WHERE (Pagado = 0) AND (CodEmpleado = '-1')"
 Me.DtaDetalleHora.Refresh
 Me.TDBGridHoras.Columns(0).Visible = False
 Me.TDBGridHoras.Columns(1).Visible = False
 Me.TDBGridHoras.Columns(2).Visible = False
 Me.TDBGridHoras.Columns(3).Width = 1200
 Me.TDBGridHoras.Columns(4).Width = 1200
 Me.TDBGridHoras.Columns(5).Width = 1200
 Me.TDBGridHoras.Columns(6).Width = 1200
 Me.TDBGridHoras.Columns(7).Width = 1200
 Me.TDBGridHoras.Columns(8).Width = 1200
 Me.TDBGridHoras.Columns(9).Width = 1200
 Me.TDBGridHoras.Columns(10).Width = 1200
 Me.TDBGridHoras.Columns(11).Locked = True
 Me.TDBGridHoras.Columns(11).Locked = True
 Me.TDBGridHoras.Columns(11).Width = 1500
 Me.TDBGridHoras.Columns(12).Width = 1500
' Me.TDBGridHoras.Columns(13).Visible = False
 
 Me.DtaDetalleProduccion.RecordSource = "SELECT CodProceso,CodReferencia,Ref, Linea,Precio, unidad, Lunes, Martes, Miercoles, Jueves, Viernes, Sabado, Domingo,Incentivo, TotalUnidades,SalarioPieza,CodEmpleado, NumNomina From DetalleProduccion WHERE (CodEmpleado = '-1')"
 Me.DtaDetalleProduccion.Refresh
Me.TDBGProduccion.Columns(0).Button = True
Me.TDBGProduccion.Columns(0).Caption = "CodProc"
Me.TDBGProduccion.Columns(0).Width = 850
Me.TDBGProduccion.Columns(1).Caption = "CodRef"
Me.TDBGProduccion.Columns(1).Width = 850
Me.TDBGProduccion.Columns(2).Width = 750
Me.TDBGProduccion.Columns(4).NumberFormat = "##,##0.00"
Me.TDBGProduccion.Columns(4).Width = 850
Me.TDBGProduccion.Columns(5).Caption = "Unidad"
Me.TDBGProduccion.Columns(5).Width = 850
Me.TDBGProduccion.Columns(6).Width = 850
Me.TDBGProduccion.Columns(7).Width = 850
Me.TDBGProduccion.Columns(8).Width = 850
Me.TDBGProduccion.Columns(9).Width = 850
Me.TDBGProduccion.Columns(10).Width = 850
Me.TDBGProduccion.Columns(11).Width = 850
Me.TDBGProduccion.Columns(12).Width = 850
Me.TDBGProduccion.Columns(13).Width = 850
Me.TDBGProduccion.Columns(14).NumberFormat = "##,##0.00"
Me.TDBGProduccion.Columns(14).Width = 1000
Me.TDBGProduccion.Columns(15).NumberFormat = "##,##0.00"
Me.TDBGProduccion.Columns(15).Width = 1000
'Me.TDBGProduccion.Columns(16).Visible = False
'Me.TDBGProduccion.Columns(17).Visible = False

     Me.TxtNombres.Text = ""
     Me.TxtCodTipoNomina.Text = ""
     Me.txtPeriodo.Text = ""
     Me.TxtTipoNomina.Text = ""
     Me.TxtAño.Text = ""
     Me.txtMes.Text = ""
     Me.TxtNperiodo.Text = ""
     Me.txtDesde.Text = ""
   Me.Frame2.Enabled = False
 End If

 
 
 
 
End Sub

Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd
With Me.DtaDetalleProduccion
  .ConnectionString = Conexion
End With

With Me.DtaDetalleHora
  .ConnectionString = Conexion
End With

With Me.DtaEmpleados
  .ConnectionString = Conexion
  .RecordSource = "SELECT Empleado.CodEmpleado,Empleado.CodEmpleado1, TipoNomina.CodTipoNomina, TipoNomina.TipoPago FROM Empleado INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina WHERE (Empleado.Activo = 1)and (TipoNomina.TipoPago = 'Salario Destajo') OR (TipoNomina.TipoPago = 'Salario Fijo,Destajo y Comision') OR (TipoNomina.TipoPago = 'Salario Destajo y Comision') ORDER BY CodEmpleado"
  .Refresh
End With

With Me.DtaConsulta
  .ConnectionString = Conexion
End With

 Me.DtaDetalleHora.RecordSource = "SELECT CodEmpleado, NumNomina, NumLinea, Lunes, Martes, Miercoles, Jueves, Viernes, Sabado, Domingo, TotalHoras, SalarioHora, TotalSalarioHora,Pagado From dbo.DetalleHorasProduccion WHERE (Pagado = 0) AND (CodEmpleado = '-1')"
 Me.DtaDetalleHora.Refresh
' Me.TDBGridHoras.Columns(0).Visible = False
' Me.TDBGridHoras.Columns(1).Visible = False
' Me.TDBGridHoras.Columns(2).Visible = False
' Me.TDBGridHoras.Columns(3).Width = 1200
' Me.TDBGridHoras.Columns(4).Width = 1200
' Me.TDBGridHoras.Columns(5).Width = 1200
' Me.TDBGridHoras.Columns(6).Width = 1200
' Me.TDBGridHoras.Columns(7).Width = 1200
' Me.TDBGridHoras.Columns(8).Width = 1200
' Me.TDBGridHoras.Columns(9).Width = 1200
' Me.TDBGridHoras.Columns(10).Width = 1200
' Me.TDBGridHoras.Columns(11).Locked = True
' Me.TDBGridHoras.Columns(11).Locked = True
' Me.TDBGridHoras.Columns(11).Width = 1500
' Me.TDBGridHoras.Columns(12).Width = 1500
' Me.TDBGridHoras.Columns(13).Visible = False

  



'  Me.DtaDetalleProduccion.RecordSource = "SELECT DetalleProduccion.CodEmpleado, DetalleProduccion.NumNomina, DetalleProduccion.CodReferencia, Referencia.CodReferencia1, DetalleProduccion.CodProceso, DetalleProduccion.Ref, DetalleProduccion.Linea, DetalleProduccion.Lunes, DetalleProduccion.Martes, DetalleProduccion.Miercoles, DetalleProduccion.Jueves, DetalleProduccion.Viernes, DetalleProduccion.Sabado, DetalleProduccion.Domingo,DetalleProduccion.Incentivo, DetalleProduccion.TotalUnidades, DetalleProduccion.SalarioPieza, DetalleProduccion.Precio, DetalleProduccion.Unidad,DetalleProduccion.Pagado , Nomina.Activa " & _
'                                       "FROM  DetalleProduccion INNER JOIN Nomina ON DetalleProduccion.NumNomina = Nomina.NumNomina INNER JOIN Referencia ON DetalleProduccion.CodReferencia = Referencia.CodReferencia " & _
'                                       "WHERE (CodEmpleado = '-1')"
                                       
   Me.DtaDetalleProduccion.RecordSource = "SELECT CodProceso,CodReferencia,CodReferencia1, Ref, Precio, unidad, Lunes, Martes, Miercoles, Jueves, Viernes, Sabado, Domingo,Incentivo, TotalUnidades,SalarioPieza,CodEmpleado, NumNomina From DetalleProduccion WHERE (CodEmpleado = '-1')"
   Me.DtaDetalleProduccion.Refresh
Me.TDBGProduccion.Columns(0).Button = True
'Me.TDBGProduccion.Columns(0).Caption = "CodProc"
'Me.TDBGProduccion.Columns(0).Width = 900
'Me.TDBGProduccion.Columns(1).Caption = "CodRef"
'Me.TDBGProduccion.Columns(1).Width = 900
'Me.TDBGProduccion.Columns(2).Width = 800
'Me.TDBGProduccion.Columns(3).NumberFormat = "##,##0.00"
'Me.TDBGProduccion.Columns(3).Width = 900
'Me.TDBGProduccion.Columns(4).Caption = "Unidad"
'Me.TDBGProduccion.Columns(4).Width = 900
'Me.TDBGProduccion.Columns(5).Width = 900
'Me.TDBGProduccion.Columns(6).Width = 900
'Me.TDBGProduccion.Columns(7).Width = 900
'Me.TDBGProduccion.Columns(8).Width = 900
'Me.TDBGProduccion.Columns(9).Width = 900
'Me.TDBGProduccion.Columns(10).Width = 900
'Me.TDBGProduccion.Columns(11).Width = 900
'Me.TDBGProduccion.Columns(12).NumberFormat = "##,##0.00"
'Me.TDBGProduccion.Columns(12).Width = 1000
'Me.TDBGProduccion.Columns(13).NumberFormat = "##,##0.00"
'Me.TDBGProduccion.Columns(13).Width = 1000
'Me.TDBGProduccion.Columns(14).Visible = False
'Me.TDBGProduccion.Columns(15).Visible = False
'Me.TDBGProduccion.Columns(16).Visible = False
 Me.TDBGProduccion.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.TDBGProduccion.OddRowStyle.BackColor = &H80000005
 Me.TDBGProduccion.AlternatingRowStyle = True
 
  Me.TDBGridHoras.EvenRowStyle.BackColor = &HC0FFFF
 Me.TDBGridHoras.OddRowStyle.BackColor = &H80000005
 Me.TDBGridHoras.AlternatingRowStyle = True
 
' Me.BackColor = RGB(236, 233, 216)
' Me.CmdAnterior.BackColor = RGB(236, 233, 216)
' Me.CmdBuscarEmpleado.BackColor = RGB(236, 233, 216)
' Me.CmdSalir.BackColor = RGB(236, 233, 216)
' Me.CmdSiguiente.BackColor = RGB(236, 233, 216)
' Me.Frame1.BackColor = RGB(236, 233, 216)
' Me.Frame2.BackColor = RGB(236, 233, 216)
 
 
End Sub

Private Sub TDBGProduccion_AfterColEdit(ByVal ColIndex As Integer)
Dim Precio As Double, unidad As Double
 Dim Lunes As Double, Martes As Double, Miercoles As Double, Jueves As Double, Viernes As Double
 Dim Sabado As Double, Domingo As Double, TotalUnidad As Double, Incentivo As Double
 Dim SalPieza As Double, CodReferencia As String, CodProcesos As String
 
 

 
 If Not Me.TDBGProduccion.Columns(0).Text = "" Then
 CodProcesos = Me.TDBGProduccion.Columns(0).Text
 Else
 CodProcesos = 0
 End If
 
 If Not Me.TDBGProduccion.Columns(1).Text = "" Then
  CodReferencia = Me.TDBGProduccion.Columns(1).Text
 Else
   CodReferencia = "N"
 End If
 
If CodProcesos <> 0 And CodReferencia <> "N" Then
    Me.DtaConsulta.RecordSource = "SELECT   Procesos.Ref, Referencia.CodReferencia1, Procesos.CodProceso, Procesos.Descrip, Procesos.Precio, Procesos.Unid, Procesos.CodReferencia " & _
                                  "FROM  Procesos INNER JOIN  Referencia ON Procesos.CodReferencia = Referencia.CodReferencia " & _
                                  "WHERE  (Procesos.CodProceso = '" & CodProcesos & "') AND (Referencia.CodReferencia1 = '" & CodReferencia & "') AND (Referencia.Activo = 1)"
    Me.DtaConsulta.Refresh
    If Not DtaConsulta.Recordset.EOF Then
      Me.TDBGProduccion.Columns(2).Text = Me.DtaConsulta.Recordset("Ref")
      Me.TDBGProduccion.Columns(4).Text = Me.DtaConsulta.Recordset("Precio")
      Me.TDBGProduccion.Columns(5).Text = Me.DtaConsulta.Recordset("Unid")
      Me.TDBGProduccion.Columns(19).Text = Me.DtaConsulta.Recordset("CodReferencia")
    Else
      MsgBox "Este FG no existe o bien esta inactivo", vbCritical, "Sistema de nominas"
      Exit Sub
    End If
 End If
 
 
    Me.DtaConsulta.RecordSource = "SELECT Nomina.*, TipoNomina.TipoPago FROM TipoNomina INNER JOIN Nomina ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina Where (((Nomina.Activa) = 1)) And Nomina.CodTipoNomina = '" & Me.TxtCodTipoNomina.Text & "' "
   'Me.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, Periodo, Activa From TipoNomina WHERE     (CodTipoNomina = '" & Me.TxtCodTipoNomina.Text & "') AND (Activa = 1) "
   Me.DtaConsulta.Refresh
   If DtaConsulta.Recordset.EOF Then
     MsgBox "No Exste Nomina Activa para Este Empleado", vbCritical, "Sistema de Nominas"
     Me.DBCodEmpleado.Text = ""
     Me.TxtCodEmpleado.Text = ""
     Exit Sub
   Else
     NumNomina = Me.DtaConsulta.Recordset("NumNomina")
    Me.TDBGProduccion.Columns(16).Text = Me.TxtCodEmpleado.Text
    Me.TDBGProduccion.Columns(17).Text = NumNomina
   End If

 If Not Me.TDBGProduccion.Columns(4).Text = "" Then
  Precio = Me.TDBGProduccion.Columns(4).Text
 Else
  Precio = 0
 End If
 
 If Not Me.TDBGProduccion.Columns(5).Text = "" Then
 unidad = Me.TDBGProduccion.Columns(5).Text
 Else
  unidad = 0
 End If
 
 If Not Me.TDBGProduccion.Columns(6).Text = "" Then
  Lunes = Me.TDBGProduccion.Columns(6).Text
 Else
  Lunes = 0
 End If
 If Not Me.TDBGProduccion.Columns(7).Text = "" Then
   Martes = val(Me.TDBGProduccion.Columns(7).Text)
 Else
   Martes = 0
 End If
 If Not Me.TDBGProduccion.Columns(8).Text = "" Then
  Miercoles = val(Me.TDBGProduccion.Columns(8).Text)
 Else
  Miercoles = 0
 End If
 
 If Not Me.TDBGProduccion.Columns(9).Text = "" Then
 Jueves = val(Me.TDBGProduccion.Columns(9).Text)
 Else
  Jueves = 0
 End If
 If Not Me.TDBGProduccion.Columns(10).Text = "" Then
  Viernes = val(Me.TDBGProduccion.Columns(10).Text)
 Else
  Viernes = 0
 End If
 If Not Me.TDBGProduccion.Columns(11).Text = "" Then
  Sabado = val(Me.TDBGProduccion.Columns(11).Text)
 Else
  Sabado = 0
 End If
 If Not Me.TDBGProduccion.Columns(12).Text = "" Then
  Domingo = val(Me.TDBGProduccion.Columns(12).Text)
 Else
  Domingo = 0
 End If
 
  If Not Me.TDBGProduccion.Columns(13).Text = "" Then
  Incentivo = val(Me.TDBGProduccion.Columns(13).Text)
 Else
  Incentivo = 0
 End If
 TotalUnidad = Lunes + Martes + Miercoles + Jueves + Viernes + Sabado + Domingo + Incentivo
 If Not unidad = 0 Then
  SalPieza = (TotalUnidad / unidad) * Precio
 Else
  SalPieza = 0
 End If
 Me.TDBGProduccion.Columns(15).Text = SalPieza
 Me.TDBGProduccion.Columns(14).Text = TotalUnidad
 


End Sub

Private Sub TDBGProduccion_AfterColUpdate(ByVal ColIndex As Integer)
Dim Precio As Double, unidad As Double
 Dim Lunes As Double, Martes As Double, Miercoles As Double, Jueves As Double, Viernes As Double
 Dim Sabado As Double, Domingo As Double, TotalUnidad As Double, Incentivo As Double
 Dim SalPieza As Double, CodReferencia As String, CodProcesos As String
 

 
 If Not Me.TDBGProduccion.Columns(0).Text = "" Then
 CodProcesos = Me.TDBGProduccion.Columns(0).Text
 Else
 CodProcesos = 0
 End If
 
 If Not Me.TDBGProduccion.Columns(1).Text = "" Then
  CodReferencia = Me.TDBGProduccion.Columns(1).Text
 Else
   CodReferencia = 0
 End If
 
Me.DtaConsulta.RecordSource = "SELECT   Procesos.Ref, Referencia.CodReferencia1, Procesos.CodProceso, Procesos.Descrip, Procesos.Precio, Procesos.Unid, Procesos.CodReferencia " & _
                              "FROM  Procesos INNER JOIN  Referencia ON Procesos.CodReferencia = Referencia.CodReferencia " & _
                              "WHERE  (Procesos.CodProceso = '" & CodProcesos & "') AND (Referencia.CodReferencia1 = '" & CodReferencia & "') AND (Referencia.Activo = 1)"
' Me.DtaConsulta.RecordSource = "SELECT Ref, CodReferencia,CodReferencia1, CodProceso, Descrip, Precio, Unid From Procesos WHERE (CodProceso = '" & CodProcesos & "') AND (CodReferencia = '" & CodReferencia & "')"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
   Me.TDBGProduccion.Columns(2).Text = Me.DtaConsulta.Recordset("Ref")
   Me.TDBGProduccion.Columns(4).Text = Me.DtaConsulta.Recordset("Precio")
   Me.TDBGProduccion.Columns(5).Text = Me.DtaConsulta.Recordset("Unid")
   Me.TDBGProduccion.Columns(19).Text = Me.DtaConsulta.Recordset("CodReferencia")
 End If
 
 '/////////////////////////////////////////////////////////////////////////////////////////////////
 '////////////CON ESTA CONSULTA BUSCO LA ULTIMA DESCRIPCION/////////////////////////////
 '///////////////////////////////////////////////////////////////////////////////////////////
 Me.DtaConsulta.RecordSource = "SELECT DetalleProduccion.CodEmpleado, DetalleProduccion.NumNomina, DetalleProduccion.CodReferencia, Referencia.CodReferencia1, DetalleProduccion.CodProceso, DetalleProduccion.Ref, DetalleProduccion.Linea, DetalleProduccion.Lunes, DetalleProduccion.Martes, DetalleProduccion.Miercoles, DetalleProduccion.Jueves, DetalleProduccion.Viernes, DetalleProduccion.Sabado, DetalleProduccion.Domingo,DetalleProduccion.Incentivo, DetalleProduccion.TotalUnidades, DetalleProduccion.SalarioPieza, DetalleProduccion.Precio, DetalleProduccion.Unidad,DetalleProduccion.Pagado , Nomina.Activa " & _
                                       "FROM  DetalleProduccion INNER JOIN Nomina ON DetalleProduccion.NumNomina = Nomina.NumNomina INNER JOIN Referencia ON DetalleProduccion.CodReferencia = Referencia.CodReferencia " & _
                                       "WHERE  (DetalleProduccion.CodEmpleado = '" & Me.TxtCodEmpleado & "') And (DetalleProduccion.NumNomina = " & NumNomina & ") And (Nomina.Activa = 1)"
 
 Me.DtaConsulta.Refresh
 
 If Not Me.DtaConsulta.Recordset.EOF Then
  Me.DtaConsulta.Recordset.MoveLast
   If Not IsNull(Me.DtaConsulta.Recordset("Linea")) Then
    Me.TDBGProduccion.Columns(3).Text = Me.DtaConsulta.Recordset("Linea")
   End If
 End If
 
 
    Me.DtaConsulta.RecordSource = "SELECT Nomina.*, TipoNomina.TipoPago FROM TipoNomina INNER JOIN Nomina ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina Where (((Nomina.Activa) = 1)) And Nomina.CodTipoNomina = '" & Me.TxtCodTipoNomina.Text & "' "
   'Me.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, Periodo, Activa From TipoNomina WHERE     (CodTipoNomina = '" & Me.TxtCodTipoNomina.Text & "') AND (Activa = 1) "
   Me.DtaConsulta.Refresh
   If DtaConsulta.Recordset.EOF Then
     MsgBox "No Exste Nomina Activa para Este Empleado", vbCritical, "Sistema de Nominas"
     Me.DBCodEmpleado.Text = ""
     Me.TxtCodEmpleado.Text = ""
     Exit Sub
   Else
     NumNomina = Me.DtaConsulta.Recordset("NumNomina")
    Me.TDBGProduccion.Columns(16).Text = Me.TxtCodEmpleado.Text
    Me.TDBGProduccion.Columns(17).Text = NumNomina
   End If

 If Not Me.TDBGProduccion.Columns(4).Text = "" Then
  Precio = Me.TDBGProduccion.Columns(4).Text
 Else
  Precio = 0
  Me.TDBGProduccion.Columns(4).Text = 0
 End If
 
 If Not Me.TDBGProduccion.Columns(5).Text = "" Then
 unidad = Me.TDBGProduccion.Columns(5).Text
 Else
  unidad = 0
  Me.TDBGProduccion.Columns(5).Text = 0
 End If
 
 If Not Me.TDBGProduccion.Columns(6).Text = "" Then
  Lunes = Me.TDBGProduccion.Columns(6).Text
 Else
  Lunes = 0
  Me.TDBGProduccion.Columns(6).Text = 0
 End If
 If Not Me.TDBGProduccion.Columns(7).Text = "" Then
   Martes = val(Me.TDBGProduccion.Columns(7).Text)
 Else
   Martes = 0
   Me.TDBGProduccion.Columns(7).Text = 0
 End If
 If Not Me.TDBGProduccion.Columns(8).Text = "" Then
  Miercoles = val(Me.TDBGProduccion.Columns(8).Text)
 Else
  Miercoles = 0
  Me.TDBGProduccion.Columns(8).Text = 0
 End If
 
 If Not Me.TDBGProduccion.Columns(9).Text = "" Then
 Jueves = val(Me.TDBGProduccion.Columns(9).Text)
 Else
  Jueves = 0
  Me.TDBGProduccion.Columns(9).Text = 0
 End If
 If Not Me.TDBGProduccion.Columns(10).Text = "" Then
  Viernes = val(Me.TDBGProduccion.Columns(10).Text)
 Else
  Viernes = 0
  Me.TDBGProduccion.Columns(10).Text = 0
 End If
 If Not Me.TDBGProduccion.Columns(11).Text = "" Then
  Sabado = val(Me.TDBGProduccion.Columns(11).Text)
 Else
  Sabado = 0
  Me.TDBGProduccion.Columns(11).Text = 0
 End If
 If Not Me.TDBGProduccion.Columns(12).Text = "" Then
  Domingo = val(Me.TDBGProduccion.Columns(12).Text)
 Else
  Domingo = 0
  Me.TDBGProduccion.Columns(12).Text = 0
 End If
 
  If Not Me.TDBGProduccion.Columns(13).Text = "" Then
  Incentivo = val(Me.TDBGProduccion.Columns(13).Text)
 Else
  Incentivo = 0
  Me.TDBGProduccion.Columns(13).Text = 0
 End If
 
 TotalUnidad = Lunes + Martes + Miercoles + Jueves + Viernes + Sabado + Domingo + Incentivo
 If Not unidad = 0 Then
  SalPieza = (TotalUnidad / unidad) * Precio
 Else
  SalPieza = 0
 End If
 Me.TDBGProduccion.Columns(15).Text = SalPieza
 Me.TDBGProduccion.Columns(14).Text = TotalUnidad

 Select Case ColIndex
   Case 1
         Me.TDBGProduccion.SetFocus
         Me.TDBGProduccion.PostMsg (3)
       
       
            
   Case 3
         Me.TDBGProduccion.PostMsg (6)
         Me.TDBGProduccion.SetFocus
End Select
 

 
End Sub

Private Sub TDBGProduccion_AfterUpdate()
  Me.TDBGProduccion.SetFocus
  Me.TDBGProduccion.PostMsg (0)
End Sub

Private Sub TDBGProduccion_BeforeUpdate(Cancel As Integer)
 Dim NumNomina As Integer
 Dim Precio As Double, unidad As Double
 Dim Lunes As Double, Martes As Double, Miercoles As Double, Jueves As Double, Viernes As Double
 Dim Sabado As Double, Domingo As Double, TotalUnidad As Double, Incentivo As Double
 Dim SalPieza As Double
 

 
   Me.DtaConsulta.RecordSource = "SELECT Nomina.*, TipoNomina.TipoPago FROM TipoNomina INNER JOIN Nomina ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina Where (((Nomina.Activa) = 1)) And Nomina.CodTipoNomina = '" & Me.TxtCodTipoNomina.Text & "' "
   'Me.DtaConsulta.RecordSource = "SELECT CodTipoNomina, Nomina, Periodo, Activa From TipoNomina WHERE     (CodTipoNomina = '" & Me.TxtCodTipoNomina.Text & "') AND (Activa = 1) "
   Me.DtaConsulta.Refresh
   If DtaConsulta.Recordset.EOF Then
     MsgBox "No Exste Nomina Activa para Este Empleado", vbCritical, "Sistema de Nominas"
     Me.DBCodEmpleado.Text = ""
     Me.TxtCodEmpleado.Text = ""
     Exit Sub
   Else
     NumNomina = Me.DtaConsulta.Recordset("NumNomina")
    Me.TDBGProduccion.Columns(16).Text = Me.TxtCodEmpleado.Text
    Me.TDBGProduccion.Columns(17).Text = NumNomina
   End If

 If Not Me.TDBGProduccion.Columns(4).Text = "" Then
  Precio = Me.TDBGProduccion.Columns(4).Text
 Else
  Precio = 0
 End If
 
 If Not Me.TDBGProduccion.Columns(5).Text = "" Then
 unidad = Me.TDBGProduccion.Columns(5).Text
 Else
  unidad = 0
 End If
 
 If Not Me.TDBGProduccion.Columns(6).Text = "" Then
  Lunes = Me.TDBGProduccion.Columns(6).Text
 Else
  Lunes = 0
 End If
 If Not Me.TDBGProduccion.Columns(7).Text = "" Then
   Martes = val(Me.TDBGProduccion.Columns(7).Text)
 Else
   Martes = 0
 End If
 If Not Me.TDBGProduccion.Columns(8).Text = "" Then
  Miercoles = val(Me.TDBGProduccion.Columns(8).Text)
 Else
  Miercoles = 0
 End If
 
 If Not Me.TDBGProduccion.Columns(9).Text = "" Then
 Jueves = val(Me.TDBGProduccion.Columns(9).Text)
 Else
  Jueves = 0
 End If
 If Not Me.TDBGProduccion.Columns(10).Text = "" Then
  Viernes = val(Me.TDBGProduccion.Columns(10).Text)
 Else
  Viernes = 0
 End If
 If Not Me.TDBGProduccion.Columns(11).Text = "" Then
  Sabado = val(Me.TDBGProduccion.Columns(11).Text)
 Else
  Sabado = 0
 End If
 If Not Me.TDBGProduccion.Columns(12).Text = "" Then
  Domingo = val(Me.TDBGProduccion.Columns(12).Text)
 Else
  Domingo = 0
 End If
 
  If Not Me.TDBGProduccion.Columns(13).Text = "" Then
  Incentivo = val(Me.TDBGProduccion.Columns(13).Text)
 Else
  Incentivo = 0
 End If
 
 TotalUnidad = Lunes + Martes + Miercoles + Jueves + Viernes + Sabado + Domingo + Incentivo
 If Not unidad = 0 Then
  SalPieza = (TotalUnidad / unidad) * Precio
 Else
  SalPieza = 0
 End If
 Me.TDBGProduccion.Columns(15).Text = SalPieza
 Me.TDBGProduccion.Columns(14).Text = TotalUnidad






End Sub

Private Sub TDBGProduccion_ButtonClick(ByVal ColIndex As Integer)
 QueProducto = "Produccion"
 FrmConsulta.Show 1
End Sub

Private Sub TDBGProduccion_PostEvent(ByVal MsgId As Integer)
   Select Case MsgId
       Case 0
            Me.TDBGProduccion.Split = 0
            Me.TDBGProduccion.col = 0
       Case 1
             Me.TDBGProduccion.Refresh
       Case 2
'            DBGTransacciones.SetFocus
            'Set focus to split zero and column 0
            Me.TDBGProduccion.Split = 0
            Me.TDBGProduccion.col = 2
       Case 3
            Me.TDBGProduccion.SetFocus
            'Set focus to split zero and column 0
            Me.TDBGProduccion.Split = 0
            Me.TDBGProduccion.col = 3
      Case 4
            Me.TDBGProduccion.SetFocus
            'Set focus to split zero and column 0
            Me.TDBGProduccion.Split = 0
            Me.TDBGProduccion.col = 4
      Case 6
            Me.TDBGProduccion.Split = 0
            Me.TDBGProduccion.col = 6
   
   End Select
End Sub

Private Sub TDBGridHoras_BeforeUpdate(Cancel As Integer)
 Dim Lunes As Double, Martes As Double, Miercoles As Double, Jueves As Double, Viernes As Double
 Dim Sabado As Double, Domingo As Double
 
   Me.DtaConsulta.RecordSource = "SELECT Nomina.*, TipoNomina.TipoPago FROM TipoNomina INNER JOIN Nomina ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina Where (((Nomina.Activa) = 1)) And Nomina.CodTipoNomina = '" & Me.TxtCodTipoNomina.Text & "' "
   Me.DtaConsulta.Refresh
   If DtaConsulta.Recordset.EOF Then
     MsgBox "No Exste Nomina Activa para Este Empleado", vbCritical, "Sistema de Nominas"
     Me.DBCodEmpleado.Text = ""
     Me.TxtCodEmpleado.Text = ""
     Exit Sub
   Else
    NumNomina = Me.DtaConsulta.Recordset("NumNomina")
    Me.TDBGridHoras.Columns(0).Text = Me.TxtCodEmpleado.Text
    Me.TDBGridHoras.Columns(1).Text = NumNomina
    
    If Me.TDBGridHoras.Columns(2).Text = "" Then
        Me.DtaConsulta.RecordSource = "SELECT CodEmpleado, NumNomina, NumLinea, Lunes, Martes, Miercoles, Jueves, Viernes, Sabado, Domingo, TotalHoras, SalarioHora, TotalSalarioHora, Pagado From dbo.DetalleHorasProduccion ORDER BY NumLinea"
        Me.DtaConsulta.Refresh
        If DtaConsulta.Recordset.EOF Then
          CodLinea = 1
        Else
          Me.DtaConsulta.Recordset.MoveLast
          CodLinea = DtaConsulta.Recordset("NumLinea") + 1
        End If
           Me.TDBGridHoras.Columns(2).Text = CodLinea
     End If
   End If
   
  If Not Me.TDBGridHoras.Columns(3).Text = "" Then
   Lunes = Me.TDBGridHoras.Columns(3).Text
  Else
    Lunes = 0
    Me.TDBGridHoras.Columns(3).Text = 0
  End If
  If Not Me.TDBGridHoras.Columns(4).Text = "" Then
   Martes = Me.TDBGridHoras.Columns(4).Text
  Else
   Martes = 0
   Martes = Me.TDBGridHoras.Columns(4).Text = 0
  End If
  If Not Me.TDBGridHoras.Columns(5).Text = "" Then
   Miercoles = Me.TDBGridHoras.Columns(5).Text
  Else
   Miercoles = 0
   Me.TDBGridHoras.Columns(5).Text = 0
  End If
  If Not Me.TDBGridHoras.Columns(6).Text = "" Then
   Jueves = Me.TDBGridHoras.Columns(6).Text
  Else
   Jueves = 0
   Me.TDBGridHoras.Columns(6).Text = 0
  End If
  If Not Me.TDBGridHoras.Columns(7).Text = "" Then
   Viernes = Me.TDBGridHoras.Columns(7).Text
  Else
   Viernes = 0
   Me.TDBGridHoras.Columns(7).Text = 0
  End If
  If Not Me.TDBGridHoras.Columns(8).Text = "" Then
   Sabado = Me.TDBGridHoras.Columns(8).Text
  Else
   Sabado = 0
   Me.TDBGridHoras.Columns(8).Text = 0
  End If
  If Not Me.TDBGridHoras.Columns(9).Text = "" Then
   Domingo = Me.TDBGridHoras.Columns(9).Text
  Else
    Domingo = 0
    Me.TDBGridHoras.Columns(9).Text = 0
  End If
   
   
   TotalHoras = Lunes + Martes + Miercoles + Jueves + Viernes + Sabado + Domingo
   Me.TDBGridHoras.Columns(10).Text = TotalHoras
'   Me.DtaConsulta.RecordSource = "SELECT CodEmpleado, TarifaHoraria From dbo.Empleado WHERE (CodEmpleado = '" & Me.TxtCodEmpleado.Text & "' )"
'   Me.DtaConsulta.Refresh
'   If Not DtaConsulta.Recordset.EOF Then
'     TarifaHoraria = Me.DtaConsulta.Recordset("TarifaHoraria")
'   End If
   TarifaHoraria = BuscaTarifaHoraria(Me.TxtCodEmpleado.Text)
   Me.TDBGridHoras.Columns(11).Text = Format(TarifaHoraria, "##,##0.0000")
   Me.TDBGridHoras.Columns(12).Text = Format(TarifaHoraria * TotalHoras, "##,##0.00")
End Sub

