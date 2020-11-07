VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmTraslados 
   Caption         =   "Trasladar Historicos"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   9750
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoTraslado 
      Height          =   375
      Left            =   5400
      Top             =   10440
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
      Caption         =   "AdoTraslado"
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
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9000
      TabIndex        =   39
      Top             =   9240
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Height          =   2895
      Left            =   120
      TabIndex        =   31
      Top             =   6240
      Width           =   10455
      Begin TrueOleDBGrid80.TDBGrid TDBGrid1 
         Bindings        =   "TrasladoHistoricos.frx":0000
         Height          =   2415
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   4260
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
         Splits.Count    =   1
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
         _StyleDefs(114) =   "Named:id=33:Normal"
         _StyleDefs(115) =   ":id=33,.parent=0"
         _StyleDefs(116) =   "Named:id=34:Heading"
         _StyleDefs(117) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(118) =   ":id=34,.wraptext=-1"
         _StyleDefs(119) =   "Named:id=35:Footing"
         _StyleDefs(120) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(121) =   "Named:id=36:Selected"
         _StyleDefs(122) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(123) =   "Named:id=37:Caption"
         _StyleDefs(124) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(125) =   "Named:id=38:HighlightRow"
         _StyleDefs(126) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(127) =   "Named:id=39:EvenRow"
         _StyleDefs(128) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(129) =   "Named:id=40:OddRow"
         _StyleDefs(130) =   ":id=40,.parent=33"
         _StyleDefs(131) =   "Named:id=41:RecordSelector"
         _StyleDefs(132) =   ":id=41,.parent=34"
         _StyleDefs(133) =   "Named:id=42:FilterBar"
         _StyleDefs(134) =   ":id=42,.parent=33"
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Criterios"
      Height          =   975
      Left            =   240
      TabIndex        =   30
      Top             =   5160
      Width           =   10455
      Begin VB.CommandButton CmdTrasladar 
         Caption         =   "Trasladar"
         Height          =   375
         Left            =   8880
         TabIndex        =   38
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdConsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   7200
         TabIndex        =   37
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTFecha1 
         Height          =   375
         Left            =   1200
         TabIndex        =   34
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   75169793
         CurrentDate     =   39928
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "TrasladoHistoricos.frx":001B
         TabIndex        =   33
         Top             =   360
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   375
         Left            =   3720
         OleObjectBlob   =   "TrasladoHistoricos.frx":008F
         TabIndex        =   35
         Top             =   360
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTFecha2 
         Height          =   375
         Left            =   4560
         TabIndex        =   36
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   75169793
         CurrentDate     =   39928
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   240
      TabIndex        =   13
      Top             =   2880
      Width           =   10455
      Begin VB.TextBox TxtCodEmpleadoDestino 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   6720
         TabIndex        =   28
         Top             =   1680
         Width           =   975
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
         Picture         =   "TrasladoHistoricos.frx":00FF
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtNombre1Destino 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox TxtApellido2Destino 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   23
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox TxtApellido1Destino 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   22
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox TxtNombre2Destino 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   21
         Top             =   960
         Width           =   3615
      End
      Begin VB.PictureBox Picture2 
         Height          =   1095
         Left            =   5520
         ScaleHeight     =   1035
         ScaleWidth      =   1035
         TabIndex        =   14
         Top             =   720
         Width           =   1095
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   1020
            Left            =   0
            Picture         =   "TrasladoHistoricos.frx":024D
            Top             =   0
            Width           =   1020
         End
      End
      Begin TrueOleDBList80.TDBCombo TDBDestino 
         Bindings        =   "TrasladoHistoricos.frx":328F
         Height          =   315
         Left            =   1680
         TabIndex        =   15
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
         _PropDict       =   $"TrasladoHistoricos.frx":32AA
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "TrasladoHistoricos.frx":3354
         TabIndex        =   16
         Top             =   1680
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "TrasladoHistoricos.frx":33D2
         TabIndex        =   17
         Top             =   1320
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "TrasladoHistoricos.frx":344E
         TabIndex        =   18
         Top             =   960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "TrasladoHistoricos.frx":34C8
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "TrasladoHistoricos.frx":3540
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   435
         Left            =   6720
         TabIndex        =   26
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
         Caption         =   "Empleado Destino"
         ButtonStyle     =   4
         AutoRepeat      =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   10455
      Begin VB.TextBox TxtCodEmpleadoOrigen 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   6840
         TabIndex        =   27
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox TxtNombre1Origen 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   11
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox TxtApellido2Origen 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox TxtApellido1Origen 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox TxtNombre2Origen 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   8
         Top             =   960
         Width           =   3615
      End
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   5520
         ScaleHeight     =   1035
         ScaleWidth      =   1035
         TabIndex        =   1
         Top             =   720
         Width           =   1095
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   1020
            Left            =   0
            Picture         =   "TrasladoHistoricos.frx":35AA
            Top             =   0
            Width           =   1020
         End
      End
      Begin TrueOleDBList80.TDBCombo TDBOrigen 
         Bindings        =   "TrasladoHistoricos.frx":65EC
         Height          =   315
         Left            =   1680
         TabIndex        =   2
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
         _PropDict       =   $"TrasladoHistoricos.frx":6607
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
         OleObjectBlob   =   "TrasladoHistoricos.frx":66B1
         TabIndex        =   3
         Top             =   1680
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "TrasladoHistoricos.frx":672F
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "TrasladoHistoricos.frx":67AB
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "TrasladoHistoricos.frx":6825
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "TrasladoHistoricos.frx":689D
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin Threed.SSCommand SSCommand1Origen 
         Height          =   435
         Left            =   6720
         TabIndex        =   12
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
         Caption         =   "Empleado Origen"
         ButtonStyle     =   4
         AutoRepeat      =   -1  'True
      End
   End
   Begin MSAdodcLib.Adodc DtaEmpleados 
      Height          =   375
      Left            =   480
      Top             =   10560
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
   Begin Threed.SSCommand CmdAcercade 
      Height          =   435
      Left            =   240
      TabIndex        =   29
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
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
      Caption         =   "Traslado de Saldos Historicos"
      ButtonStyle     =   4
      AutoRepeat      =   -1  'True
   End
   Begin MSAdodcLib.Adodc AdoHistorial 
      Height          =   375
      Left            =   480
      Top             =   10200
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
End
Attribute VB_Name = "FrmTraslados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub CmdConsultar_Click()
Dim Fecha1 As String, Fecha2 As String
Dim CodEmpleado As Double


CodEmpleado = Me.TxtCodEmpleadoOrigen.Text

Fecha1 = Format(Me.DTFecha1.Value, "yyyy-mm-dd")
Fecha2 = Format(Me.DTFecha2.Value, "yyyy-mm-dd")

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
SQLHistorial = SQLHistorial & "GROUP BY DetalleNomina.CodEmpleado, Nomina.Mes, Nomina.Ano, Nomina.FechaNominaINI" & vbLf
SQLHistorial = SQLHistorial & "Having (Nomina.FechaNominaINI BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) AND (dbo.DetalleNomina.CodEmpleado = " & CodEmpleado & ")"

Me.AdoHistorial.RecordSource = SQLHistorial
Me.AdoHistorial.Refresh




End Sub

Private Sub Command2_Click()

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdTrasladar_Click()
Dim Fecha1 As String, Fecha2 As String
Dim CodEmpleado As Double


CodEmpleado = Me.TxtCodEmpleadoOrigen.Text

Fecha1 = Format(Me.DTFecha1.Value, "yyyy-mm-dd")
Fecha2 = Format(Me.DTFecha2.Value, "yyyy-mm-dd")



Me.AdoTraslado.RecordSource = "SELECT * FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina " & _
                              "WHERE (Nomina.FechaNominaINI BETWEEN CONVERT(DATETIME, '" & Fecha1 & "', 102) AND CONVERT(DATETIME, '" & Fecha2 & "', 102)) AND (dbo.DetalleNomina.CodEmpleado = " & CodEmpleado & ")"
Me.AdoTraslado.Refresh

Do While Not Me.AdoTraslado.Recordset.EOF
  Me.AdoTraslado.Recordset("CodEmpleado") = Me.TxtCodEmpleadoDestino.Text
  Me.AdoTraslado.Recordset.Update
  Me.AdoTraslado.Recordset.MoveNext
Loop

Me.AdoHistorial.RecordSource = "SELECT * FROM DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina " & _
                              "WHERE (Nomina.FechaNominaINI BETWEEN CONVERT(DATETIME, '" & Fecha2 & "', 102) AND CONVERT(DATETIME, '" & Fecha1 & "', 102)) AND (dbo.DetalleNomina.CodEmpleado = " & CodEmpleado & ")"
Me.AdoHistorial.Refresh
MsgBox "Empleado Trasladado con Exito!!!"


End Sub

Private Sub Form_Load()

MDIPrimero.Skin1.ApplySkin hWnd

 With Me.DtaEmpleados
   .ConnectionString = Conexion
   '.RecordSource = "SELECT   Empleado.CodEmpleado1, Empleado.CodEmpleado,Empleado.Nombre1 + Empleado.Nombre2 + Empleado.Apellido1 + Empleado.Apellido2 AS Nombres, Departamento.Departamento, Cargo.Cargo,Empleado.Activo FROM Departamento INNER JOIN Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento ORDER BY Empleado.CodEmpleado1"
   .RecordSource = "SELECT Empleado.CodEmpleado1, Empleado.CodEmpleado,Empleado.Nombre1 +' '+ Empleado.Nombre2+' ' + Empleado.Apellido1+' ' + Empleado.Apellido2 AS Nombres, Empleado.Activo FROM Departamento INNER JOIN  Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento ORDER BY Empleado.CodEmpleado1 "
   .Refresh
 End With
 
  With Me.AdoHistorial
    .ConnectionString = Conexion
 End With
 
 With Me.AdoTraslado
    .ConnectionString = Conexion
 End With
 
 Me.TDBGrid1.EvenRowStyle.BackColor = &H80000013
 Me.TDBGrid1.OddRowStyle.BackColor = &H80000005
 Me.TDBGrid1.AlternatingRowStyle = True
 
 Me.DTFecha1.Value = Now
 Me.DTFecha2.Value = Now
 
End Sub

Private Sub TDBDestino_ItemChange()
Me.TxtCodEmpleadoDestino.Text = Me.TDBDestino.Columns(1).Text
End Sub

Private Sub TDBOrigen_ItemChange()
Me.TxtCodEmpleadoOrigen.Text = Me.TDBOrigen.Columns(1).Text
End Sub

Private Sub TxtCodEmpleadoDestino_Change()

Dim SqlPagos As String
Dim SqlEmpleados As String
Dim MesIni As Byte
Dim Annoini As Integer
Dim MesFin As Byte
Dim AnnoFin As Integer
Dim CodEmpleado As Double
Annoini = val(FrmHistorial.CmdAnnoIni.Text)

CodEmpleado = Me.TxtCodEmpleadoOrigen.Text



SqlEmpleados = "SELECT CodEmpleado1, CodEmpleado, Nombre1 + Nombre2 + Apellido1 + Apellido2 AS Nombres, Nombre1, Nombre2, Apellido1, Apellido2, NumHijos,Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula, Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss,NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria, PorcentajeComision, OtrosIngresos, DescripOtrIngre, ExentoInss,ExentoIr , PagoInssPatronal, SalarioMinimo, Observaciones, Liquidado, Ausente, Activo From Empleado Where (CodEmpleado = " & CodEmpleado & ") ORDER BY CodEmpleado1"
FrmHistorial.AdoBusca.RecordSource = SqlEmpleados
FrmHistorial.AdoBusca.Refresh

If Not FrmHistorial.AdoBusca.Recordset.EOF Then
    TxtNombre1Destino.Text = FrmHistorial.AdoBusca.Recordset("Nombre1")
    TxtNombre2Destino.Text = FrmHistorial.AdoBusca.Recordset("Nombre2")
    TxtApellido1Destino.Text = FrmHistorial.AdoBusca.Recordset("Apellido1")
    TxtApellido2Destino.Text = FrmHistorial.AdoBusca.Recordset("Apellido2")
'    TxtCargo.Text = Me.AdoBusca.Recordset("Cargo")
'    txtDepartamento = Me.AdoBusca.Recordset("departamento")
Else
    TxtNombre1Destino.Text = ""
    TxtNombre2Destino.Text = ""
    TxtApellido1Destino.Text = ""
    TxtApellido2Destino.Text = ""

End If

Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub TxtCodEmpleadoOrigen_Change()
'On Error GoTo TipoErr
Dim SqlPagos As String
Dim SqlEmpleados As String
Dim MesIni As Byte
Dim Annoini As Integer
Dim MesFin As Byte
Dim AnnoFin As Integer
Dim CodEmpleado As Double
Annoini = val(FrmHistorial.CmdAnnoIni.Text)

CodEmpleado = Me.TxtCodEmpleadoOrigen.Text



SqlEmpleados = "SELECT CodEmpleado1, CodEmpleado, Nombre1 + Nombre2 + Apellido1 + Apellido2 AS Nombres, Nombre1, Nombre2, Apellido1, Apellido2, NumHijos,Direccion, Nacionalidad, CodigoPostal, Sexo, CodInss, CodIr, NumCedula, Sindicalista, CodDepartamento, CodGrupo, CodCargo, NumeroInss,NumeroRuc, CodTipoNomina, DiasDescuento, SueldoPeriodo, TarifaHoraria, PorcentajeComision, OtrosIngresos, DescripOtrIngre, ExentoInss,ExentoIr , PagoInssPatronal, SalarioMinimo, Observaciones, Liquidado, Ausente, Activo From Empleado Where (CodEmpleado = " & CodEmpleado & ") ORDER BY CodEmpleado1"
FrmHistorial.AdoBusca.RecordSource = SqlEmpleados
FrmHistorial.AdoBusca.Refresh

If Not FrmHistorial.AdoBusca.Recordset.EOF Then
    TxtNombre1Origen.Text = FrmHistorial.AdoBusca.Recordset("Nombre1")
    TxtNombre2Origen.Text = FrmHistorial.AdoBusca.Recordset("Nombre2")
    TxtApellido1Origen.Text = FrmHistorial.AdoBusca.Recordset("Apellido1")
    TxtApellido2Origen.Text = FrmHistorial.AdoBusca.Recordset("Apellido2")
'    TxtCargo.Text = Me.AdoBusca.Recordset("Cargo")
'    txtDepartamento = Me.AdoBusca.Recordset("departamento")
Else
    TxtNombre1Origen.Text = ""
    TxtNombre2Origen.Text = ""
    TxtApellido1Origen.Text = ""
    TxtApellido2Origen.Text = ""

End If



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
