VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmSolicitud 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de Permisos"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10125
   Begin MSComctlLib.ProgressBar pbEmpleados 
      Height          =   375
      Left            =   3600
      TabIndex        =   46
      Top             =   5160
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc AdoAuxiliar 
      Height          =   495
      Left            =   4200
      Top             =   8520
      Width           =   2895
      _ExtentX        =   5106
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
   Begin MSAdodcLib.Adodc AdoAux 
      Height          =   330
      Left            =   7200
      Top             =   8040
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "AdoAux"
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
   Begin VB.TextBox txtHorasSolicitud 
      Height          =   285
      Left            =   7200
      TabIndex        =   44
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox TxtCodEmpleado 
      Height          =   285
      Left            =   6120
      TabIndex        =   39
      Top             =   6480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox TxtObservaciones 
      Height          =   615
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   5640
      Width           =   8535
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   9855
      _Version        =   786432
      _ExtentX        =   17383
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "Datos Periodo Vacacional"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox TxtDisponibles 
         Height          =   315
         Left            =   8280
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtDisfrutados 
         Height          =   315
         Left            =   4800
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtDiasVacaciones 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmSolicitud.frx":0000
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmSolicitud.frx":007C
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   6960
         OleObjectBlob   =   "FrmSolicitud.frx":00FA
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox TxtNumero 
      Height          =   315
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin XtremeSuiteControls.DateTimePicker DtpFechaSolicitud 
      Height          =   315
      Left            =   7920
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   556
      _StockProps     =   68
      Format          =   1
      CurrentDate     =   41714.4187384259
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   6720
      OleObjectBlob   =   "FrmSolicitud.frx":0178
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   6600
      OleObjectBlob   =   "FrmSolicitud.frx":01F4
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   255
      Left            =   6840
      OleObjectBlob   =   "FrmSolicitud.frx":0272
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   9855
      _Version        =   786432
      _ExtentX        =   17383
      _ExtentY        =   2355
      _StockProps     =   79
      Caption         =   "Datos de Solicitud"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox TxtDiasSolicitud 
         Height          =   285
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
      Begin XtremeSuiteControls.DateTimePicker DtpFechaInicio 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   480
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   41714.4187384259
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmSolicitud.frx":02EC
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin XtremeSuiteControls.DateTimePicker DtpFechaFin 
         Height          =   315
         Left            =   4080
         TabIndex        =   15
         Top             =   480
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   41714.4187384259
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   3240
         OleObjectBlob   =   "FrmSolicitud.frx":0362
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   5880
         OleObjectBlob   =   "FrmSolicitud.frx":03D2
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   5760
         OleObjectBlob   =   "FrmSolicitud.frx":0452
         TabIndex        =   43
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "."
         Height          =   255
         Left            =   8280
         TabIndex        =   45
         Top             =   840
         Width           =   15
      End
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   390
      Left            =   9480
      TabIndex        =   22
      Top             =   360
      Width           =   390
      _Version        =   786432
      _ExtentX        =   688
      _ExtentY        =   688
      _StockProps     =   79
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmSolicitud.frx":04D4
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   8520
      TabIndex        =   23
      Top             =   6480
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmSolicitud.frx":09D6
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdGrabar 
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   6480
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Autorizar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmSolicitud.frx":0EDA
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdBorrar 
      Height          =   375
      Left            =   1680
      TabIndex        =   25
      Top             =   6480
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Imprimir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmSolicitud.frx":323E
      ImageAlignment  =   0
   End
   Begin MSAdodcLib.Adodc DtaEmpleados 
      Height          =   375
      Left            =   360
      Top             =   7080
      Width           =   2895
      _ExtentX        =   5106
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
   Begin MSAdodcLib.Adodc DtaEmpleado 
      Height          =   375
      Left            =   240
      Top             =   7560
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "DtaEmpleado"
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
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1455
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Width           =   9855
      _Version        =   786432
      _ExtentX        =   17383
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Datos del Empleado"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox TxtNombre 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         TabIndex        =   29
         Top             =   360
         Width           =   5415
      End
      Begin VB.TextBox TxtDepartamento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         TabIndex        =   27
         Top             =   840
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmSolicitud.frx":36F2
         TabIndex        =   28
         Top             =   360
         Width           =   735
      End
      Begin TrueOleDBList80.TDBCombo DBCodigoEmpleado 
         Bindings        =   "FrmSolicitud.frx":375E
         Height          =   315
         Left            =   840
         TabIndex        =   30
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   8811
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
         DropdownPosition=   1
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         AddItemSeparator=   ";"
         _PropDict       =   $"FrmSolicitud.frx":3779
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
      Begin MSDataListLib.DataCombo DBCCargo 
         Bindings        =   "FrmSolicitud.frx":3823
         Height          =   315
         Left            =   8040
         TabIndex        =   31
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         ListField       =   "Cargo"
         Text            =   ""
      End
      Begin XtremeSuiteControls.DateTimePicker DtpFechaIngreso 
         Height          =   315
         Left            =   1320
         TabIndex        =   32
         Top             =   840
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   68
         Enabled         =   0   'False
         Format          =   1
         CurrentDate     =   41714.4187384259
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmSolicitud.frx":383A
         TabIndex        =   33
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   7320
         OleObjectBlob   =   "FrmSolicitud.frx":38A8
         TabIndex        =   34
         Top             =   840
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmSolicitud.frx":3912
         TabIndex        =   35
         Top             =   840
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "FrmSolicitud.frx":398C
         TabIndex        =   36
         Top             =   840
         Width           =   1095
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Height          =   390
         Left            =   3000
         TabIndex        =   37
         Top             =   240
         Width           =   390
         _Version        =   786432
         _ExtentX        =   688
         _ExtentY        =   688
         _StockProps     =   79
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmSolicitud.frx":3A02
         ImageAlignment  =   0
      End
   End
   Begin XtremeSuiteControls.ComboBox CmbTipoSolicitud 
      Height          =   315
      Left            =   7920
      TabIndex        =   38
      Top             =   840
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   375
      Left            =   3240
      TabIndex        =   40
      Top             =   6480
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Anular"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmSolicitud.frx":3F04
      ImageAlignment  =   0
   End
   Begin MSAdodcLib.Adodc AdoSolicitud 
      Height          =   375
      Left            =   3720
      Top             =   7320
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "AdoSolicitud"
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
   Begin MSAdodcLib.Adodc AdoConsecutivo 
      Height          =   375
      Left            =   3720
      Top             =   7800
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "AdoConsecutivo"
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
      Left            =   240
      Top             =   8160
      Width           =   2895
      _ExtentX        =   5106
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
   Begin XtremeSuiteControls.ComboBox CmbClasificado 
      Height          =   315
      Left            =   1440
      TabIndex        =   42
      Top             =   5160
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Style           =   2
      Appearance      =   6
   End
   Begin MSAdodcLib.Adodc AdoNomina 
      Height          =   375
      Left            =   6720
      Top             =   7320
      Width           =   2895
      _ExtentX        =   5106
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
   Begin VB.Label txtEmpleados 
      Caption         =   "Label5"
      Height          =   255
      Left            =   5880
      TabIndex        =   47
      Top             =   5160
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Clasificacion"
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitud y Autorizacion de Dias de Vacaciones"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label Label2 
      Caption         =   "Observaciones:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   5640
      Width           =   1335
   End
End
Attribute VB_Name = "FrmSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command4_Click()
Unload Me
End Sub



Private Sub CmbTipoSolicitud_LostFocus()
If Me.CmbTipoSolicitud.Text = "Ausente" Or CmbTipoSolicitud.Text = "Suspension" Then
        Me.txtdiasvacaciones.Visible = False
        Me.TxtDisfrutados.Visible = False
        Me.txtDisponibles.Visible = False
    Else
        Me.txtdiasvacaciones.Visible = True
        Me.TxtDisfrutados.Visible = True
        Me.txtDisponibles.Visible = True
    End If
End Sub

Private Sub cmdborrar_Click()
Dim rpt As Object
Dim fPreview As New FrmPreview

If CDbl(Me.TxtDiasSolicitud.Text) > 3 Then

  Set rpt = New ArepSolicitudVacaLargo


    rpt.txtNombre.Text = Me.txtNombre.Text
    rpt.txtCargo.Text = Me.DBCCargo.Text
    rpt.txtDesde.Text = Format(Me.DtpFechaInicio.Value, "dd/MM/yyyy")
    rpt.txtHasta.Text = Format(Me.DtpFechaFin.Value, "dd/MM/yyyy")
    rpt.txtRegresa.Text = Format(DateAdd("d", 1, Me.DtpFechaFin.Value), "dd/MM/yyyy")
    
    rpt.txtDia.Text = Format(DateTime.Now, "d")
    rpt.txtMes.Text = Format(DateTime.Now, "MMMM")
    rpt.txtCantidad.Text = Me.TxtDiasSolicitud.Text
    
    MDIPrimero.DtaConsulta.RecordSource = "SELECT     TOP (1) Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombre, Historico.FechaContratoVac, DATEADD(month,    (YEAR(Historico.FechaContratoVac) - 1900) * 12 + MONTH(Historico.FechaContratoVac), - 1) AS UdMes  FROM         Empleado INNER JOIN   Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (Empleado.CodEmpleado1 = '" & Me.DBCodigoEmpleado.Text & "')"
    MDIPrimero.DtaConsulta.Refresh

'///////// Inicializo parametros generales ////////////
Dim VacacionesAcumuladas, VacacionesSolicitadas, TotalVacacionesAcumuladas, TotalVacacionSolicitada, SaldoActual, tempVacacionSolicitada, tempVacacionAcumulada As Double
Dim NombreCompleto As String
NombreCompleto = MDIPrimero.DtaConsulta.Recordset("Nombre")
TotalVacacionesAcumuladas = 0
TotalVacacionSolicitada = 0
Dim Inicio, tempInicio, Fin, tempFin As Date

Inicio = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
tempInicio = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
Fin = DateTime.Now
tempFin = MDIPrimero.DtaConsulta.Recordset("udMes")


         SaldoActual = 0

    Do While (Inicio < Fin)
     tempVacacionSolicitada = 0
     tempVacacionAcumulada = 0
          
          If CDbl(Format(tempFin, "d")) > 30 Then
                tempFin = tempFin - 1
            End If
            
            Dim DiasMes As Double
            
            DiasMes = DateDiff("d", DateSerial(Year(tempInicio), Month(tempInicio), 1), DateSerial(Year(tempInicio), Month(tempInicio) + 1, 0)) + 1
            
            If Format(tempInicio, "MMMM") = "Febrero" Then
                If (DateDiff("d", tempInicio, tempFin) + 1) < DiasMes Then
                    tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
                Else
                    tempVacacionAcumulada = 2.5
                End If
            Else
                If (DateDiff("d", tempInicio, tempFin) + 1) < 30 Then
                    tempVacacionAcumulada = (DateDiff("d", tempInicio, tempFin) + 1) / 12
                Else
                    tempVacacionAcumulada = 2.5
                End If
            End If
          
          TotalVacacionesAcumuladas = tempVacacionAcumulada + TotalVacacionesAcumuladas
     
             
             '////////     /////////     ////////        /////////       /////////       ////////        //////
             '////////// Calculo el total de dias y horas solicitadas en el rango de fechas recorrido /////////
             '/////////////////////////////////////////////////////////////////////////////////////////////////
           
             AdoAux.RecordSource = "SELECT     SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM         SolicitudVacaciones WHERE     (CodigoEmpleado = '" & Me.DBCodigoEmpleado.Text & "') AND (FechaInicio BETWEEN '" & Format(tempInicio, "dd/MM/yyyy") & "' AND '" & Format(tempFin, "dd/MM/yyyy") & "')"
             AdoAux.Refresh
             If Not AdoAux.Recordset.EOF Then
                If Not IsNull(AdoAux.Recordset("VacacionesSolicitadas")) Then
                    tempVacacionSolicitada = AdoAux.Recordset("VacacionesSolicitadas")
                    TotalVacacionSolicitada = TotalVacacionSolicitada + tempVacacionSolicitada
                Else
                    tempVacacionSolicitada = 0
                End If
             Else
                 tempVacacionSolicitada = 0
             End If
             
             SaldoActual = SaldoActual + (tempVacacionAcumulada - tempVacacionSolicitada)
          
             
             tempFin = DateAdd("d", 2, tempFin)
             
             tempInicio = DateSerial(Year(tempFin), Month(tempFin), 1) ' Inicio
             tempFin = DateSerial(Year(tempFin), Month(tempFin) + 1, 0) 'Fin
                         'ponerle temp inicio para que los dias  no varien
            Inicio = tempInicio
    Loop
    
    rpt.txtDiasAcumulados.Text = Format(TotalVacacionesAcumuladas, "##,##0.00")
    rpt.txtDiasSolicitados.Text = Format(TotalVacacionSolicitada, "##,##0.00")
    rpt.txtSaldoActual.Text = Format(SaldoActual, "##,##0.00")
    rpt.txtTotalDisponibles.Text = Format(CDbl(Me.txtDisponibles.Text), "##,##0.00")
    rpt.txtTotalVacaciones.Text = Format(CDbl(Me.txtdiasvacaciones.Text), "##,#0.00")
    rpt.txtTotalSolicitados.Text = Format(CDbl(Me.TxtDisfrutados.Text), "##,#0.00")
    
     fPreview.arv.ReportSource = rpt
               fPreview.Show 1
Else
Set rpt = New ArepSolicitudVacaCorto
    rpt.txtNosolicitud.Text = "Solicitud No: " & Me.TxtNumero.Text
    rpt.txtFechaSolicitud.Text = Format(Me.DtpFechaSolicitud.Value, "dd/MM/yyyy")
    rpt.txtFechaInicio.Text = Format(Me.DtpFechaInicio.Value, "dd/MM/yyyy")
    rpt.txtFechaFin.Caption = Format(Me.DtpFechaFin.Value, "dd/MM/yyyy")
    rpt.txtCantidadDias.Text = Me.TxtDiasSolicitud.Text
    rpt.txtNombre.Text = Me.txtNombre.Text
    rpt.txtCargo.Text = Me.DBCCargo.Text
    rpt.txtCodigo.Caption = Me.DBCodigoEmpleado.Text
    rpt.txtDepartamento.Caption = Me.txtDepartamento.Text
    rpt.txtHorasdisfruta.Caption = Me.txtHorasSolicitud.Text
    rpt.txtdiasvacaciones.Caption = Me.txtdiasvacaciones.Text
    rpt.txtDiasdisfrutados.Caption = Me.TxtDisfrutados.Text
    rpt.txtdiasdisponibles.Caption = Me.txtDisponibles.Text
    
    
     fPreview.arv.ReportSource = rpt
               fPreview.Show 1
End If
     DisfrutadosConsulta = 0
     pActualiza = False
End Sub

Private Sub CmdGrabar_Click()



If DBCodigoEmpleado.Text = "" Then
MsgBox "Tenes que agregar un empelado", vbCritical, "Zeus Nominas"
    Exit Sub
End If

Me.CmdGrabar.Enabled = False
 Dim Fecha As String, FechaIni As Date, FechaFin As Date, Resultado As Double, CantDias As Double, Horas As Double
 Dim rs As New ADODB.Recordset
 
If Not IsNumeric(txtHorasSolicitud.Text) Then
 MsgBox "El campo Horas Solicitud esta vacio o tiene un formato incorrecto", vbCritical, "Zeus Nominas"
 Me.CmdGrabar.Enabled = True
Exit Sub
End If

    If ExisteEmpleado(Me.DBCodigoEmpleado.Text) = False Then
      MsgBox "Empleado no Existe", vbCritical, "Zeus Nominas"
      Me.CmdGrabar.Enabled = True
      Exit Sub
    End If
    
    AdoNomina.RecordSource = "SELECT TipoNomina.Horas FROM TipoNomina INNER JOIN  Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina WHERE        (Empleado.CodEmpleado1 = '" & Me.DBCodigoEmpleado.Text & "')"
    AdoNomina.Refresh
    
    If Not Me.AdoNomina.Recordset.EOF Then
    Horas = Me.AdoNomina.Recordset("Horas")
    End If
    
    If (CDbl(txtHorasSolicitud.Text) > Horas) Then
    MsgBox "Segun la nomina " & txtNombre.Text & " labora al dia unicamente " & Horas & " horas", vbCritical, "Zeus Nominas"
      Me.CmdGrabar.Enabled = True
      Exit Sub
    End If
    
    '//////////////////////////////////////////
    '/// Comienzo a recorrer los registros ////
    '//////////////////////////////////////////
    
    AdoAux.RecordSource = "SELECT     Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombre     FROM         Empleado INNER JOIN       TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina      WHERE  empleado.Activo = 'true' and    (TipoNomina.Nomina = N'" & Me.DBCodigoEmpleado.Text & "')"
    AdoAux.Refresh
    
    
    If Not Me.AdoAux.Recordset.EOF Then
            pbEmpleados.Visible = True
            txtEmpleados.Visible = True
            pbEmpleados.Min = 0
            pbEmpleados.Max = AdoAux.Recordset.RecordCount
            Dim CodigoEmpleadoR As String
            Dim Contador As Integer
            Contador = 0
         Do While Not AdoAux.Recordset.EOF
            CodigoEmpleadoR = AdoAux.Recordset("CodEmpleado1")
            pbEmpleados.Value = Contador
            txtEmpleados.Caption = AdoAux.Recordset("Nombre")
                    '///////////////////////////////////////
                    '///////////////////////////////////////
                    
                    
                        Fecha = Format(Me.DtpFechaSolicitud.Value, "yyyy-mm-dd")
                        Me.AdoSolicitud.RecordSource = "SELECT  * From SolicitudVacaciones WHERE (NumeroSolicitud = '" & Me.TxtNumero.Text & "')"
                        Me.AdoSolicitud.Refresh
                        If Me.AdoSolicitud.Recordset.EOF Then
                          Me.AdoSolicitud.Recordset.AddNew
                          
                          If txtdiasvacaciones.Text = "" Then
                           txtdiasvacaciones.Text = 0
                          End If
                          
                          If TxtDisfrutados.Text = "" Then
                           TxtDisfrutados.Text = 0
                          End If
                          
                           Me.AdoSolicitud.Recordset("FechaSolicitud") = Format(Me.DtpFechaSolicitud.Value, "dd/mm/yyyy")
                           Me.AdoSolicitud.Recordset("NumeroSolicitud") = Me.TxtNumero.Text
                           Me.AdoSolicitud.Recordset("TipoSolicitud") = Me.CmbTipoSolicitud.Text
                           Me.AdoSolicitud.Recordset("CodigoEmpleado") = CodigoEmpleadoR
                           Me.AdoSolicitud.Recordset("DiasVacaciones") = Me.txtdiasvacaciones.Text
                           Me.AdoSolicitud.Recordset("DiasDisfrutados") = Me.TxtDisfrutados.Text
                           Me.AdoSolicitud.Recordset("FechaInicio") = Me.DtpFechaInicio.Value
                           Me.AdoSolicitud.Recordset("FechaFin") = Me.DtpFechaFin.Value
                           
                            If CDbl(Me.txtHorasSolicitud.Text) = 0 And CDbl(TxtDiasSolicitud.Text) >= 1 Then 'Significa que son Dias
                                   Me.AdoSolicitud.Recordset("DiasDisfrutar") = CDbl(Me.TxtDiasSolicitud.Text)
                            Else 'Significa que son horas
                                   If CDbl(TxtDiasSolicitud.Text) = 1 And CDbl(txtHorasSolicitud.Text) > 0 Then
                                   Me.AdoSolicitud.Recordset("DiasDisfrutar") = CDbl(Me.txtHorasSolicitud.Text) / Horas
                                   End If
                            End If
                           
                           Me.AdoSolicitud.Recordset("Observaciones") = Me.txtobservaciones.Text
                           Me.AdoSolicitud.Recordset.Update
                           
                           Me.AdoConsecutivo.Recordset("Solicitud") = ConsecutivoSolicitud
                           Me.AdoConsecutivo.Recordset.Update
                        End If
                        
                        
                        '///////////////////////////////////////ELIMINO LOS REGISTROS ////////////////////////////////////
                        rs.Open "DELETE FROM [DescuentoDiasVacaciones] WHERE (NumeroSolicitud = '" & Me.TxtNumero.Text & "') AND (TipoDescuento = '" & Me.CmbTipoSolicitud.Text & "') AND (CodigoEmpleado = '" & Me.DBCodigoEmpleado.Text & "')", Conexion
                        
                        'Supongo que elimina por si existen registros para luego actualizarlos.
                        
                        'Condicionales del DoWhile
                        FechaIni = Format(Me.DtpFechaInicio.Value, "dd/mm/yyyy")
                        FechaFin = Format(Me.DtpFechaFin.Value, "dd/mm/yyyy")
                        
                        
                        Do While FechaIni <= FechaFin
                          If Me.CmbTipoSolicitud.Text = "Vacaciones Pagadas" Then
                          
                               If CDbl(Me.txtHorasSolicitud.Text) = 0 And CDbl(TxtDiasSolicitud.Text) >= 1 Then 'Significa que son Dias
                                   CantDias = DateDiff("d", Format(Me.DtpFechaInicio.Value, "dd/mm/yyyy"), Format(Me.DtpFechaFin.Value, "dd/mm/yyyy")) + 1
                               Else 'Significa que son horas
                                   If CDbl(TxtDiasSolicitud.Text) = 1 And CDbl(txtHorasSolicitud.Text) > 0 Then
                                   CantDias = CDbl(txtHorasSolicitud.Text) / Horas
                                   End If
                               End If
                               Resultado = GrabaDescuentoDias(FechaIni, Me.DBCodigoEmpleado.Text, Me.CmbTipoSolicitud.Text, Me.TxtNumero.Text, CantDias)
                               Exit Do
                          
                           
                              
                          Else
                           CantDias = 1
                           Resultado = GrabaDescuentoDias(FechaIni, Me.DBCodigoEmpleado.Text, Me.CmbTipoSolicitud.Text, Me.TxtNumero.Text, 1)
                          
                          End If
                          FechaIni = DateAdd("d", 1, FechaIni)
                        Loop
                        
                    Me.TxtNumero.Text = Format(CInt(Me.TxtNumero.Text) + 1, "0000#")
                    '///////////////////////////////////////////
                    '///////////////////////////////////////////
                    DoEvents
            AdoAux.Recordset.MoveNext
            Contador = Contador + 1
         Loop
         
            Me.TxtDiasSolicitud.Text = 1
            Me.txtHorasSolicitud.Text = 0
            Me.DBCodigoEmpleado.Text = ""
            Me.CmdGrabar.Enabled = True
            pActualiza = False
            BackToLife
            DisfrutadosConsulta = 0
            Me.DtpFechaInicio.Value = DateTime.Now
            Me.DtpFechaFin.Value = DateTime.Now
            Me.TxtDisfrutados.Text = "0"
            Me.txtDisponibles.Text = "0"
            pbEmpleados.Visible = False
            txtEmpleados.Caption = ""
            txtEmpleados.Visible = False
            
         Exit Sub
    End If
    
    
    '////////////// Si no existen empleados con el Tipo de Nomina no continua ////////////
 Fecha = Format(Me.DtpFechaSolicitud.Value, "yyyy-mm-dd")
 Me.AdoSolicitud.RecordSource = "SELECT  * From SolicitudVacaciones WHERE (NumeroSolicitud = '" & Me.TxtNumero.Text & "')"
 Me.AdoSolicitud.Refresh
 If Me.AdoSolicitud.Recordset.EOF Then
   Me.AdoSolicitud.Recordset.AddNew
   
   If txtdiasvacaciones.Text = "" Then
    txtdiasvacaciones.Text = 0
   End If
   
   If TxtDisfrutados.Text = "" Then
    TxtDisfrutados.Text = 0
   End If
   
    Me.AdoSolicitud.Recordset("FechaSolicitud") = Format(Me.DtpFechaSolicitud.Value, "dd/mm/yyyy")
    Me.AdoSolicitud.Recordset("NumeroSolicitud") = Me.TxtNumero.Text
    Me.AdoSolicitud.Recordset("TipoSolicitud") = Me.CmbTipoSolicitud.Text
    Me.AdoSolicitud.Recordset("CodigoEmpleado") = Me.DBCodigoEmpleado.Text
    Me.AdoSolicitud.Recordset("DiasVacaciones") = Me.txtdiasvacaciones.Text
    Me.AdoSolicitud.Recordset("DiasDisfrutados") = Me.TxtDisfrutados.Text
    Me.AdoSolicitud.Recordset("FechaInicio") = Me.DtpFechaInicio.Value
    Me.AdoSolicitud.Recordset("FechaFin") = Me.DtpFechaFin.Value
    
     If CDbl(Me.txtHorasSolicitud.Text) = 0 And CDbl(TxtDiasSolicitud.Text) >= 1 Then 'Significa que son Dias
            Me.AdoSolicitud.Recordset("DiasDisfrutar") = CDbl(Me.TxtDiasSolicitud.Text)
     Else 'Significa que son horas
            If CDbl(TxtDiasSolicitud.Text) = 1 And CDbl(txtHorasSolicitud.Text) > 0 Then
            Me.AdoSolicitud.Recordset("DiasDisfrutar") = CDbl(Me.txtHorasSolicitud.Text) / Horas
            End If
     End If
    
    Me.AdoSolicitud.Recordset("Observaciones") = Me.txtobservaciones.Text
    Me.AdoSolicitud.Recordset.Update
    
    Me.AdoConsecutivo.Recordset("Solicitud") = ConsecutivoSolicitud
    Me.AdoConsecutivo.Recordset.Update
 Else
 
  If txtdiasvacaciones.Text = "" Then
    txtdiasvacaciones.Text = 0
   End If
   
   If TxtDisfrutados.Text = "" Then
    TxtDisfrutados.Text = 0
   End If
 
    Me.AdoSolicitud.Recordset("DiasVacaciones") = Me.txtdiasvacaciones.Text
    Me.AdoSolicitud.Recordset("DiasDisfrutados") = Me.TxtDisfrutados.Text
    Me.AdoSolicitud.Recordset("FechaInicio") = Me.DtpFechaInicio.Value
    Me.AdoSolicitud.Recordset("FechaFin") = Me.DtpFechaFin.Value
    Me.AdoSolicitud.Recordset("TipoSolicitud") = Me.CmbTipoSolicitud.Text
    
    If CDbl(Me.txtHorasSolicitud.Text) = 0 And CDbl(TxtDiasSolicitud.Text) >= 1 Then 'Significa que son Dias
            Me.AdoSolicitud.Recordset("DiasDisfrutar") = CDbl(Me.TxtDiasSolicitud.Text)
     Else 'Significa que son horas
            If CDbl(TxtDiasSolicitud.Text) = 1 And CDbl(txtHorasSolicitud.Text) > 0 Then
            Me.AdoSolicitud.Recordset("DiasDisfrutar") = CDbl(Me.txtHorasSolicitud.Text) / Horas
            End If
     End If
    
    Me.AdoSolicitud.Recordset("Observaciones") = Me.txtobservaciones.Text
    Me.AdoSolicitud.Recordset.Update
 End If
 
 
 '///////////////////////////////////////ELIMINO LOS REGISTROS ////////////////////////////////////
 rs.Open "DELETE FROM [DescuentoDiasVacaciones] WHERE (NumeroSolicitud = '" & Me.TxtNumero.Text & "') AND (TipoDescuento = '" & Me.CmbTipoSolicitud.Text & "') AND (CodigoEmpleado = '" & Me.DBCodigoEmpleado.Text & "')", Conexion
 
 'Supongo que elimina por si existen registros para luego actualizarlos.
 
 'Condicionales del DoWhile
 FechaIni = Format(Me.DtpFechaInicio.Value, "dd/mm/yyyy")
 FechaFin = Format(Me.DtpFechaFin.Value, "dd/mm/yyyy")
 
 
 Do While FechaIni <= FechaFin
   If Me.CmbTipoSolicitud.Text = "Vacaciones Pagadas" Then
   
        If CDbl(Me.txtHorasSolicitud.Text) = 0 And CDbl(TxtDiasSolicitud.Text) >= 1 Then 'Significa que son Dias
            CantDias = DateDiff("d", Format(Me.DtpFechaInicio.Value, "dd/mm/yyyy"), Format(Me.DtpFechaFin.Value, "dd/mm/yyyy")) + 1
        Else 'Significa que son horas
            If CDbl(TxtDiasSolicitud.Text) = 1 And CDbl(txtHorasSolicitud.Text) > 0 Then
            CantDias = CDbl(txtHorasSolicitud.Text) / Horas
            End If
        End If
        Resultado = GrabaDescuentoDias(FechaIni, Me.DBCodigoEmpleado.Text, Me.CmbTipoSolicitud.Text, Me.TxtNumero.Text, CantDias)
        Exit Do
   
    
       
   Else
    CantDias = 1
    Resultado = GrabaDescuentoDias(FechaIni, Me.DBCodigoEmpleado.Text, Me.CmbTipoSolicitud.Text, Me.TxtNumero.Text, 1)
   
   End If
   FechaIni = DateAdd("d", 1, FechaIni)
 Loop
 Me.TxtDiasSolicitud.Text = 1
 Me.txtHorasSolicitud.Text = 0
 Me.DBCodigoEmpleado.Text = ""
 Me.CmdGrabar.Enabled = True
 pActualiza = False
 BackToLife
 DisfrutadosConsulta = 0
 Me.DtpFechaInicio.Value = DateTime.Now
 Me.DtpFechaFin.Value = DateTime.Now
 
 Me.TxtDisfrutados.Text = "0"
 Me.txtDisponibles.Text = "0"
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
    QueProducto = "EmpleadosSoli"
    FrmConsulta.Show
    'QueProducto = ""
End Sub

Private Sub DBCodigoEmpleado_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    Dim DiasAcumulados As Double, FechaInicio As Date, DiasDisfrute As Double
           Me.DtaEmpleado.RecordSource = "SELECT Empleado.CodEmpleado, Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Empleado.NumHijos, Empleado.Direccion, Empleado.Nacionalidad, Empleado.CodigoPostal, Empleado.Sexo, Empleado.CodInss, Empleado.CodIr,Empleado.NumCedula, Empleado.Sindicalista, Empleado.CodDepartamento, Empleado.CodGrupo, Empleado.CodCargo, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodTipoNomina, Empleado.DiasDescuento, Empleado.SueldoPeriodo, Empleado.TarifaHoraria, Empleado.OtrosIngresos, Empleado.PorcentajeComision, Empleado.DescripOtrIngre, Empleado.ExentoInss, Empleado.ExentoIr, Empleado.PagoInssPatronal, Empleado.SalarioMinimo, Empleado.Observaciones, Empleado.Activo, Empleado.Ausente, Empleado.SalarioFijo, Empleado.SumarSubsidio, Empleado.PorcientoIncentivo, Empleado.Dolarizado, Empleado.CuentaBanco, Empleado.SueldoActualBasico, Empleado.HorasTurno, Departamento.Departamento, Cargo.Cargo, " & _
                                          "Historico.FechaContrato FROM Empleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN  Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE (Empleado.CodEmpleado1 = '" & DBCodigoEmpleado.Text & "') AND (Empleado.Activo = 1)"
            Me.DtaEmpleado.Refresh

 
 
 If Not Me.DtaEmpleado.Recordset.EOF Then
        
     Me.txtNombre.Text = Me.DtaEmpleado.Recordset("Nombres")
     Me.txtDepartamento.Text = Me.DtaEmpleado.Recordset("Departamento")
     Me.DBCCargo.Text = Me.DtaEmpleado.Recordset("Cargo")
    Me.DtpFechaIngreso.Value = Me.DtaEmpleado.Recordset("FechaContrato")
     FechaInicio = Format(Me.DtpFechaInicio.Value, "dd/mm/yyyy")
     
     
     'DiasVacaAcumulados = CalcularDiasVaca(CDate(Me.DtpFechaIngreso.Value), FechaInicio) - DiasVacaDesAcumulados(Me.DBCodigoEmpleado.Text, FechaInicio)
    ' Me.txtdiasvacaciones.Text = DiasVacaAcumulados
     'DiasDisfrute = DiasDifrutados(Me.DBCodigoEmpleado.Text, Me.DtpFechaInicio.Value, Me.DtpFechaFin.Value)
    ' Me.TxtDisfrutados.Text = DiasDisfrute
    ' Me.TxtDisponibles.Text = DiasVacaAcumulados - DiasDisfrute
    
    
    '//////////////////////////////////////////////////////////////////////////////////////////////
    
    
   
    TotalVacacionesDisfrutadas = 0
    TotalDiasVacaciones = 0
    TotalDiasDisponibles = 0
    
    
    Dim fs As Boolean
    fs = False
'//////////////////////////////// Saco datos generales del empleado ///////////////////////
MDIPrimero.DtaConsulta.RecordSource = "SELECT     TOP (1) Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombre, Historico.FechaContratoVac, DATEADD(month,    (YEAR(Historico.FechaContratoVac) - 1900) * 12 + MONTH(Historico.FechaContratoVac), - 1) AS UdMes  FROM         Empleado INNER JOIN   Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (Empleado.CodEmpleado1 = '" & DBCodigoEmpleado.Text & "') And Empleado.Activo = 'True'"
MDIPrimero.DtaConsulta.Refresh

'///////// Inicializo parametros generales ////////////
Dim VacacionesAcumuladas, VacacionesSolicitadas, SaldoActual, tempVacacionSolicitada, tempVacacionAcumulada As Double
Dim NombreCompleto As String


NombreCompleto = MDIPrimero.DtaConsulta.Recordset("Nombre")
Dim Inicio, tempInicio, Fin, tempFin As Date

pFechaContrato = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
Inicio = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
tempInicio = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
Me.DtpFechaIngreso.Value = MDIPrimero.DtaConsulta.Recordset("FechaContratoVac")
Fin = DateTime.Now
tempFin = MDIPrimero.DtaConsulta.Recordset("udMes")


         SaldoActual = 0
         
         Dim Tipo As String
         MDIPrimero.DtaControles.Refresh
         Tipo = MDIPrimero.DtaControles.Recordset("DiasMes")


    Do While (Inicio < Fin)
             tempVacacionSolicitada = 0
             tempVacacionAcumulada = 0
             
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
            
             TotalDiasVacaciones = TotalDiasVacaciones + tempVacacionAcumulada
             
             '////////     /////////     ////////        /////////       /////////       ////////        //////
             '////////// Calculo el total de dias y horas solicitadas en el rango de fechas recorrido /////////
             '/////////////////////////////////////////////////////////////////////////////////////////////////
           
           If DiasMes = 31 Then
                AdoAuxiliar.RecordSource = "SELECT     SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM         SolicitudVacaciones WHERE   not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'   and    (CodigoEmpleado = '" & Me.DBCodigoEmpleado.Text & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(DateAdd("d", 1, tempFin), "dd/MM/yyyy") & " 23:59')"
           Else
                AdoAuxiliar.RecordSource = "SELECT     SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM         SolicitudVacaciones Where   not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'   and    (CodigoEmpleado = '" & Me.DBCodigoEmpleado.Text & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(tempFin, "dd/MM/yyyy") & " 23:59')"
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
                AdoAuxiliar.RecordSource = "SELECT     SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM SolicitudVacaciones WHERE   not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'       and    (CodigoEmpleado = '" & Me.DBCodigoEmpleado.Text & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(DateAdd("d", 1, tempFin), "dd/MM/yyyy") & " 23:59')"
           Else
                AdoAuxiliar.RecordSource = "SELECT     SUM(DiasDisfrutar) AS VacacionesSolicitadas  FROM SolicitudVacaciones WHERE   not TipoSolicitud = 'Ausente' and not TipoSolicitud = 'Subsidio' and not TipoSolicitud = 'Suspension' and not TipoSolicitud = 'Feriado'      and     (CodigoEmpleado = '" & Me.DBCodigoEmpleado.Text & "' or CodigoEmpleado = 'Todos') AND (FechaInicio >= '" & Format(tempInicio, "dd/MM/yyyy") & " 00:00') AND (FechaInicio <= '" & Format(tempFin, "dd/MM/yyyy") & " 23:59')"
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
    
        Me.txtdiasvacaciones.Text = TotalDiasVacaciones
     
     If CDbl(txtHorasSolicitud.Text) > 0 Then
        Dim Horas As Double
        AdoNomina.RecordSource = "SELECT TipoNomina.Horas FROM TipoNomina INNER JOIN  Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina WHERE        (Empleado.CodEmpleado1 = '" & Me.DBCodigoEmpleado.Text & "')"
        AdoNomina.Refresh
        
        If Not Me.AdoNomina.Recordset.EOF Then
                Horas = Me.AdoNomina.Recordset("Horas")
            Else
                Horas = 0
        End If
        
        
     Else
     
     End If
     
     If pActualiza = True Then
        Me.TxtDisfrutados.Text = TotalVacacionesDisfrutadas
        Me.txtDisponibles.Text = TotalDiasDisponibles
        Me.txtdiasvacaciones.Text = TotalDiasVacaciones
     Else
        Me.TxtDisfrutados.Text = Format(TotalVacacionesDisfrutadas + 1, "##,##0.00")
        Me.txtDisponibles.Text = Format(TotalDiasDisponibles - 1, "##,##0.00")
        Me.txtdiasvacaciones.Text = Format(TotalDiasVacaciones, "##,##0.00")
     End If
     
     pTotalDiasVacaciones = TotalDiasVacaciones
     pTotalVacacionesDisfrutadas = TotalVacacionesDisfrutadas
     pTotalDiasDisponibles = TotalDiasDisponibles
     
     
 End If
 End If
End Sub

Private Sub DBCodigoEmpleado_SelChange(Cancel As Integer)
    DBCodigoEmpleado_KeyPress (13)
End Sub

Private Sub DtpFechaFin_Change()
    Me.TxtDiasSolicitud.Text = DateDiff("d", Format(Me.DtpFechaInicio.Value, "dd/mm/yyyy"), Format(Me.DtpFechaFin.Value, "dd/mm/yyyy")) + 1

   If val(Me.TxtDiasSolicitud.Text) > 1 Then
       Me.TxtDiasSolicitud.Locked = False
        Me.txtHorasSolicitud.Enabled = True
   Else
       Me.TxtDiasSolicitud.Locked = True
       Me.txtHorasSolicitud.Enabled = True
   End If

End Sub

Private Sub DtpFechaInicio_Change()
 Me.TxtDiasSolicitud.Text = DateDiff("d", Format(Me.DtpFechaInicio.Value, "dd/mm/yyyy"), Format(Me.DtpFechaFin.Value, "dd/mm/yyyy")) + 1

   If val(Me.TxtDiasSolicitud.Text) > 1 Then
       Me.TxtDiasSolicitud.Locked = False
        Me.txtHorasSolicitud.Enabled = True
   Else
       Me.TxtDiasSolicitud.Locked = True
       Me.txtHorasSolicitud.Enabled = True
   End If
End Sub

Private Sub Form_Load()
pActualiza = False
DisfrutadosConsulta = 0
Me.txtHorasSolicitud.Text = "00"
'MDIPrimero.Skin1.ApplySkin hWnd
Me.BackColor = RGB(222, 227, 247)
Me.GroupBox1.BackColor = RGB(222, 227, 247)
Me.GroupBox2.BackColor = RGB(222, 227, 247)
Me.GroupBox3.BackColor = RGB(222, 227, 247)
Me.Label1.BackColor = RGB(222, 227, 247)
Me.Label2.BackColor = RGB(222, 227, 247)

Me.DtpFechaIngreso.Value = Now
Me.DtpFechaSolicitud.Value = Now
Me.DtpFechaInicio.Value = Now
Me.DtpFechaFin.Value = Now
Me.TxtDiasSolicitud.Text = DateDiff("d", Format(Me.DtpFechaInicio.Value, "dd/mm/yyyy"), Format(Me.DtpFechaFin.Value, "dd/mm/yyyy")) + 1
Me.TxtNumero.Text = Format(ConsecutivoSolicitud, "0000#")

Me.CmbTipoSolicitud.AddItem "Vacaciones"
Me.CmbTipoSolicitud.AddItem "Vacaciones Pagadas"
Me.CmbTipoSolicitud.AddItem "Vacaciones Programadas"
Me.CmbTipoSolicitud.AddItem "Subsidio"
Me.CmbTipoSolicitud.AddItem "Ausente"
Me.CmbTipoSolicitud.AddItem "Feriado"
Me.CmbTipoSolicitud.AddItem "Permiso Programado"
Me.CmbTipoSolicitud.AddItem "Suspension"
Me.CmbTipoSolicitud.Text = "Vacaciones"
Me.CmbClasificado.AddItem "Riesgo Laboral"
Me.CmbClasificado.AddItem "Accidente Comun"
Me.CmbClasificado.AddItem "Embarazo"
Me.CmbClasificado.AddItem "Enfermedades"



With Me.DtaEmpleados
   .ConnectionString = Conexion
   .RecordSource = "SELECT CodEmpleado, CodEmpleado1, Activo, Nombre1 + ' '+ Nombre2 +' '+Apellido1+' '+Apellido2 as Nombres From Empleado Where (Activo = 1) ORDER BY CodEmpleado1"
   .Refresh
End With

With Me.AdoConsecutivo
   .ConnectionString = Conexion
   .RecordSource = "SELECT  * From Consecutivos"
   .Refresh
End With



With Me.AdoAux
   .ConnectionString = Conexion
End With

With Me.AdoNomina
   .ConnectionString = Conexion
End With
With Me.AdoAuxiliar
   .ConnectionString = Conexion
End With

With Me.DtaEmpleado
   .ConnectionString = Conexion
End With

With Me.AdoSolicitud
   .ConnectionString = Conexion
End With

With Me.AdoConsulta
   .ConnectionString = Conexion
End With

Me.DBCodigoEmpleado.Columns(0).Visible = False
Me.DBCodigoEmpleado.Columns(1).Caption = "Codigo"
Me.DBCodigoEmpleado.Columns(1).Width = 800
Me.DBCodigoEmpleado.Columns(2).Visible = False


End Sub
Public Sub BackToLife()
Me.TxtDisfrutados.Visible = True
Me.txtDisponibles.Visible = True
Me.txtdiasvacaciones.Visible = True
    Me.CmbTipoSolicitud.Text = "Vacaciones"
    Me.DBCCargo.Text = ""
    Me.txtNombre.Text = ""
    Me.DtpFechaIngreso.Value = DateTime.Now
    Me.txtDepartamento.Text = ""
    Me.DBCCargo.Text = ""
    Me.txtdiasvacaciones.Text = ""
    Me.TxtDisfrutados.Text = ""
    Me.txtDisponibles.Text = ""
    Me.DtpFechaInicio.Value = DateTime.Now
    Me.DtpFechaFin.Value = DateTime.Now
    Me.txtHorasSolicitud.Text = "00"
    Me.CmbClasificado.Text = ""
    Me.txtobservaciones.Text = ""
    Me.TxtNumero.Text = Format(ConsecutivoSolicitud, "0000#")
End Sub



Private Sub PushButton1_Click()
pActualiza = True
Dim Fecha As String
QueProducto = "Solicitud"
FrmConsulta.Show 1


If FrmConsulta.NumeroSolicitud = "" Then
Exit Sub
End If


Me.TxtNumero.Text = FrmConsulta.NumeroSolicitud
Me.DtpFechaSolicitud.Value = FrmConsulta.FechaSolicitud
Me.CmbTipoSolicitud.Text = FrmConsulta.TipoSolicitud
Me.DBCodigoEmpleado.Text = FrmConsulta.CodigoEmpleado1
If FrmConsulta.TipoSolicitud = "Ausente" Or FrmConsulta.TipoSolicitud = "Suspension" Then
    Me.TxtDisfrutados.Visible = False
    Me.txtDisponibles.Visible = False
    Me.txtdiasvacaciones.Visible = False
Else
    Me.TxtDisfrutados.Visible = True
    Me.txtDisponibles.Visible = True
    Me.txtdiasvacaciones.Visible = True
End If


'-----------------------BUSCO CANTIDAD DE HORAS DEL EMPLEADO ---------------------------
AdoNomina.RecordSource = "SELECT TipoNomina.Horas FROM TipoNomina INNER JOIN  Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina WHERE        (Empleado.CodEmpleado1 = '" & Me.DBCodigoEmpleado.Text & "')"
AdoNomina.Refresh
    
    If Not Me.AdoNomina.Recordset.EOF Then
    Horas = Me.AdoNomina.Recordset("Horas")
    Else
    Horas = 0
    End If

'-----------------------BUSCO LA INFORMACION RESTANTE DE LA SOLICITUD -------------------
Fecha = Format(Me.DtpFechaSolicitud.Value, "yyyy-mm-dd")
AdoConsulta.RecordSource = "SELECT  * From SolicitudVacaciones WHERE (NumeroSolicitud = '" & Me.TxtNumero.Text & "') "
AdoConsulta.Refresh
If Not AdoConsulta.Recordset.EOF Then
  Me.DtpFechaInicio.Value = AdoConsulta.Recordset("FechaInicio")
  Me.DtpFechaFin.Value = AdoConsulta.Recordset("FechaFin")
  Me.txtobservaciones.Text = AdoConsulta.Recordset("Observaciones")
    If AdoConsulta.Recordset("DiasDisfrutar") < 1 Then
        txtHorasSolicitud.Text = AdoConsulta.Recordset("DiasDisfrutar") * Horas
        TxtDiasSolicitud.Text = "1"
    Else
        txtHorasSolicitud.Text = "00"
    End If
    
    If Not IsNull(AdoConsulta.Recordset("DiasDisfrutar")) Then
        DisfrutadosConsulta = AdoConsulta.Recordset("DiasDisfrutar")
    End If
    
  DtpFechaInicio_Change
  DBCodigoEmpleado_KeyPress (13)
  
'  Me.TxtDiasVacaciones.Text = AdoConsulta.Recordset("DiasVacaciones")
'  Me.TxtDisfrutados.Text = AdoConsulta.Recordset("DiasDisfrutados")
End If



End Sub

Private Sub PushButton2_Click()


k% = MsgBox("Desea eliminar la solicitud?", vbYesNo)
If k% <> 6 Then
    Cancel = 1
    Exit Sub
End If

Dim strSQL As String
Dim rs As New ADODB.Recordset
 strSQL = "DELETE FROM SolicitudVacaciones WHERE     (NumeroSolicitud = '" & Me.TxtNumero.Text & "')"
 rs.Open strSQL, Conexion
 
 Set rs = New ADODB.Recordset
 strSQL = "DELETE FROM DescuentoDiasVacaciones WHERE     (NumeroSolicitud = '" & Me.TxtNumero.Text & "')"
 rs.Open strSQL, Conexion
 BackToLife
 pActualiza = False
 
End Sub

Private Sub TxtDiasSolicitud_Change()

If Me.TxtDiasSolicitud.Text = "" Then
  Exit Sub
End If

If Me.TxtDiasSolicitud.Text = 1 Then
     Me.TxtDiasSolicitud.Text = DateDiff("d", Format(Me.DtpFechaInicio.Value, "dd/mm/yyyy"), Format(Me.DtpFechaFin.Value, "dd/mm/yyyy")) + 1

   If val(Me.TxtDiasSolicitud.Text) > 1 Then
       Me.TxtDiasSolicitud.Locked = False
   Else
       Me.TxtDiasSolicitud.Locked = True
   End If
End If

If TxtDiasSolicitud.Text > 1 Then
TxtDiasSolicitud.Enabled = True
   Me.txtHorasSolicitud.Enabled = False
   Me.txtHorasSolicitud.Text = "00"
   
   Me.txtDisponibles.Text = Format((pTotalDiasDisponibles + DisfrutadosConsulta) - TxtDiasSolicitud, "##,##0.00")
   Me.TxtDisfrutados.Text = Format((pTotalVacacionesDisfrutadas - DisfrutadosConsulta) + TxtDiasSolicitud, "##,##0.00")
   
Else

    If CDbl(TxtDiasSolicitud.Text) = 1 And CDbl(txtHorasSolicitud.Text) = 0 Then
     Me.txtDisponibles.Text = Format((pTotalDiasDisponibles + DisfrutadosConsulta) - TxtDiasSolicitud, "##,##0.00")
     Me.TxtDisfrutados.Text = Format((pTotalVacacionesDisfrutadas - DisfrutadosConsulta) + TxtDiasSolicitud, "##,##0.00")
    End If
    
    If TxtDiasSolicitud.Text = 1 Then
        TxtDiasSolicitud.Enabled = True
        txtHorasSolicitud.Enabled = True
    Else
        TxtDiasSolicitud.Enabled = False
    End If
End If



End Sub
 
Private Sub txtHorasSolicitud_Change()
Dim Horas As Double

If txtHorasSolicitud.Text = "" Then
    txtHorasSolicitud.Text = 0
End If

If CDbl(txtHorasSolicitud.Text) > 0 Then
    AdoNomina.RecordSource = "SELECT TipoNomina.Horas FROM TipoNomina INNER JOIN  Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina WHERE        (Empleado.CodEmpleado1 = '" & Me.DBCodigoEmpleado.Text & "')"
    AdoNomina.Refresh
    
    If Not Me.AdoNomina.Recordset.EOF Then
            Horas = Me.AdoNomina.Recordset("Horas")
        Else
            Horas = 0
    End If
    
    Me.txtDisponibles.Text = Format((pTotalDiasDisponibles - DisfrutadosConsulta) - (CDbl(Me.TxtDiasSolicitud.Text) + (CDbl(Me.txtHorasSolicitud.Text) / Horas)), "##,##0.00")
    Me.TxtDisfrutados.Text = Format((pTotalVacacionesDisfrutadas + DisfrutadosConsulta) + CDbl(Me.TxtDiasSolicitud.Text) + (CDbl(Me.txtHorasSolicitud.Text) / Horas), "##,##0.00")
Else
     Me.txtDisponibles.Text = Format(pTotalDiasDisponibles - 1, "##,##0.00")
     Me.TxtDisfrutados.Text = Format(pTotalVacacionesDisfrutadas + 1, "##,##0.00")
End If

End Sub

Private Sub txtHorasSolicitud_Click()
'txtHorasSolicitud.SelStart = 0
'txtHorasSolicitud.SelLength = Len(txtHorasSolicitud.Text)
End Sub
