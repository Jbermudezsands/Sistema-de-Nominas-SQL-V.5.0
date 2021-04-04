VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmReferencias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Referencias"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   586
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   21
      Top             =   4560
      Width           =   3135
      Begin XtremeSuiteControls.PushButton CmdAnterior 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Anterior"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmReferencias.frx":0000
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdSiguiente 
         Height          =   375
         Left            =   1560
         TabIndex        =   23
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Siguiente"
         ForeColor       =   0
         TextAlignment   =   0
         Appearance      =   6
         Picture         =   "FrmReferencias.frx":0502
         ImageAlignment  =   1
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton CmdPrimero 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Primero"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmReferencias.frx":0A06
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdUltimo 
         Height          =   375
         Left            =   1560
         TabIndex        =   25
         Top             =   720
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Siguiente"
         ForeColor       =   0
         TextAlignment   =   0
         Appearance      =   6
         Picture         =   "FrmReferencias.frx":0F08
         ImageAlignment  =   1
         TextImageRelation=   4
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
      ScaleWidth      =   8775
      TabIndex        =   14
      Top             =   0
      Width           =   8775
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTROS DE REFERENCIAS"
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
         Left            =   2640
         TabIndex        =   15
         Top             =   360
         Width           =   3840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   8760
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image2 
         Height          =   1020
         Left            =   120
         Picture         =   "FrmReferencias.frx":140A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1290
      End
   End
   Begin MSAdodcLib.Adodc AdoFGProcesos 
      Height          =   375
      Left            =   3720
      Top             =   6360
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
      Caption         =   "AdoFGProcesos"
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
      Top             =   6600
      Width           =   2415
      _ExtentX        =   4260
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
      Top             =   6360
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
   Begin MSAdodcLib.Adodc AdoMetas 
      Height          =   375
      Left            =   3720
      Top             =   6360
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
      Caption         =   "AdoMetas"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5741
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Referencias"
      TabPicture(0)   =   "FrmReferencias.frx":1EC8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DBCodigoReferencia"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TxtMeta"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TxtProduccion"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TxtReferencia"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TxtDescripcion"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TxtDiferencia"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TxtCodReferencia"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ChkActivo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtCodReferencia2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "FG-Procesos"
      TabPicture(1)   =   "FrmReferencias.frx":1EE4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CmdPrnNomina"
      Tab(1).Control(1)=   "CmdPrnDetalle"
      Tab(1).Control(2)=   "TDBGridProcesos"
      Tab(1).ControlCount=   3
      Begin VB.TextBox TxtCodReferencia2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         TabIndex        =   20
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   6840
         TabIndex        =   18
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox ChkActivo 
         Caption         =   "Activo"
         Height          =   375
         Left            =   6840
         TabIndex        =   17
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtCodReferencia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         TabIndex        =   16
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtDiferencia 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1680
         Width           =   7455
      End
      Begin VB.TextBox TxtReferencia 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox TxtProduccion 
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TxtMeta 
         Height          =   315
         Left            =   600
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   2640
         Width           =   1335
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGridProcesos 
         Bindings        =   "FrmReferencias.frx":1F00
         Height          =   2175
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   3836
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "CodReferencia"
         Columns(0).DataField=   "CodReferencia"
         Columns(0).DataWidth=   5
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "CodProceso"
         Columns(1).DataField=   "CodProceso"
         Columns(1).DataWidth=   5
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descrip"
         Columns(2).DataField=   "Descrip"
         Columns(2).DataWidth=   100
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Lunes"
         Columns(3).DataField=   "Lunes"
         Columns(3).DataWidth=   22
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Martes"
         Columns(4).DataField=   "Martes"
         Columns(4).DataWidth=   22
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Miercoles"
         Columns(5).DataField=   "Miercoles"
         Columns(5).DataWidth=   22
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Jueves"
         Columns(6).DataField=   "Jueves"
         Columns(6).DataWidth=   22
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Viernes"
         Columns(7).DataField=   "Viernes"
         Columns(7).DataWidth=   22
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Sabado"
         Columns(8).DataField=   "Sabado"
         Columns(8).DataWidth=   22
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Domingo"
         Columns(9).DataField=   "Domingo"
         Columns(9).DataWidth=   22
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "TotalUnidades"
         Columns(10).DataField=   "TotalUnidades"
         Columns(10).DataWidth=   23
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Produccion"
         Columns(11).DataField=   "Produccion"
         Columns(11).DataWidth=   22
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "SalarioPieza"
         Columns(12).DataField=   "SalarioPieza"
         Columns(12).DataWidth=   22
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Meta"
         Columns(13).DataField=   "Meta"
         Columns(13).DataWidth=   23
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Diferencia"
         Columns(14).DataField=   "Diferencia"
         Columns(14).DataWidth=   23
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   15
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=15"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2090"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
         Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1746"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1667"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=5292"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=5212"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(17)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(19)=   "Column(3)._AlignLeft=0"
         Splits(0)._ColumnProps(20)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(23)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(25)=   "Column(4)._AlignLeft=0"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(29)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(5)._AlignLeft=0"
         Splits(0)._ColumnProps(32)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(35)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(37)=   "Column(6)._AlignLeft=0"
         Splits(0)._ColumnProps(38)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(41)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(42)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(43)=   "Column(7)._AlignLeft=0"
         Splits(0)._ColumnProps(44)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(47)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(49)=   "Column(8)._AlignLeft=0"
         Splits(0)._ColumnProps(50)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(51)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(52)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(53)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(54)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(55)=   "Column(9)._AlignLeft=0"
         Splits(0)._ColumnProps(56)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(57)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(58)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(59)=   "Column(10).Visible=0"
         Splits(0)._ColumnProps(60)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(61)=   "Column(10)._AlignLeft=0"
         Splits(0)._ColumnProps(62)=   "Column(11).Width=1773"
         Splits(0)._ColumnProps(63)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(64)=   "Column(11)._WidthInPix=1693"
         Splits(0)._ColumnProps(65)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(66)=   "Column(11)._AlignLeft=0"
         Splits(0)._ColumnProps(67)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(68)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(69)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(70)=   "Column(12).Visible=0"
         Splits(0)._ColumnProps(71)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(72)=   "Column(12)._AlignLeft=0"
         Splits(0)._ColumnProps(73)=   "Column(13).Width=1773"
         Splits(0)._ColumnProps(74)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(75)=   "Column(13)._WidthInPix=1693"
         Splits(0)._ColumnProps(76)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(77)=   "Column(13)._AlignLeft=0"
         Splits(0)._ColumnProps(78)=   "Column(14).Width=1773"
         Splits(0)._ColumnProps(79)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(80)=   "Column(14)._WidthInPix=1693"
         Splits(0)._ColumnProps(81)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(82)=   "Column(14)._AlignLeft=0"
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
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
         _StyleDefs(71)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
         _StyleDefs(74)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
         _StyleDefs(75)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
         _StyleDefs(76)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
         _StyleDefs(78)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
         _StyleDefs(79)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
         _StyleDefs(80)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
         _StyleDefs(81)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
         _StyleDefs(82)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
         _StyleDefs(83)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
         _StyleDefs(84)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
         _StyleDefs(85)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
         _StyleDefs(86)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
         _StyleDefs(87)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
         _StyleDefs(88)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
         _StyleDefs(89)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
         _StyleDefs(90)  =   "Named:id=33:Normal"
         _StyleDefs(91)  =   ":id=33,.parent=0"
         _StyleDefs(92)  =   "Named:id=34:Heading"
         _StyleDefs(93)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(94)  =   ":id=34,.wraptext=-1"
         _StyleDefs(95)  =   "Named:id=35:Footing"
         _StyleDefs(96)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(97)  =   "Named:id=36:Selected"
         _StyleDefs(98)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(99)  =   "Named:id=37:Caption"
         _StyleDefs(100) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(101) =   "Named:id=38:HighlightRow"
         _StyleDefs(102) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(103) =   "Named:id=39:EvenRow"
         _StyleDefs(104) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(105) =   "Named:id=40:OddRow"
         _StyleDefs(106) =   ":id=40,.parent=33"
         _StyleDefs(107) =   "Named:id=41:RecordSelector"
         _StyleDefs(108) =   ":id=41,.parent=34"
         _StyleDefs(109) =   "Named:id=42:FilterBar"
         _StyleDefs(110) =   ":id=42,.parent=33"
      End
      Begin TrueOleDBList80.TDBCombo DBCodigoReferencia 
         Bindings        =   "FrmReferencias.frx":1F1C
         Height          =   315
         Left            =   1560
         TabIndex        =   19
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
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
         _PropDict       =   $"FrmReferencias.frx":1F38
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
      Begin XtremeSuiteControls.PushButton CmdPrnDetalle 
         Height          =   375
         Left            =   -69840
         TabIndex        =   30
         Top             =   2760
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Detalle"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmReferencias.frx":1FE2
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdPrnNomina 
         Height          =   375
         Left            =   -68280
         TabIndex        =   31
         Top             =   2760
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Resumen"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmReferencias.frx":42CE
         ImageAlignment  =   0
      End
      Begin VB.Label Label7 
         Caption         =   "Diferencia:"
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Produccion:"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Meta:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion Referencia"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Referencia"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo Referencia"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.PushButton CmdGrabar 
      Height          =   375
      Left            =   3480
      TabIndex        =   26
      Top             =   4680
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Grabar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmReferencias.frx":65BA
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdBorrar 
      Height          =   375
      Left            =   3480
      TabIndex        =   27
      Top             =   5160
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Borrar"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmReferencias.frx":891E
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   7200
      TabIndex        =   28
      Top             =   5280
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmReferencias.frx":8DD2
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton CmdNuevo 
      Height          =   375
      Left            =   5280
      TabIndex        =   29
      Top             =   4680
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Nuevo"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmReferencias.frx":92D6
      ImageAlignment  =   0
   End
   Begin VB.Label Label5 
      Caption         =   "Meta:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   615
   End
End
Attribute VB_Name = "FrmReferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnterior_Click()
 Me.DtaReferencia.Recordset.MovePrevious
 If Not Me.DtaReferencia.Recordset.BOF Then
   Me.DBCodigoReferencia.Text = Me.DtaReferencia.Recordset("CodReferencia")
 Else
   MsgBox "Este es el Primer Registro", vbExclamation, "sistema de Nominas"
   Me.DtaReferencia.Recordset.MoveNext
 End If
End Sub

Private Sub cmdborrar_Click()
On Error GoTo TipoErrs
Dim Respuesta, Rsp
If IsNull(Me.DtaReferencia.Recordset("CodReferencia")) = True Then
        MsgBox "No Existe Registro"
        Exit Sub
End If
If Me.DBCodigoReferencia.Text = "" Then
 MsgBox "No se Puede Eliminar este Registro", vbInformation, "Sistema de Nominas"
 Exit Sub
End If
'Elimino el registro activo en la pantalla

  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando la Referencia: " & Me.TxtDescripcion.Text)
   If Respuesta = 6 Then
     Me.DtaConsulta.RecordSource = "SELECT Ref, CodReferencia, Descrip From Referencia WHERE     (CodReferencia = " & Me.DBCodigoReferencia.Columns(1).Text & ")"
     Me.DtaConsulta.Refresh
     If Not Me.DtaConsulta.Recordset.EOF Then
       Me.DtaConsulta.Recordset.Delete
     End If
     If IsNull(Me.DtaReferencia.Recordset("CodReferencia")) = True Then
        MsgBox "No Existe Registro"
        Exit Sub
     Else
      Me.DBCodigoReferencia.Text = ""
      Me.DtaReferencia.Refresh
      CmdAnterior.Enabled = True
     End If
 End If
Exit Sub
TipoErrs:
   ControlErrores
   MsgBox Err.Description
 Unload Me
End Sub

Private Sub CmdGrabar_Click()
'On Error GoTo TipoErrs
 Me.DtaConsulta.RecordSource = "SELECT Activo,Meta, Ref, CodReferencia1, Descrip From Referencia WHERE  (CodReferencia1 = '" & Me.DBCodigoReferencia.Text & "') and (CodReferencia = '" & val(Me.TxtCodReferencia.Text) & "') "    'And Activo=1
 Me.DtaConsulta.Refresh
    If Me.TxtMeta.Text = "" Then
      Me.TxtMeta.Text = 0
    End If
 If Not Me.DtaConsulta.Recordset.EOF Then

    Me.DtaConsulta.Recordset("Ref") = Me.TxtReferencia.Text
    Me.DtaConsulta.Recordset("Descrip") = Me.TxtDescripcion.Text
    Me.DtaConsulta.Recordset("Meta") = CDbl(Me.TxtMeta.Text)
    If Me.ChkActivo.Value = 1 Then
     Me.DtaConsulta.Recordset("Activo") = 1
    Else
     Me.DtaConsulta.Recordset("Activo") = 0
    End If
    Me.DtaConsulta.Recordset.Update
 Else
    Me.DtaConsulta.Refresh
    Me.DtaConsulta.Recordset.AddNew
    Me.DtaConsulta.Recordset("CodReferencia1") = Me.DBCodigoReferencia.Text
    Me.DtaConsulta.Recordset("Ref") = Me.TxtReferencia.Text
    Me.DtaConsulta.Recordset("Descrip") = Me.TxtDescripcion.Text
    Me.DtaConsulta.Recordset("Meta") = CDbl(Me.TxtMeta.Text)
    Me.DtaConsulta.Recordset("Activo") = 1
    Me.DtaConsulta.Recordset.Update
    
    
 End If
 Me.TxtCodReferencia.Text = ""
 Me.DBCodigoReferencia.Text = ""
 Me.TxtReferencia.Text = ""
 Me.TxtCodReferencia2.Text = ""
 Me.TxtDescripcion.Text = ""
 Me.TxtDiferencia.Text = ""
 Me.TxtMeta.Text = ""
 Me.TxtProduccion.Text = ""

 Me.DtaReferencia.Refresh
Exit Sub
TipoErrs:
  MsgBox Err.Description
End Sub

Private Sub CmdNuevo_Click()
Me.TxtCodReferencia.Text = ""
Me.TxtDescripcion.Text = ""
End Sub

Private Sub CmdPirmero_Click()
 Me.DtaReferencia.Recordset.MoveFirst
 If Not Me.DtaReferencia.Recordset.BOF Then
   Me.DBCodigoReferencia.Text = Me.DtaReferencia.Recordset("CodReferencia")
 Else
   MsgBox "Este es el Primer Registro", vbExclamation, "sistema de Nominas"
   Me.DtaReferencia.Recordset.MoveNext
 End If
End Sub

Private Sub CmdPrnDetalle_Click()
  ArepFGDetalle.DataControl1.ConnectionString = ConexionReporte
  ArepFGDetalle.LblTitulo.Caption = Titulo
  ArepFGDetalle.LblSubtitulo.Caption = "REPORTE DETALLE FG POR PROCESOS"
  ArepFGDetalle.ImgLogo.Picture = LoadPicture(RutaLogo)
  
  ArepFGDetalle.DataControl1.Source = "SELECT  Referencia.CodReferencia1, Referencia.CodReferencia, Referencia.Descrip, Referencia.Meta, Procesos.CodProceso, Procesos.Descrip AS DescripcionProceso, " & _
                      "Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, " & _
                      "DetalleProduccion.Linea, DetalleProduccion.Lunes, DetalleProduccion.Martes, DetalleProduccion.Miercoles, DetalleProduccion.Jueves, " & _
                      "DetalleProduccion.Viernes, DetalleProduccion.Sabado, DetalleProduccion.Domingo, DetalleProduccion.SalarioPieza, " & _
                      "DetalleProduccion.NumNomina " & _
"FROM         Referencia INNER JOIN " & _
                      "Procesos ON Referencia.CodReferencia = Procesos.CodReferencia INNER JOIN " & _
                      "DetalleProduccion ON Procesos.CodReferencia = DetalleProduccion.CodReferencia AND " & _
                      "Procesos.CodProceso = DetalleProduccion.CodProceso INNER JOIN " & _
                      "Empleado ON DetalleProduccion.CodEmpleado = Empleado.CodEmpleado " & _
"WHERE     (Referencia.CodReferencia = '" & Me.DBCodigoReferencia.Columns(1) & "') ORDER BY Referencia.CodReferencia, Procesos.CodProceso"


   ArepFGDetalle.Show 1

End Sub

Private Sub CmdPrnNomina_Click()
  ArepFG.DataControl1.ConnectionString = ConexionReporte
  ArepFG.LblTitulo.Caption = Titulo
  ArepFG.LblSubtitulo.Caption = "REPORTE DE FG POR PROCESOS"
  ArepFG.ImgLogo.Picture = LoadPicture(RutaLogo)
                     
   
   ArepFG.DataControl1.Source = "SELECT  DetalleProduccion.CodReferencia AS CodReferencia, DetalleProduccion.CodProceso AS CodProceso, Procesos.Descrip, SUM(DetalleProduccion.Lunes) AS Lunes, SUM(DetalleProduccion.Martes) AS Martes, SUM(DetalleProduccion.Miercoles) AS Miercoles, SUM(DetalleProduccion.Jueves) AS Jueves, SUM(DetalleProduccion.Viernes) AS Viernes, SUM(DetalleProduccion.Sabado) AS Sabado, SUM(DetalleProduccion.Domingo) AS Domingo, SUM(DetalleProduccion.TotalUnidades) AS TotalUnidades, SUM(DetalleProduccion.Lunes + DetalleProduccion.Martes + DetalleProduccion.Miercoles + DetalleProduccion.Jueves + DetalleProduccion.Viernes + DetalleProduccion.Sabado + DetalleProduccion.Domingo) AS Produccion, MAX(DetalleProduccion.SalarioPieza) AS SalarioPieza, Referencia.Meta, Referencia.Meta - SUM(DetalleProduccion.Lunes + DetalleProduccion.Martes + DetalleProduccion.Miercoles + DetalleProduccion.Jueves + DetalleProduccion.Viernes " & _
                                "+ DetalleProduccion.Sabado + DetalleProduccion.Domingo) AS Diferencia, Referencia.Descrip AS DescripcionFG, Referencia.CodReferencia1 FROM DetalleProduccion INNER JOIN Referencia ON DetalleProduccion.CodReferencia = Referencia.CodReferencia INNER JOIN Procesos ON DetalleProduccion.CodReferencia = Procesos.CodReferencia AND DetalleProduccion.CodProceso = Procesos.CodProceso AND Referencia.CodReferencia = Procesos.CodReferencia  GROUP BY DetalleProduccion.CodProceso, DetalleProduccion.CodReferencia, Referencia.Meta, Procesos.Descrip, Referencia.Descrip, Referencia.CodReferencia1 " & _
                                "HAVING  (DetalleProduccion.CodReferencia = '" & Me.DBCodigoReferencia.Columns(1) & "')"
   ArepFG.Show



  
  
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub CmdSiguiente_Click()
 Me.DtaReferencia.Recordset.MoveNext
 If Not Me.DtaReferencia.Recordset.EOF Then
   Me.DBCodigoReferencia.Text = Me.DtaReferencia.Recordset("CodReferencia")
 Else
   MsgBox "Este es el ultimo Registro", vbExclamation, "sistema de Nominas"
   Me.DtaReferencia.Recordset.MovePrevious
 End If
End Sub

Private Sub CmdUltimo_Click()
 Me.DtaReferencia.Recordset.MoveLast
 If Not Me.DtaReferencia.Recordset.EOF Then
   Me.DBCodigoReferencia.Text = Me.DtaReferencia.Recordset("CodReferencia")
 Else
   MsgBox "Este es el ultimo Registro", vbExclamation, "sistema de Nominas"
   Me.DtaReferencia.Recordset.MovePrevious
 End If
End Sub

Private Sub Command1_Click()
 Me.DtaConsulta.RecordSource = "SELECT Activo,Meta, Ref, CodReferencia1, CodReferencia, Descrip From Referencia "
 Me.DtaConsulta.Refresh
 Do While Not Me.DtaConsulta.Recordset.EOF

    Me.DtaConsulta.Recordset("CodReferencia1") = Me.DtaConsulta.Recordset("CodReferencia")

    Me.DtaConsulta.Recordset.Update
    
    Me.DtaConsulta.Recordset.MoveNext
  Loop
  
MsgBox "proceso terminado"


End Sub

Private Sub DBCodigoEmpleado_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub DBCodigoReferencia_ItemChange()
 
     Me.TxtMeta.Text = "0"
    Me.TxtDiferencia.Text = "0"
    Me.TxtProduccion.Text = "0"
    
 Me.DtaConsulta.RecordSource = "SELECT Meta,Ref, CodReferencia1,CodReferencia, Descrip,Activo From Referencia WHERE     (CodReferencia = " & Me.DBCodigoReferencia.Columns(1).Text & ") "
 Me.DtaConsulta.Refresh
 If Not Me.DtaConsulta.Recordset.EOF Then
    Me.TxtReferencia.Text = Me.DtaConsulta.Recordset("Ref")
    Me.TxtDescripcion.Text = Me.DtaConsulta.Recordset("Descrip")
    Me.TxtCodReferencia.Text = Me.DtaConsulta.Recordset("CodReferencia")
    If Not IsNull(Me.DtaConsulta.Recordset("Meta")) Then
     Me.TxtMeta.Text = Me.DtaConsulta.Recordset("Meta")
    End If
    
    If Me.DtaConsulta.Recordset("Activo") = True Then
       Me.ChkActivo.Value = 1
    Else
       Me.ChkActivo.Value = 0
    End If
 
 
 '/////////////////////////////////////////////////////////////////////////////////
 '////////EN ESTA CONSULTA CARGO AL GRID, LOS PROCESOS DEL FG////////////////////////
 '//////////////////////////////////////////////////////////////////////////////////
 
 Me.AdoFGProcesos.RecordSource = "SELECT DetalleProduccion.CodReferencia AS CodReferencia, DetalleProduccion.CodProceso AS CodProceso,  Procesos.Descrip, SUM(DetalleProduccion.Lunes) AS Lunes, " & _
                      "SUM(DetalleProduccion.Martes) AS Martes, SUM(DetalleProduccion.Miercoles) AS Miercoles, SUM(DetalleProduccion.Jueves) AS Jueves, " & _
                      "SUM(DetalleProduccion.Viernes) AS Viernes, SUM(DetalleProduccion.Sabado) AS Sabado, SUM(DetalleProduccion.Domingo) AS Domingo, " & _
                      "SUM(DetalleProduccion.TotalUnidades) AS TotalUnidades, " & _
                      "SUM(DetalleProduccion.Lunes + DetalleProduccion.Martes + DetalleProduccion.Miercoles + DetalleProduccion.Jueves + DetalleProduccion.Viernes + DetalleProduccion.Sabado " & _
                      "+ DetalleProduccion.Domingo) AS Produccion, MAX(DetalleProduccion.SalarioPieza) AS SalarioPieza, Referencia.Meta, " & _
                      "Referencia.Meta - SUM(DetalleProduccion.Lunes + DetalleProduccion.Martes + DetalleProduccion.Miercoles + DetalleProduccion.Jueves + DetalleProduccion.Viernes " & _
                      "+ DetalleProduccion.Sabado + DetalleProduccion.Domingo) AS Diferencia " & _
                      "FROM DetalleProduccion INNER JOIN " & _
                      "Referencia ON DetalleProduccion.CodReferencia = Referencia.CodReferencia INNER JOIN " & _
                      "Procesos ON DetalleProduccion.CodReferencia = Procesos.CodReferencia AND DetalleProduccion.CodProceso = Procesos.CodProceso AND " & _
                      "Referencia.CodReferencia = Procesos.CodReferencia " & _
                      "GROUP BY DetalleProduccion.CodProceso, DetalleProduccion.CodReferencia, Referencia.Meta, Procesos.Descrip " & _
                      "HAVING      (DetalleProduccion.CodReferencia =  " & Me.DBCodigoReferencia.Columns(1) & ") "
 
     Me.AdoFGProcesos.Refresh

 
   
    Me.AdoMetas.RecordSource = "SELECT CodReferencia AS Expr1, CodProceso AS CodProceso, SUM(Lunes) AS Lunes, SUM(Martes) AS Martes, SUM(Miercoles) AS Miercoles, SUM(Jueves) AS Jueves, SUM(Viernes) AS Viernes, SUM(Sabado) AS Sabado, SUM(Domingo) AS Domingo, SUM(TotalUnidades) AS TotalUnidades, MAX(SalarioPieza) As SalarioPieza From DetalleProduccion GROUP BY CodProceso, CodReferencia HAVING  (CodReferencia =" & Me.DBCodigoReferencia.Columns(1).Text & ")"
    Me.AdoMetas.Refresh
    
    
    If Not Me.AdoMetas.Recordset.EOF Then
      Me.TxtProduccion.Text = Format(Me.AdoMetas.Recordset("TotalUnidades"), "##,##0")
      Me.TxtDiferencia.Text = Format(CDbl(Me.TxtMeta.Text) - CDbl(Me.TxtProduccion.Text), "##,##0")
      
      If CDbl(Me.TxtMeta.Text) < CDbl(Me.TxtProduccion.Text) Then
       Me.TxtDiferencia.BackColor = &HFF&
      Else
       Me.TxtDiferencia.BackColor = &HFFFFFF
      End If
      
    End If
 Else
    Me.TxtReferencia.Text = ""
    Me.TxtDescripcion.Text = ""
    Me.TxtMeta.Text = "0"
    Me.TxtDiferencia.Text = "0"
    Me.TxtProduccion.Text = "0"
 End If
End Sub

Private Sub DBCodigoReferencia_KeyUp(KeyCode As Integer, Shift As Integer)

 
   Me.TxtCodReferencia.Text = Me.DBCodigoReferencia.Columns(1)
   Me.TxtCodReferencia2.Text = Me.DBCodigoReferencia.Columns(1)
End Sub

Private Sub Form_Load()


  ' Define a New Style that will be used within this Application
  Set OrderItems = Me.TDBGridProcesos.Styles.Add("ItemSelected")
  OrderItems.BackColor = vbBlue
  OrderItems.ForeColor = vbWhite

  ' Define the Style that will be used for items that are 0
  Dim NoItem As New TrueOleDBGrid80.Style
  NoItem.BackColor = vbRed
  Me.TDBGridProcesos.Columns("Diferencia").AddRegexCellStyle dbgNormalCell, NoItem, "^ *-"     '"^0"
  Me.TDBGridProcesos.Columns("Diferencia").AddRegexCellStyle dbgNormalCell + dbgCurrentCell, NoItem, "^ *-"
  
  
  
 Me.TDBGridProcesos.EvenRowStyle.BackColor = &H80000013
 Me.TDBGridProcesos.OddRowStyle.BackColor = &H80000005
 Me.TDBGridProcesos.AlternatingRowStyle = True

Me.TxtMeta.Text = 0
Me.TxtProduccion.Text = 0
Me.TxtDiferencia.Text = 0

With Me.AdoMetas
  .ConnectionString = Conexion

End With

With Me.DtaConsulta
  .ConnectionString = Conexion

End With

With Me.AdoFGProcesos
  .ConnectionString = Conexion

End With

With Me.DtaReferencia
  .ConnectionString = Conexion
  .RecordSource = "SELECT CodReferencia1, CodReferencia, Descrip, Activo From Referencia ORDER BY CodReferencia1"
  .Refresh
End With

Me.AdoFGProcesos.RecordSource = "SELECT DetalleProduccion.CodReferencia AS CodReferencia, DetalleProduccion.CodProceso AS CodProceso, SUM(DetalleProduccion.Lunes) AS Lunes, " & _
                      "SUM(DetalleProduccion.Martes) AS Martes, SUM(DetalleProduccion.Miercoles) AS Miercoles, SUM(DetalleProduccion.Jueves) AS Jueves, " & _
                      "SUM(DetalleProduccion.Viernes) AS Viernes, SUM(DetalleProduccion.Sabado) AS Sabado, SUM(DetalleProduccion.Domingo) AS Domingo, " & _
                      "SUM(DetalleProduccion.TotalUnidades) AS TotalUnidades, " & _
                      "SUM(DetalleProduccion.Lunes + DetalleProduccion.Martes + DetalleProduccion.Miercoles + DetalleProduccion.Jueves + DetalleProduccion.Viernes + DetalleProduccion.Sabado " & _
                      "+ DetalleProduccion.Domingo) AS Produccion, MAX(DetalleProduccion.SalarioPieza) AS SalarioPieza, Referencia.Meta, " & _
                      "Referencia.Meta - SUM(DetalleProduccion.Lunes + DetalleProduccion.Martes + DetalleProduccion.Miercoles + DetalleProduccion.Jueves + DetalleProduccion.Viernes " & _
                      "+ DetalleProduccion.Sabado + DetalleProduccion.Domingo) AS Diferencia, Procesos.Descrip " & _
                      "FROM DetalleProduccion INNER JOIN " & _
                      "Referencia ON DetalleProduccion.CodReferencia = Referencia.CodReferencia INNER JOIN " & _
                      "Procesos ON DetalleProduccion.CodReferencia = Procesos.CodReferencia AND DetalleProduccion.CodProceso = Procesos.CodProceso AND " & _
                      "Referencia.CodReferencia = Procesos.CodReferencia " & _
                      "GROUP BY DetalleProduccion.CodProceso, DetalleProduccion.CodReferencia, Referencia.Meta, Procesos.Descrip " & _
                      "HAVING      (DetalleProduccion.CodReferencia = '-0071') "
Me.AdoFGProcesos.Refresh


End Sub

Private Sub TxtCodReferencia2_Change()
    Me.TxtMeta.Text = "0"
    Me.TxtDiferencia.Text = "0"
    Me.TxtProduccion.Text = "0"
    
 Me.DtaConsulta.RecordSource = "SELECT Meta,Ref, CodReferencia1,CodReferencia, Descrip,Activo From Referencia WHERE     (CodReferencia = " & Me.DBCodigoReferencia.Columns(1).Text & ") "
 Me.DtaConsulta.Refresh
 If Not Me.DtaConsulta.Recordset.EOF Then
    Me.TxtReferencia.Text = Me.DtaConsulta.Recordset("Ref")
    Me.TxtDescripcion.Text = Me.DtaConsulta.Recordset("Descrip")
    Me.TxtCodReferencia.Text = Me.DtaConsulta.Recordset("CodReferencia")
    If Not IsNull(Me.DtaConsulta.Recordset("Meta")) Then
     Me.TxtMeta.Text = Me.DtaConsulta.Recordset("Meta")
    End If
    
    If Me.DtaConsulta.Recordset("Activo") = True Then
       Me.ChkActivo.Value = 1
    Else
       Me.ChkActivo.Value = 0
    End If
 
 
 '/////////////////////////////////////////////////////////////////////////////////
 '////////EN ESTA CONSULTA CARGO AL GRID, LOS PROCESOS DEL FG////////////////////////
 '//////////////////////////////////////////////////////////////////////////////////
 
 Me.AdoFGProcesos.RecordSource = "SELECT DetalleProduccion.CodReferencia AS CodReferencia, DetalleProduccion.CodProceso AS CodProceso,  Procesos.Descrip, SUM(DetalleProduccion.Lunes) AS Lunes, " & _
                      "SUM(DetalleProduccion.Martes) AS Martes, SUM(DetalleProduccion.Miercoles) AS Miercoles, SUM(DetalleProduccion.Jueves) AS Jueves, " & _
                      "SUM(DetalleProduccion.Viernes) AS Viernes, SUM(DetalleProduccion.Sabado) AS Sabado, SUM(DetalleProduccion.Domingo) AS Domingo, " & _
                      "SUM(DetalleProduccion.TotalUnidades) AS TotalUnidades, " & _
                      "SUM(DetalleProduccion.Lunes + DetalleProduccion.Martes + DetalleProduccion.Miercoles + DetalleProduccion.Jueves + DetalleProduccion.Viernes + DetalleProduccion.Sabado " & _
                      "+ DetalleProduccion.Domingo) AS Produccion, MAX(DetalleProduccion.SalarioPieza) AS SalarioPieza, Referencia.Meta, " & _
                      "Referencia.Meta - SUM(DetalleProduccion.Lunes + DetalleProduccion.Martes + DetalleProduccion.Miercoles + DetalleProduccion.Jueves + DetalleProduccion.Viernes " & _
                      "+ DetalleProduccion.Sabado + DetalleProduccion.Domingo) AS Diferencia " & _
                      "FROM DetalleProduccion INNER JOIN " & _
                      "Referencia ON DetalleProduccion.CodReferencia = Referencia.CodReferencia INNER JOIN " & _
                      "Procesos ON DetalleProduccion.CodReferencia = Procesos.CodReferencia AND DetalleProduccion.CodProceso = Procesos.CodProceso AND " & _
                      "Referencia.CodReferencia = Procesos.CodReferencia " & _
                      "GROUP BY DetalleProduccion.CodProceso, DetalleProduccion.CodReferencia, Referencia.Meta, Procesos.Descrip " & _
                      "HAVING      (DetalleProduccion.CodReferencia =  " & Me.DBCodigoReferencia.Columns(1) & ") "
 
     Me.AdoFGProcesos.Refresh

 
   
    Me.AdoMetas.RecordSource = "SELECT CodReferencia AS Expr1, CodProceso AS CodProceso, SUM(Lunes) AS Lunes, SUM(Martes) AS Martes, SUM(Miercoles) AS Miercoles, SUM(Jueves) AS Jueves, SUM(Viernes) AS Viernes, SUM(Sabado) AS Sabado, SUM(Domingo) AS Domingo, SUM(TotalUnidades) AS TotalUnidades, MAX(SalarioPieza) As SalarioPieza From DetalleProduccion GROUP BY CodProceso, CodReferencia HAVING  (CodReferencia =" & Me.DBCodigoReferencia.Columns(1).Text & ")"
    Me.AdoMetas.Refresh
    
    
    If Not Me.AdoMetas.Recordset.EOF Then
      Me.TxtProduccion.Text = Format(Me.AdoMetas.Recordset("TotalUnidades"), "##,##0")
      Me.TxtDiferencia.Text = Format(CDbl(Me.TxtMeta.Text) - CDbl(Me.TxtProduccion.Text), "##,##0")
      
      If CDbl(Me.TxtMeta.Text) < CDbl(Me.TxtProduccion.Text) Then
       Me.TxtDiferencia.BackColor = &HFF&
      Else
       Me.TxtDiferencia.BackColor = &HFFFFFF
      End If
      
    End If
 Else
    Me.TxtReferencia.Text = ""
    Me.TxtDescripcion.Text = ""
    Me.TxtMeta.Text = "0"
    Me.TxtDiferencia.Text = "0"
    Me.TxtProduccion.Text = "0"
 End If
End Sub

Private Sub TxtMeta_Change()
         Me.TxtDiferencia.BackColor = &HFFFFFF
   
      If Me.TxtMeta.Text = "" Then
       Me.TxtDiferencia.Text = Format(CDbl(0) - CDbl(Me.TxtProduccion.Text), "##,##0")
       Me.TxtDiferencia.BackColor = &HFF&
      ElseIf Not Me.TxtProduccion.Text = "" Then
       Me.TxtDiferencia.Text = Format(CDbl(Me.TxtMeta.Text) - CDbl(Me.TxtProduccion.Text), "##,##0")
      
         If CDbl(Me.TxtMeta.Text) < CDbl(Me.TxtProduccion.Text) Then
           Me.TxtDiferencia.BackColor = &HFF&
         Else
           Me.TxtDiferencia.BackColor = &HFFFFFF
         End If
         
      End If
      

End Sub

Private Sub TxtMeta_LostFocus()
   Me.TxtMeta.Text = Format(Me.TxtMeta.Text, "##,##0")
End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub
