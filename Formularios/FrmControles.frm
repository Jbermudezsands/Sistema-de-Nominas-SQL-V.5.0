VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmControles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controles Personalizados"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AdoEmpleados 
      Height          =   375
      Left            =   3840
      Top             =   8400
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
   Begin MSAdodcLib.Adodc AdoDetalleNominas 
      Height          =   375
      Left            =   840
      Top             =   9240
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
      Caption         =   "AdoDetalleNominas"
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
      Caption         =   "Aceptar"
      DownPicture     =   "FrmControles.frx":0000
      Height          =   735
      Left            =   6360
      MouseIcon       =   "FrmControles.frx":1AE2
      MousePointer    =   99  'Custom
      Picture         =   "FrmControles.frx":1F24
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      DownPicture     =   "FrmControles.frx":2856
      Height          =   735
      Left            =   7440
      MouseIcon       =   "FrmControles.frx":4338
      MousePointer    =   99  'Custom
      Picture         =   "FrmControles.frx":477A
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6600
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   -120
      ScaleHeight     =   1095
      ScaleWidth      =   8415
      TabIndex        =   41
      Top             =   0
      Width           =   8415
      Begin VB.Image Image2 
         Height          =   1020
         Left            =   0
         Picture         =   "FrmControles.frx":53BC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1290
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   8400
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "CONTROLES PERSONALIZADOS"
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
         Left            =   1920
         TabIndex        =   42
         Top             =   360
         Width           =   3840
      End
   End
   Begin MSAdodcLib.Adodc AdoDatosEmpresa 
      Height          =   375
      Left            =   720
      Top             =   8880
      Width           =   3375
      _ExtentX        =   5953
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
   Begin MSAdodcLib.Adodc DtaControles 
      Height          =   375
      Left            =   840
      Top             =   8400
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
      Caption         =   "DtaControles"
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
   Begin vbskfree.Skinner Skinner1 
      Left            =   0
      Top             =   7080
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales de la Empesa"
      TabPicture(0)   =   "FrmControles.frx":5D99
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Configuracion del Sistema"
      TabPicture(1)   =   "FrmControles.frx":5DB5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "ChkTasa"
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(5)=   "Frame8"
      Tab(1).Control(6)=   "Frame10"
      Tab(1).Control(7)=   "Frame14"
      Tab(1).Control(8)=   "ckRedondeo"
      Tab(1).Control(9)=   "TxtValorHorasExtra"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Ajustes"
      TabPicture(2)   =   "FrmControles.frx":5DD1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame16"
      Tab(2).Control(1)=   "Frame15"
      Tab(2).Control(2)=   "Frame7"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Conexiones"
      TabPicture(3)   =   "FrmControles.frx":5DED
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame9"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Salario Mínimo y Valor de Puntos"
      TabPicture(4)   =   "FrmControles.frx":5E09
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame11"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Procesar"
      TabPicture(5)   =   "FrmControles.frx":5E25
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame17"
      Tab(5).ControlCount=   1
      Begin VB.TextBox TxtValorHorasExtra 
         Height          =   285
         Left            =   -71760
         TabIndex        =   115
         Top             =   2040
         Width           =   735
      End
      Begin VB.Frame Frame17 
         Caption         =   "Agregar User "
         Height          =   975
         Left            =   -74760
         TabIndex        =   111
         Top             =   1080
         Width           =   7215
         Begin VB.CommandButton Command6 
            Caption         =   "Procesar"
            Height          =   615
            Left            =   5880
            Picture         =   "FrmControles.frx":5E41
            Style           =   1  'Graphical
            TabIndex        =   112
            Top             =   240
            Width           =   1095
         End
         Begin XtremeSuiteControls.ProgressBar ProgressBar2 
            Height          =   375
            Left            =   120
            TabIndex        =   113
            Top             =   360
            Width           =   5055
            _Version        =   786432
            _ExtentX        =   8916
            _ExtentY        =   661
            _StockProps     =   93
            BackColor       =   14737632
            Scrolling       =   1
            Appearance      =   6
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Asignacion Horarios e Historicos"
         Height          =   975
         Left            =   -74760
         TabIndex        =   108
         Top             =   4080
         Width           =   7215
         Begin VB.CommandButton Command7 
            Caption         =   "Historicos"
            Height          =   615
            Left            =   6000
            Picture         =   "FrmControles.frx":63CB
            Style           =   1  'Graphical
            TabIndex        =   116
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Procesar"
            Height          =   615
            Left            =   4800
            Picture         =   "FrmControles.frx":6955
            Style           =   1  'Graphical
            TabIndex        =   110
            Top             =   240
            Width           =   1095
         End
         Begin XtremeSuiteControls.ProgressBar ProgressBar1 
            Height          =   375
            Left            =   120
            TabIndex        =   109
            Top             =   360
            Width           =   4575
            _Version        =   786432
            _ExtentX        =   8070
            _ExtentY        =   661
            _StockProps     =   93
            BackColor       =   14737632
            Scrolling       =   1
            Appearance      =   6
         End
      End
      Begin VB.CheckBox ckRedondeo 
         Caption         =   "Calcular Redondeado"
         Height          =   255
         Left            =   -70200
         TabIndex        =   107
         Top             =   960
         Width           =   2055
      End
      Begin VB.Frame Frame15 
         Caption         =   "Ajustar de Salarios"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   96
         Top             =   2520
         Width           =   7215
         Begin VB.TextBox TxtPorciento 
            Height          =   315
            Left            =   4920
            TabIndex        =   106
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Procesar"
            Height          =   735
            Left            =   5880
            Picture         =   "FrmControles.frx":6EDF
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox TxtHasta 
            Height          =   315
            Left            =   2520
            TabIndex        =   100
            Top             =   310
            Width           =   1095
         End
         Begin VB.TextBox TxtDesde 
            Height          =   315
            Left            =   840
            TabIndex        =   99
            Top             =   310
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Ajustar Salario Basico"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   1200
            Value           =   -1  'True
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmControles.frx":7469
            TabIndex        =   98
            Top             =   1200
            Width           =   4935
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmControles.frx":74C7
            TabIndex        =   102
            Top             =   360
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
            Height          =   255
            Left            =   2040
            OleObjectBlob   =   "FrmControles.frx":752F
            TabIndex        =   103
            Top             =   360
            Width           =   855
         End
         Begin XtremeSuiteControls.ProgressBar ProgressBar 
            Height          =   375
            Left            =   240
            TabIndex        =   104
            Top             =   720
            Width           =   5055
            _Version        =   786432
            _ExtentX        =   8916
            _ExtentY        =   661
            _StockProps     =   93
            BackColor       =   14737632
            Scrolling       =   1
            Appearance      =   6
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
            Height          =   255
            Left            =   3840
            OleObjectBlob   =   "FrmControles.frx":7597
            TabIndex        =   105
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Configurar Liquidacion"
         Height          =   1455
         Left            =   -70320
         TabIndex        =   92
         Top             =   3720
         Width           =   2895
         Begin VB.CheckBox ChkAntiguedadMenor 
            Caption         =   "Pag Antiguedad Menor 1año"
            Height          =   255
            Left            =   240
            TabIndex        =   95
            Top             =   1080
            Width           =   2415
         End
         Begin VB.OptionButton OptSalarioPromedioPeriodo 
            Caption         =   "Salario Promedio Entre Periodo"
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   720
            Width           =   2535
         End
         Begin VB.OptionButton OptSalarioPromedioReal 
            Caption         =   "Salario Promedio Dif Dias"
            Height          =   255
            Left            =   240
            TabIndex        =   93
            Top             =   360
            Value           =   -1  'True
            Width           =   2535
         End
      End
      Begin VB.Frame Frame11 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   75
         Top             =   1080
         Width           =   7455
         Begin VB.Frame Frame13 
            Caption         =   "Valor de Puntos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   3720
            TabIndex        =   84
            Top             =   960
            Width           =   3495
            Begin VB.TextBox txtPtsAnt 
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   86
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txtPtsAct 
               Height          =   285
               Left            =   1800
               TabIndex        =   85
               Top             =   600
               Width           =   1575
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmControles.frx":760D
               TabIndex        =   87
               Top             =   360
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
               Height          =   255
               Left            =   1800
               OleObjectBlob   =   "FrmControles.frx":767B
               TabIndex        =   88
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Salario Mínimo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   79
            Top             =   960
            Width           =   3495
            Begin VB.TextBox txtSalAnt 
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   81
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txtSalAct 
               Height          =   285
               Left            =   1800
               TabIndex        =   80
               Top             =   600
               Width           =   1575
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmControles.frx":76E5
               TabIndex        =   82
               Top             =   360
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
               Height          =   255
               Left            =   1800
               OleObjectBlob   =   "FrmControles.frx":7753
               TabIndex        =   83
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.CommandButton cmdSalPts 
            Caption         =   "Procesar"
            Height          =   735
            Left            =   120
            Picture         =   "FrmControles.frx":77BD
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   2640
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmControles.frx":7D47
            TabIndex        =   77
            Top             =   240
            Width           =   1215
         End
         Begin TrueOleDBList80.TDBCombo tdbcNomina 
            Height          =   315
            Left            =   120
            TabIndex        =   78
            Top             =   480
            Width           =   7095
            _ExtentX        =   12515
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
            _PropDict       =   $"FrmControles.frx":7DB3
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
         Begin XtremeSuiteControls.ProgressBar ospSalPts 
            Height          =   375
            Left            =   120
            TabIndex        =   89
            Top             =   2040
            Width           =   7095
            _Version        =   786432
            _ExtentX        =   12515
            _ExtentY        =   661
            _StockProps     =   93
            BackColor       =   14737632
            Scrolling       =   1
            Appearance      =   6
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Configuracion del IR"
         Height          =   1095
         Left            =   -70320
         TabIndex        =   67
         Top             =   2640
         Width           =   2295
         Begin VB.OptionButton Option1 
            Caption         =   "Calcular Ajustando IR"
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Calcular IR x 12"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "CONEXION SISTEMA CONTABLE"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   63
         Top             =   1020
         Width           =   6855
         Begin VB.CommandButton Command3 
            Height          =   375
            Left            =   5880
            Picture         =   "FrmControles.frx":7E5D
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox TxtConexionString 
            Height          =   1515
            Left            =   1920
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   64
            Top             =   360
            Width           =   3855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "FrmControles.frx":8313
            TabIndex        =   65
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Configuracion Vacaciones"
         Height          =   1095
         Left            =   -70320
         TabIndex        =   60
         Top             =   1560
         Width           =   2415
         Begin VB.OptionButton OptVacacionesMensuales 
            Caption         =   "Vacaciones Mensuales"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   720
            Width           =   2175
         End
         Begin VB.OptionButton OptVacacionesSemestrales 
            Caption         =   "Vacaciones Semestrales"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Ajustar de Salarios"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   53
         Top             =   960
         Width           =   6615
         Begin VB.ComboBox CmbSimbolo 
            Height          =   315
            ItemData        =   "FrmControles.frx":838F
            Left            =   960
            List            =   "FrmControles.frx":83A2
            TabIndex        =   73
            Text            =   "="
            Top             =   310
            Width           =   735
         End
         Begin VB.OptionButton OptTarifa 
            Caption         =   "Ajustar Tarifa Horaria"
            Height          =   255
            Left            =   2520
            TabIndex        =   72
            Top             =   1200
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton OptSalario 
            Caption         =   "Ajustar Salario Basico"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   1200
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblProgreso 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmControles.frx":83B7
            TabIndex        =   59
            Top             =   1200
            Width           =   4935
         End
         Begin VB.TextBox TxtSalarioAjuste 
            Height          =   315
            Left            =   3960
            TabIndex        =   58
            Top             =   310
            Width           =   1095
         End
         Begin VB.TextBox TxtSalario 
            Height          =   315
            Left            =   1800
            TabIndex        =   56
            Top             =   310
            Width           =   1095
         End
         Begin VB.CommandButton CmdProcesarSalario 
            Caption         =   "Procesar"
            Height          =   735
            Left            =   5280
            Picture         =   "FrmControles.frx":8415
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   240
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblNombre1 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmControles.frx":899F
            TabIndex        =   54
            Top             =   360
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblNombre2 
            Height          =   255
            Left            =   3000
            OleObjectBlob   =   "FrmControles.frx":8A0D
            TabIndex        =   57
            Top             =   360
            Width           =   855
         End
         Begin XtremeSuiteControls.ProgressBar BarraEmpleados 
            Height          =   375
            Left            =   120
            TabIndex        =   91
            Top             =   720
            Width           =   5055
            _Version        =   786432
            _ExtentX        =   8916
            _ExtentY        =   661
            _StockProps     =   93
            BackColor       =   14737632
            Scrolling       =   1
            Appearance      =   6
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Configuracion de Reportes"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   36
         Top             =   2340
         Width           =   4455
         Begin VB.Frame Frame6 
            Height          =   1455
            Left            =   120
            TabIndex        =   48
            Top             =   1200
            Width           =   4095
            Begin VB.CheckBox Chk7mo 
               Caption         =   "No Calc 7mo conforme Produccion"
               Height          =   375
               Left            =   1080
               TabIndex        =   74
               Top             =   960
               Width           =   2895
            End
            Begin MSComCtl2.DTPicker DTFecha 
               Height          =   285
               Left            =   2160
               TabIndex        =   51
               Top             =   240
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   503
               _Version        =   393216
               Format          =   77266945
               CurrentDate     =   39737
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Procesar"
               Height          =   735
               Left            =   120
               Picture         =   "FrmControles.frx":8A81
               Style           =   1  'Graphical
               TabIndex        =   49
               Top             =   120
               Width           =   1215
            End
            Begin XtremeSuiteControls.ProgressBar Barra 
               Height          =   255
               Left            =   1560
               TabIndex        =   90
               Top             =   600
               Visible         =   0   'False
               Width           =   2415
               _Version        =   786432
               _ExtentX        =   4260
               _ExtentY        =   450
               _StockProps     =   93
               BackColor       =   14737632
               Scrolling       =   1
               Appearance      =   6
            End
            Begin VB.Label Label3 
               Caption         =   "Hasta"
               Height          =   255
               Left            =   1560
               TabIndex        =   50
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.ComboBox CmbNominas 
            Height          =   315
            ItemData        =   "FrmControles.frx":900B
            Left            =   1560
            List            =   "FrmControles.frx":9024
            TabIndex        =   40
            Text            =   "Predeterminado"
            Top             =   840
            Width           =   2775
         End
         Begin VB.ComboBox CmbColillas 
            Height          =   315
            ItemData        =   "FrmControles.frx":90BD
            Left            =   1560
            List            =   "FrmControles.frx":90D9
            TabIndex        =   38
            Text            =   "Predeterminado"
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label LblProcesos 
            Height          =   255
            Left            =   1680
            TabIndex        =   52
            Top             =   360
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Formatos Nominas"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Formatos Colillas"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.CheckBox ChkTasa 
         Caption         =   "Verificar Tasa al Entrar"
         Height          =   375
         Left            =   -70200
         TabIndex        =   35
         ToolTipText     =   "Verifica si la tasa del día ya ha sido Grabada"
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dias a la Semana"
         Height          =   975
         Left            =   -72960
         TabIndex        =   32
         Top             =   960
         Width           =   1695
         Begin VB.OptionButton OPT6 
            Caption         =   "Seis Dias"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Opt7 
            Caption         =   "Siete Dias"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dias al Mes"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   29
         Top             =   960
         Width           =   1575
         Begin VB.OptionButton Opt25 
            Caption         =   "25 Días"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton OptExacto 
            Caption         =   "(365/12) Días"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Opt30 
            Caption         =   "30 Días"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3375
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   6855
         Begin VB.TextBox TxtRutaFoto 
            Height          =   375
            Left            =   2880
            TabIndex        =   44
            Top             =   2760
            Width           =   3495
         End
         Begin VB.CommandButton Command1 
            Height          =   375
            Left            =   6360
            Picture         =   "FrmControles.frx":918B
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   2760
            Width           =   375
         End
         Begin VB.TextBox TxtNombreEmpresa 
            Height          =   285
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   21
            Top             =   120
            Width           =   2655
         End
         Begin VB.CommandButton CmdBuscarLogo 
            Height          =   375
            Left            =   6360
            Picture         =   "FrmControles.frx":9641
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   2280
            Width           =   375
         End
         Begin VB.PictureBox ImgLogo2 
            AutoSize        =   -1  'True
            Height          =   2055
            Left            =   120
            ScaleHeight     =   1995
            ScaleWidth      =   2235
            TabIndex        =   19
            Top             =   240
            Width           =   2295
            Begin VB.Image ImgLogo 
               Height          =   2055
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   2295
            End
         End
         Begin VB.TextBox TxtRutaLogo 
            Height          =   375
            Left            =   2880
            TabIndex        =   18
            Top             =   2280
            Width           =   3495
         End
         Begin VB.TextBox TxtFax 
            Height          =   285
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   17
            Top             =   1560
            Width           =   2655
         End
         Begin VB.TextBox TxtEmail 
            Height          =   285
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   16
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox TxtDireccionEmpresa 
            Height          =   765
            Left            =   4080
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox TxtRucEmpresa 
            Height          =   285
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   14
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox TxtTelefono 
            Height          =   285
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   13
            Top             =   1320
            Width           =   2655
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmControles.frx":9AF7
            TabIndex        =   22
            Top             =   2400
            Width           =   2535
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
            Height          =   255
            Left            =   2640
            OleObjectBlob   =   "FrmControles.frx":9B93
            TabIndex        =   23
            Top             =   120
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   3480
            OleObjectBlob   =   "FrmControles.frx":9C0D
            TabIndex        =   24
            Top             =   1560
            Width           =   375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
            Height          =   255
            Left            =   3360
            OleObjectBlob   =   "FrmControles.frx":9C71
            TabIndex        =   25
            Top             =   1800
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
            Height          =   255
            Left            =   3120
            OleObjectBlob   =   "FrmControles.frx":9CDD
            TabIndex        =   26
            Top             =   360
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
            Height          =   255
            Left            =   2880
            OleObjectBlob   =   "FrmControles.frx":9D4D
            TabIndex        =   27
            Top             =   1080
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
            Height          =   255
            Left            =   3120
            OleObjectBlob   =   "FrmControles.frx":9DBF
            TabIndex        =   28
            Top             =   1320
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmControles.frx":9E2F
            TabIndex        =   45
            Top             =   2880
            Width           =   2775
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   6855
         Begin VB.TextBox TxtIO 
            Height          =   405
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   6
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox TxtCO 
            Height          =   405
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   5
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox TxtOI 
            Height          =   405
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox TxtC 
            Height          =   285
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   3
            Top             =   2280
            Width           =   3135
         End
         Begin VB.TextBox TxtI 
            Height          =   285
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   2
            Top             =   1800
            Width           =   3135
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmControles.frx":9ED3
            TabIndex        =   7
            Top             =   240
            Width           =   3375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmControles.frx":9F7F
            TabIndex        =   8
            Top             =   720
            Width           =   3495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmControles.frx":A027
            TabIndex        =   9
            Top             =   1200
            Width           =   3495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmControles.frx":A0DD
            TabIndex        =   10
            Top             =   1800
            Width           =   3255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmControles.frx":A173
            TabIndex        =   11
            Top             =   2280
            Width           =   3495
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Valor Horas Extra"
         Height          =   255
         Left            =   -73200
         TabIndex        =   114
         Top             =   2040
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog CMRutaFoto 
      Left            =   120
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   256
   End
   Begin MSAdodcLib.Adodc DtaHorarioEmpleado 
      Height          =   375
      Left            =   3840
      Top             =   7920
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
      Left            =   720
      Top             =   7920
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
      CommandType     =   2
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
   Begin MSAdodcLib.Adodc AdoUser 
      Height          =   375
      Left            =   3960
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
      Caption         =   "AdoUser"
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
Attribute VB_Name = "FrmControles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnx As New ADODB.Connection
Private rs As New ADODB.Recordset
Private sql As String

Private Sub CmbNominas_Change()
  If Me.CmbNominas.Text = "Nomina Bono Produccion" Then
    Me.Frame6.Visible = True
  Else
    Me.Frame6.Visible = False
  End If
End Sub

Private Sub CmbNominas_Click()
  If Me.CmbNominas.Text = "Nomina Bono Produccion" Then
    Me.Frame6.Visible = True
  Else
    Me.Frame6.Visible = False
  End If
End Sub

Private Sub CmdAceptar_Click()
'On Error GoTo TipoErr
DtaControles.Refresh
'DtaControles.Recordset.Edit
If OptExacto Then
   DtaControles.Recordset("DiasMes") = (365 / 12)
ElseIf Me.Opt30.Value = True Then
   DtaControles.Recordset("DiasMes") = 30
ElseIf Me.Opt25.Value = True Then
   DtaControles.Recordset("DiasMes") = 25
End If

If Opt7 Then
   DtaControles.Recordset("DiasSemana") = 7
Else
   DtaControles.Recordset("DiasSemana") = 6
End If

If ChkTasa.Value = 1 Then
  DtaControles.Recordset("verificartasa") = True
Else
DtaControles.Recordset("verificartasa") = False
End If

If Me.ckRedondeo.Value = 1 Then
  DtaControles.Recordset("CalcularRedondeado") = True
Else
DtaControles.Recordset("CalcularRedondeado") = False
End If

If Me.OptSalarioPromedioReal.Value = True Then
  DtaControles.Recordset("SalarioPromedioReal") = 1
Else
  DtaControles.Recordset("SalarioPromedioReal") = 0
End If

If Me.ChkAntiguedadMenor.Value = 1 Then
  DtaControles.Recordset("AntiguedadMenor") = True
Else
  DtaControles.Recordset("AntiguedadMenor") = False
End If

DtaControles.Recordset.Update


Me.AdoDatosEmpresa.Refresh
If Not Me.AdoDatosEmpresa.Recordset.EOF Then
 Me.AdoDatosEmpresa.Recordset("NombreEmpresa") = Me.TxtNombreEmpresa.Text
 Me.AdoDatosEmpresa.Recordset("RutaFoto") = Me.TxtRutaFoto.Text
 Me.AdoDatosEmpresa.Recordset("RutaLogo") = Me.TxtRutaLogo.Text
 Me.AdoDatosEmpresa.Recordset("NumeroRUC") = Me.TxtRucEmpresa.Text
 Me.AdoDatosEmpresa.Recordset("Direccion") = Me.TxtDireccionEmpresa.Text
 Me.AdoDatosEmpresa.Recordset("Telefono") = Me.TxtTelefono.Text
 Me.AdoDatosEmpresa.Recordset("Fax") = Me.TxtFax.Text
 Me.AdoDatosEmpresa.Recordset("Email") = Me.TxtEmail.Text
 Me.AdoDatosEmpresa.Recordset("FormatoColilla") = Me.CmbColillas.Text
 Me.AdoDatosEmpresa.Recordset("FormatoNomina") = Me.CmbNominas.Text
 
 If Me.TxtValorHorasExtra.Text = "" Then
   Me.TxtValorHorasExtra.Text = 2
 End If
 
 Me.AdoDatosEmpresa.Recordset("HorasExtra") = Me.TxtValorHorasExtra.Text
 
 
 If Me.Chk7mo.Value = 0 Then
  Me.AdoDatosEmpresa.Recordset("Calcular7mo") = 1
 Else
  Me.AdoDatosEmpresa.Recordset("Calcular7mo") = 0
 End If
 
 If Me.OptVacacionesSemestrales.Value = True Then
  Me.AdoDatosEmpresa.Recordset("MetodoVacaciones") = "Vacaciones Semestrales"
 ElseIf Me.OptVacacionesMensuales.Value = True Then
  Me.AdoDatosEmpresa.Recordset("MetodoVacaciones") = "Vacaciones Mensuales"
 End If
 
 If Me.Option1.Value = True Then
  Me.AdoDatosEmpresa.Recordset("TipoCalculoIR") = "Calcular Ajustando IR"
 Else
  Me.AdoDatosEmpresa.Recordset("TipoCalculoIR") = "Calcular IR x 12"
 End If
 

 
 If Me.TxtConexionString.Text <> "" Then
   Me.AdoDatosEmpresa.Recordset("ConexionSistemaContable") = Me.TxtConexionString.Text
 End If
 
 Me.AdoDatosEmpresa.Recordset.Update
Else
 Me.AdoDatosEmpresa.Recordset.AddNew
 Me.AdoDatosEmpresa.Recordset("Numero") = 1
 Me.AdoDatosEmpresa.Recordset("NombreEmpresa") = Me.TxtNombreEmpresa.Text
 Me.AdoDatosEmpresa.Recordset("RutaFoto") = Me.TxtRutaFoto.Text
 Me.AdoDatosEmpresa.Recordset("RutaLogo") = Me.TxtRutaLogo.Text
 Me.AdoDatosEmpresa.Recordset("NumeroRUC") = Me.TxtRucEmpresa.Text
 Me.AdoDatosEmpresa.Recordset("Direccion") = Me.TxtDireccionEmpresa.Text
 Me.AdoDatosEmpresa.Recordset("Telefono") = Me.TxtTelefono.Text
 Me.AdoDatosEmpresa.Recordset("Fax") = Me.TxtFax.Text
 Me.AdoDatosEmpresa.Recordset("Email") = Me.TxtEmail.Text
 Me.AdoDatosEmpresa.Recordset("FormatoColilla") = Me.CmbColillas.Text
 Me.AdoDatosEmpresa.Recordset("FormatoNomina") = Me.CmbNominas.Text
 If Me.OptVacacionesSemestrales.Value = True Then
  Me.AdoDatosEmpresa.Recordset("MetodoVacaciones") = "Vacaciones Semestrales"
 ElseIf Me.OptVacacionesMensuales.Value = True Then
  Me.AdoDatosEmpresa.Recordset("MetodoVacaciones") = "Vacaciones Mensuales"
 End If
 
 If Me.TxtValorHorasExtra.Text = "" Then
   Me.TxtValorHorasExtra.Text = 2
 End If
 
 Me.AdoDatosEmpresa.Recordset("HorasExtra") = Me.TxtValorHorasExtra.Text
 
  If Me.TxtConexionString.Text <> "" Then
   Me.AdoDatosEmpresa.Recordset("ConexionSistemaContable") = Me.TxtConexionString.Text
  End If
 
 Me.AdoDatosEmpresa.Recordset.Update
End If



RutaLogo = Me.TxtRutaLogo.Text
RutaFoto = Me.TxtRutaFoto.Text

Unload Me

Exit Sub
TipoErr:
    ControlErrores
End Sub

Private Sub CmdBuscarLogo_Click()
Dim retval
Dim OPENFILENAME As String
    On Error Resume Next
    ' Set the commom dialog properties we need
    If Me.TxtRutaLogo.Text <> "" Then
       CMRutaFoto.InitDir = Me.TxtRutaLogo.Text
    End If
    CMRutaFoto.FileName = ""
    ' We will load BMP, JPG, and TIF files
    
    CMRutaFoto.Filter = "Image Files |*.bmp;*.gif;*.jpg;*.png;*.tif|All files |*.*"
    ' Display common dialog box
    CMRutaFoto.ShowOpen
    Me.TxtRutaLogo.Text = CMRutaFoto.FileName
    
    Me.ImgLogo.Picture = LoadPicture(Me.TxtRutaLogo.Text)
End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub CmdProcesarSalario_Click()
On Error GoTo TipoErrs
Dim Simbolo As String




Dim SalarioBasico As Double, SalarioAjuste As Double, Maximo As Double, i As Double

  Me.txtSalario.Enabled = False
  Me.TxtSalarioAjuste.Enabled = False
  
  If Me.CmbSimbolo.Text = "" Then
   MsgBox "Debe seleccionar la operacion a realizar", "Zeus Nominas"
   Exit Sub
  End If
  
  Simbolo = Me.CmbSimbolo.Text
  
  If Me.txtSalario.Text = "" Then
    MsgBox "Se necesita un Valor para el Salario", vbCritical, "Sistema de Nominas"
     Me.txtSalario.Enabled = True
     Me.TxtSalarioAjuste.Enabled = True
    Exit Sub
  ElseIf Not IsNumeric(Me.txtSalario.Text) Then
    MsgBox "Se necesita un Valor Numerico para Salario", vbCritical, "Sistema de Nominas"
         Me.txtSalario.Enabled = True
     Me.TxtSalarioAjuste.Enabled = True
    Exit Sub
  End If
  
  If Me.TxtSalarioAjuste.Text = "" Then
    MsgBox "Se necesita un Valor para el Salario de Ajuste", vbCritical, "Sistema de Nominas"
         Me.txtSalario.Enabled = True
     Me.TxtSalarioAjuste.Enabled = True
    Exit Sub
  ElseIf Not IsNumeric(Me.TxtSalarioAjuste.Text) Then
    MsgBox "Se necesita un Valor Numerico para Salario Ajuste", vbCritical, "Sistema de Nominas"
         Me.txtSalario.Enabled = True
     Me.TxtSalarioAjuste.Enabled = True
    Exit Sub
  End If
  
  SalarioBasico = Me.txtSalario.Text
  SalarioAjuste = Me.TxtSalarioAjuste.Text
  
  If Me.OptTarifa.Value = True Then
    Me.AdoEmpleados.RecordSource = "SELECT CodEmpleado, SueldoPeriodo, TarifaHoraria, PorcentajeComision, SalarioMinimo From Empleado Where (TarifaHoraria " & Simbolo & "  " & SalarioBasico & ") AND (Activo = 1)"
    Me.AdoEmpleados.Refresh
  Else
    Me.AdoEmpleados.RecordSource = "SELECT CodEmpleado, SueldoPeriodo, TarifaHoraria, PorcentajeComision, SalarioMinimo From Empleado Where (SueldoPeriodo " & Simbolo & "  " & SalarioBasico & ") AND (Activo = 1)"
    Me.AdoEmpleados.Refresh
  End If
  
   Maximo = Me.AdoEmpleados.Recordset.RecordCount
   Me.BarraEmpleados.Min = 0
   Me.BarraEmpleados.Max = Maximo
   Me.BarraEmpleados.Value = 0
   Me.BarraEmpleados.Visible = True
   i = 0
   
   MsgBox "Se procesaran un Total de " & Maximo & " Empleados"
  
  Do While Not Me.AdoEmpleados.Recordset.EOF
  
   If Me.OptTarifa.Value = True Then
 
    Me.AdoEmpleados.Recordset("TarifaHoraria") = SalarioAjuste
    Me.AdoEmpleados.Recordset.Update
   
   Else
    Me.AdoEmpleados.Recordset("SueldoPeriodo") = SalarioAjuste
    Me.AdoEmpleados.Recordset.Update
   End If
 
   i = i + 1
      Me.LblProgreso.Caption = "Procesando " & i & " de " & Maximo
   DoEvents
   Me.BarraEmpleados.Value = i
   Me.AdoEmpleados.Recordset.MoveNext
  Loop

   
     MsgBox "Proceso Terminado!!!!", vbExclamation, "Sistema de Nominas"
     Me.BarraEmpleados.Visible = False
     Me.LblProgreso.Caption = ""
     Me.txtSalario.Enabled = True
     Me.TxtSalarioAjuste.Enabled = True
 Exit Sub
TipoErrs:
   MsgBox Err.Description
End Sub

Private Sub cmdSalPts_Click()
On Error GoTo errsal

If Me.tdbcNomina.BoundText = "" Then
    MsgBox "Seleccione una nomina", vbInformation
    Me.tdbcNomina.SetFocus
    Exit Sub
ElseIf Trim(txtSalAct.Text) <> "" And val(Trim(txtSalAct.Text)) <= val(Trim(txtSalAnt.Text)) Then
    MsgBox "El monto del salario mínimo debe ser mayor al salario mínimo existente", vbInformation
    Me.txtSalAct.SetFocus
    Exit Sub
ElseIf Trim(txtPtsAct.Text) <> "" And val(Trim(txtPtsAct.Text)) <= val(Trim(txtPtsAnt.Text)) Then
    MsgBox "El valor de los puntos debe ser mayor al valor existente", vbInformation
    Me.txtPtsAct.SetFocus
    Exit Sub
ElseIf Trim(txtSalAct.Text) = Trim(txtPtsAct.Text) And Trim(txtPtsAct.Text) = "" Then
    MsgBox "Digite los valores a actualizar", vbInformation
    Me.txtSalAct.SetFocus
    Exit Sub
End If

If MsgBox("Esta seguro de realizar ésta Actualización, recuerde que no podrá revertirla", vbYesNo) = vbNo Then Exit Sub
ospSalPts.Min = 0
ospSalPts.Max = 12
ospSalPts.Value = 0

sql = "update DatosEmpresa set SalarioMinimo = " & IIf(Trim(Me.txtSalAct.Text) = "", val(Trim(Me.txtSalAnt.Text)), val(Trim(Me.txtSalAct.Text))) & ", ValorPts = " & IIf(Trim(Me.txtPtsAct.Text) = "", val(Trim(Me.txtPtsAnt.Text)), val(Trim(Me.txtPtsAct.Text))) & " where numero = 1"
ospSalPts.Value = 2
cnx.Execute sql
ospSalPts.Value = 4
sql = "UPDATE  EMP " & _
        "SET SUELDOPERIODO =    (SELECT SALORDINARIO = (SALMIN * (100 + SALPORC) / 100) + (CANTPTS * VALPTS) " & _
        "                                            FROM (SELECT SALMIN = (SELECT SALARIOMINIMO FROM DATOSEMPRESA WHERE NUMERO = 1), " & _
        "                                                                        VALPTS = (SELECT VALORPTS FROM DATOSEMPRESA WHERE NUMERO = 1), " & _
        "                                                                        SALPORC = (SELECT ISNULL(SALPORCENTAJE,0) FROM EMPLEADO WHERE CODEMPLEADO = EMP.CODEMPLEADO), " & _
        "                                                                        CANTPTS = (SELECT ISNULL(SUM(CANTPTS),0) FROM PUNTOSEMPLEADO PE INNER JOIN PUNTOS P ON PE.PUNTOS = P.ID WHERE PE.APROBADO = 1 AND EMPLEADO = EMP.CODEMPLEADO)) DAT), " & _
        "   CANTPTS = (SELECT ISNULL(SUM(CANTPTS),0) " & _
        "                        FROM PUNTOSEMPLEADO PE INNER JOIN PUNTOS P ON PE.PUNTOS = P.ID " & _
        "                        WHERE PE.APROBADO = 1 AND EMPLEADO = EMP.CODEMPLEADO) " & _
        "FROM EMPLEADO EMP WHERE EMP.CodTipoNomina = '04' AND EMP.ACTIVO = 1"
ospSalPts.Value = 6
cnx.Execute sql
ospSalPts.Value = 8
Me.AdoDatosEmpresa.Recordset.Requery
ospSalPts.Value = 10
Call SSTab1_Click(0)
ospSalPts.Value = 12
MsgBox "Actualización completada", vbInformation
ospSalPts.Value = 0
Exit Sub
errsal:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub Command1_Click()
FrmDirectorio.Show 1
End Sub

Private Sub Command2_Click()
 Dim Fecha As String, i As Double, HorasExtra As Double, Maximo As Double
  Me.SSTab1.Enabled = False
  
   Fecha = Format(Me.DTFecha.Value, "yyyy-mm-dd")
   Me.AdoDetalleNominas.RecordSource = "SELECT   Nomina.FechaNominaINI, Nomina.FechaNomina, DetalleNomina.Ajuste, DetalleNomina.HorasExtras FROM  DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina WHERE (Nomina.FechaNomina <= CONVERT(DATETIME, '" & Fecha & "', 102))"
   Me.AdoDetalleNominas.Refresh
   
   Maximo = Me.AdoDetalleNominas.Recordset.RecordCount
   Me.Barra.Min = 0
   Me.Barra.Max = Maximo
   Me.Barra.Value = 0
   Me.Barra.Visible = True
   i = 0
   Me.LblProcesos.Visible = True
   
   Do While Not Me.AdoDetalleNominas.Recordset.EOF
    
    
     HorasExtra = Me.AdoDetalleNominas.Recordset("HorasExtras")
     
     Me.AdoDetalleNominas.Recordset("Ajuste") = HorasExtra
     Me.AdoDetalleNominas.Recordset.Update
     
   
     i = i + 1
     DoEvents
     Me.Barra.Value = i
     Me.LblProcesos.Caption = "Procesando " & i & " de " & Maximo
     Me.AdoDetalleNominas.Recordset.MoveNext
   Loop


  Me.SSTab1.Enabled = True
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

Private Sub Command4_Click()
Dim SalarioBasico As Double, SalarioAjuste As Double, Maximo As Double, i As Double
Dim SalarioDesde As Double, SalarioHasta As Double, Porciento As Double

On Error GoTo TipoErrs

If Me.txtDesde.Text = "" Then
 MsgBox "Se necesita el salario de Filtro", vbCritical, "Zeus Nominas"
 Exit Sub
End If

If Me.txtHasta.Text = "" Then
 MsgBox "Se necesita el salario de Filtro", vbCritical, "Zeus Nominas"
 Exit Sub
End If

If Me.TxtPorciento.Text = "" Then
 MsgBox "Se necesita el salario de Filtro", vbCritical, "Zeus Nominas"
 Exit Sub
End If


  SalarioDesde = Me.txtDesde.Text
  SalarioHasta = Me.txtHasta.Text
  Porciento = Me.TxtPorciento.Text

  Porciento = 1 + (Porciento / 100)

'    Me.AdoEmpleados.RecordSource = "SELECT CodEmpleado, SueldoPeriodo, TarifaHoraria, PorcentajeComision, SalarioMinimo From Empleado Where (SueldoPeriodo " & Simbolo & "  " & SalarioBasico & ") AND (Activo = 1)"
    Me.AdoEmpleados.RecordSource = "SELECT Empleado.* From Empleado WHERE (SueldoPeriodo BETWEEN " & SalarioDesde & " AND " & SalarioHasta & ") AND (Activo = 1)"
    Me.AdoEmpleados.Refresh
  
   Maximo = Me.AdoEmpleados.Recordset.RecordCount
   Me.ProgressBar.Min = 0
   Me.ProgressBar.Max = Maximo
   Me.ProgressBar.Value = 0
   Me.ProgressBar.Visible = True
   i = 0
   
   MsgBox "Se procesaran un Total de " & Maximo & " Empleados"
   
 Do While Not Me.AdoEmpleados.Recordset.EOF
  
    SalarioBasico = Me.AdoEmpleados.Recordset("SueldoPeriodo") * Porciento

    Me.AdoEmpleados.Recordset("SueldoPeriodo") = SalarioBasico
    Me.AdoEmpleados.Recordset.Update
  
 
   i = i + 1
      Me.LblProgreso.Caption = "Procesando " & i & " de " & Maximo
   DoEvents
   Me.ProgressBar.Value = i
   Me.AdoEmpleados.Recordset.MoveNext
 Loop

   
     MsgBox "Proceso Terminado!!!!", vbExclamation, "Sistema de Nominas"
     Me.ProgressBar.Visible = False
     Me.LblProgreso.Caption = ""
     Me.txtSalario.Enabled = True
     Me.TxtSalarioAjuste.Enabled = True
   
  Exit Sub
TipoErrs:
   MsgBox Err.Description
   

End Sub

Private Sub Command5_Click()
   Dim CodEmpleado1 As String, Maximo As Double, CodTurno As String


    Me.AdoEmpleados.RecordSource = "SELECT Empleado.* From Empleado WHERE (Activo = 1)"
    Me.AdoEmpleados.Refresh
  
   Maximo = Me.AdoEmpleados.Recordset.RecordCount
   Me.ProgressBar1.Min = 0
   Me.ProgressBar1.Max = Maximo
   Me.ProgressBar1.Value = 0
   Me.ProgressBar1.Visible = True
   i = 0
   
   MsgBox "Se procesaran un Total de " & Maximo & " Empleados"
   
    Do While Not Me.AdoEmpleados.Recordset.EOF
    
           If Not IsNull(Me.AdoEmpleados.Recordset("CodEmpleado1")) Then
           CodEmpleado1 = Me.AdoEmpleados.Recordset("CodEmpleado1")
           End If
    
           Me.DtaHorarioEmpleado.RecordSource = "SELECT CodEmpleado, LEntrada, LSalida, MEntrada, MSalida, MCEntrada, MCSalida, JEntrada, JSalida, VEntrada, VSalida, TComida, TurnoLunes,TurnoMartes , TurnoMiercoles, TurnoJueves, TurnoViernes, TurnoSabado, TurnoDomingo, SEntrada, SSalida, DEntrada, DSalida From dbo.HorarioEmpleado WHERE(CodEmpleado ='" & CodEmpleado1 & "')"
           Me.DtaHorarioEmpleado.Refresh
           If Me.DtaHorarioEmpleado.Recordset.EOF Then
             Me.DtaTurnos.Refresh
             If Not Me.DtaTurnos.Recordset.EOF Then
               CodTurno = Me.DtaTurnos.Recordset("CodTurno")
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
                Me.DtaHorarioEmpleado.Recordset("TurnoLunes") = "Diurno"
                Me.DtaHorarioEmpleado.Recordset("TurnoMartes") = "Diurno"
                Me.DtaHorarioEmpleado.Recordset("TurnoMiercoles") = "Diurno"
                Me.DtaHorarioEmpleado.Recordset("TurnoJueves") = "Diurno"
                Me.DtaHorarioEmpleado.Recordset("TurnoViernes") = "Diurno"
                Me.DtaHorarioEmpleado.Recordset("TurnoSabado") = "Diurno"
                Me.DtaHorarioEmpleado.Recordset("TurnoDomingo") = "Diurno"
               Me.DtaHorarioEmpleado.Recordset("SEntrada") = Me.DtaTurnos.Recordset("SEntrada")
               Me.DtaHorarioEmpleado.Recordset("SSalida") = Me.DtaTurnos.Recordset("SEntrada")
               Me.DtaHorarioEmpleado.Recordset("DEntrada") = Me.DtaTurnos.Recordset("SEntrada")
               Me.DtaHorarioEmpleado.Recordset("DSalida") = Me.DtaTurnos.Recordset("SEntrada")
        
             Me.DtaHorarioEmpleado.Recordset.Update
             End If
           End If
    
    
    
            i = i + 1
               Me.LblProgreso.Caption = "Procesando " & i & " de " & Maximo
            DoEvents
            Me.ProgressBar1.Value = i
            Me.AdoEmpleados.Recordset.MoveNext
    Loop
End Sub

Private Sub Command6_Click()
   Dim CodEmpleado1 As String, Maximo As Double, CodTurno As String


    Me.AdoEmpleados.RecordSource = "SELECT Empleado.* From Empleado WHERE (Activo = 1)"
    Me.AdoEmpleados.Refresh
  
   Maximo = Me.AdoEmpleados.Recordset.RecordCount
   Me.ProgressBar2.Min = 0
   Me.ProgressBar2.Max = Maximo
   Me.ProgressBar2.Value = 0
   Me.ProgressBar2.Visible = True
   i = 0
   
   MsgBox "Se procesaran un Total de " & Maximo & " Empleados"
   
    Do While Not Me.AdoEmpleados.Recordset.EOF
    
           CodEmpleado1 = Me.AdoEmpleados.Recordset("CodEmpleado1")
           
               i = i + 1          '/////////////////////////////////////////////////////////////////////////////////////
        '//////////////////////////////////AGREGA EMPLEADOS EN LA TABLA USUARIOS ////////////
        '////////////////////////////////////////////////////////////////////////////////////
                    Dim NumeroUser As Double
                     Me.AdoUser.ConnectionString = Conexion
                     Me.AdoUser.RecordSource = "SELECT * From Userinfo WHERE (IDCard = '" & CodEmpleado1 & "')"
                     Me.AdoUser.Refresh
                    If Me.AdoUser.Recordset.EOF Then
                       NumeroUser = ConsecutivoUser(CodEmpleado1)
                       Me.AdoUser.Recordset.AddNew
                         Me.AdoUser.Recordset("Userid") = NumeroUser
                         Me.AdoUser.Recordset("Name") = Me.AdoEmpleados.Recordset("Nombre1") + " " + Me.AdoEmpleados.Recordset("Nombre2") + " " + Me.AdoEmpleados.Recordset("Apellido1") + " " + Me.AdoEmpleados.Recordset("Apellido2")
                         Me.AdoUser.Recordset("IDCard") = CodEmpleado1
                       Me.AdoUser.Recordset.Update
                    
                    End If


            Me.LblProgreso.Caption = "Procesando " & i & " de " & Maximo
            DoEvents
            Me.ProgressBar2.Value = i
            Me.AdoEmpleados.Recordset.MoveNext
    Loop
End Sub

Private Sub Command7_Click()
   Dim CodEmpleado1 As String, Maximo As Double, CodTurno As String, CodEmpleado As Double


    Me.AdoEmpleados.RecordSource = "SELECT Empleado.* From Empleado WHERE (Activo = 1)"
    Me.AdoEmpleados.Refresh
  
   Maximo = Me.AdoEmpleados.Recordset.RecordCount
   Me.ProgressBar1.Min = 0
   Me.ProgressBar1.Max = Maximo
   Me.ProgressBar1.Value = 0
   Me.ProgressBar1.Visible = True
   i = 0
   
   MsgBox "Se procesaran un Total de " & Maximo & " Empleados"
   
    Do While Not Me.AdoEmpleados.Recordset.EOF
    
           If Not IsNull(Me.AdoEmpleados.Recordset("CodEmpleado")) Then
           CodEmpleado = Me.AdoEmpleados.Recordset("CodEmpleado")
           End If
    
           Me.DtaHorarioEmpleado.RecordSource = "SELECT  Historico.* From Historico WHERE (Codempleado = " & CodEmpleado & ")"
           Me.DtaHorarioEmpleado.Refresh
           If Me.DtaHorarioEmpleado.Recordset.EOF Then

               Me.DtaHorarioEmpleado.Recordset.AddNew
               Me.DtaHorarioEmpleado.Recordset("Codempleado") = CodEmpleado1
                Me.DtaHorarioEmpleado.Recordset("FechaNacimiento") = Format(Now, "dd/mm/yyyy")
                Me.DtaHorarioEmpleado.Recordset("FechaContrato") = Format(Now, "dd/mm/yyyy")
                Me.DtaHorarioEmpleado.Recordset("FechaContratoVac") = Format(Now, "dd/mm/yyyy")
        
             Me.DtaHorarioEmpleado.Recordset.Update
           End If

    
    
    
            i = i + 1
               Me.LblProgreso.Caption = "Procesando " & i & " de " & Maximo
            DoEvents
            Me.ProgressBar1.Value = i
            Me.AdoEmpleados.Recordset.MoveNext
    Loop
End Sub

Private Sub Form_Load()
On Error GoTo TipoErrs
Dim Destino As String
Me.Top = 1500
Me.Left = 4500

Me.DTFecha.Value = Format(Now, "dd/mm/yyyy")

With Me.DtaTurnos
   .ConnectionString = Conexion
   .RecordSource = "Turno"
   .Refresh
End With

With Me.DtaHorarioEmpleado
   .ConnectionString = Conexion
End With

With Me.AdoUser
   .ConnectionString = Conexion
End With

With Me.DtaControles
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Controles"
   .Refresh
End With

With Me.AdoDetalleNominas
   .ConnectionString = Conexion
End With

With Me.AdoEmpleados
   .ConnectionString = Conexion
End With

With Me.AdoDatosEmpresa
   .ConnectionString = Conexion
   .RecordSource = "DatosEmpresa"
   .Refresh
End With

If Not Me.AdoDatosEmpresa.Recordset.EOF Then
If Not IsNull(Me.AdoDatosEmpresa.Recordset("RutaFoto")) Then
 Me.TxtRutaFoto.Text = Me.AdoDatosEmpresa.Recordset("RutaFoto")
End If
 Me.TxtNombreEmpresa.Text = Me.AdoDatosEmpresa.Recordset("NombreEmpresa")
 Me.TxtRucEmpresa.Text = Me.AdoDatosEmpresa.Recordset("NumeroRUC")
 Me.TxtDireccionEmpresa.Text = Me.AdoDatosEmpresa.Recordset("Direccion")
 Me.TxtTelefono.Text = Me.AdoDatosEmpresa.Recordset("Telefono")
 Me.TxtFax.Text = Me.AdoDatosEmpresa.Recordset("Fax")
 Me.TxtEmail.Text = Me.AdoDatosEmpresa.Recordset("Email")
 Me.TxtRutaLogo.Text = Me.AdoDatosEmpresa.Recordset("RutaLogo")
 
 If Not IsNull(Me.AdoDatosEmpresa.Recordset("HorasExtra")) Then
   Me.TxtValorHorasExtra.Text = Me.AdoDatosEmpresa.Recordset("HorasExtra")
 Else
   Me.TxtValorHorasExtra.Text = 2
 End If
 
 If Me.AdoDatosEmpresa.Recordset("Calcular7mo") = True Then
   Me.Chk7mo.Value = 0
 Else
   Me.Chk7mo.Value = 1
 End If
 
 If DtaControles.Recordset("CalcularRedondeado") = True Then
    Me.ckRedondeo.Value = 1
 Else
    Me.ckRedondeo.Value = 0
 End If
 
 If Not IsNull(Me.AdoDatosEmpresa.Recordset("TipoCalculoIR")) Then
  If Me.AdoDatosEmpresa.Recordset("TipoCalculoIR") = "Calcular Ajustando IR" Then
    Me.Option1.Value = True
  Else
    Me.Option2.Value = True
  End If
 
 End If
 
 If Me.AdoDatosEmpresa.Recordset("MetodoVacaciones") = "Vacaciones Semestrales" Then
  Me.OptVacacionesSemestrales.Value = True
 ElseIf Me.AdoDatosEmpresa.Recordset("MetodoVacaciones") = "Vacaciones Mensuales" Then
  Me.OptVacacionesMensuales.Value = True
 End If
 
 If Not IsNull(Me.AdoDatosEmpresa.Recordset("ConexionSistemaContable")) Then
   Me.TxtConexionString.Text = Me.AdoDatosEmpresa.Recordset("ConexionSistemaContable")
 End If
 
 If Not IsNull(Me.AdoDatosEmpresa.Recordset("FormatoColilla")) Then
 Me.CmbColillas.Text = Me.AdoDatosEmpresa.Recordset("FormatoColilla")
 End If
 If Not IsNull(Me.AdoDatosEmpresa.Recordset("FormatoNomina")) Then
  Me.CmbNominas.Text = Me.AdoDatosEmpresa.Recordset("FormatoNomina")
  If Me.AdoDatosEmpresa.Recordset("FormatoNomina") = "Nomina Bono Produccion" Then
    Me.Frame6.Visible = True
  Else
    Me.Frame6.Visible = False
  End If
 End If
End If

    
    
    
    If Me.TxtRutaLogo.Text <> "" Then
        If (Dir(Me.TxtRutaLogo.Text, vbDirectory) <> "") Then
          Me.ImgLogo.Picture = LoadPicture(Me.TxtRutaLogo.Text)
        Else
'          Destino = RutaFoto + "Zw.bmp"
'          Me.ImgLogo.Picture = LoadPicture(Destino)
          MsgBox "La Ruta del LOGO: " & Me.TxtRutaLogo & " ES INCORRECTA"
          
        End If
       CMRutaFoto.InitDir = Me.TxtRutaLogo.Text

    End If
        
    If Me.TxtRutaFoto.Text <> "" Then
        If (Dir(Me.TxtRutaFoto.Text & "\", vbDirectory) <> "") Then
'          Me.ImgLogo.Picture = LoadPicture(Me.TxtRutaLogo.Text)
        Else
'          Destino = RutaFoto + "Zw.bmp"
'          Me.ImgLogo.Picture = LoadPicture(Destino)
          MsgBox "La Ruta De la FOTO : " & RutaFoto & " ES INCORRECTA"
          
        End If
'       CMRutaFoto.InitDir = Me.TxtRutaLogo.Text

    End If
    



DtaControles.Refresh
If DtaControles.Recordset("DiasMes") = "30" Then
   Opt30.Value = True
ElseIf DtaControles.Recordset("DiasMes") = "30.41667" Then
   OptExacto.Value = True
Else
   Me.Opt25.Value = True
End If

If DtaControles.Recordset("DiasSemana") = 7 Then
   Opt7.Value = True
Else
   OPT6.Value = True
End If

If DtaControles.Recordset("verificartasa") = True Then
   ChkTasa.Value = 1
Else
  ChkTasa.Value = 0
End If

If DtaControles.Recordset("SalarioPromedioReal") = True Then
   Me.OptSalarioPromedioReal.Value = True
Else
  Me.OptSalarioPromedioPeriodo.Value = True
End If

If DtaControles.Recordset("AntiguedadMenor") = True Then
   Me.ChkAntiguedadMenor.Value = 1
Else
   Me.ChkAntiguedadMenor.Value = 0
End If

If SSTab1.Tab = 4 Then Call SSTab1_Click(0)

Exit Sub
TipoErrs:
MsgBox Err.Description
End Sub

Private Sub Opt25Dias_Click()

End Sub

Private Sub OptSalario_Click()
 If Me.OptTarifa.Value = True Then
   Me.LblNombre1.Caption = "Tarifa Menor que"
   Me.LblNombre2.Caption = "Ajustar Por"
Else
   Me.LblNombre1.Caption = "Salario Menor que"
   Me.LblNombre2.Caption = "Ajustar Por"
 End If
End Sub

Private Sub OptTarifa_Click()
 If Me.OptTarifa.Value = True Then
   Me.LblNombre1.Caption = "Tarifa Menor que"
   Me.LblNombre2.Caption = "Ajustar Por"
Else
   Me.LblNombre1.Caption = "Salario Menor que"
   Me.LblNombre2.Caption = "Ajustar Por"
 End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo errsst
Select Case SSTab1.Tab
Case 4
    If cnx.State = adStateClosed Then
    '    sql = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PRUEBA;Data Source=WEBMASTER\SQL2005"
        cnx.ConnectionString = Conexion
        cnx.Open
    End If
    
    'NOMINAS
    sql = "SELECT [Nomina], [CodTipoNomina] From [dbo].[TipoNomina]"
    With rs
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .Open sql, cnx, adOpenDynamic, adLockOptimistic
    End With
    
    Me.tdbcNomina.RowSource = rs
    Me.tdbcNomina.BoundColumn = "CodTipoNomina"
    Me.tdbcNomina.Refresh
    Me.tdbcNomina.Columns(1).Visible = False
    Me.tdbcNomina.Text = ""
    
    txtSalAnt.Text = AdoDatosEmpresa.Recordset!SalarioMinimo
    txtPtsAnt.Text = AdoDatosEmpresa.Recordset!valorpts
    txtSalAct.Text = ""
    txtPtsAct.Text = ""
Case Else
    Set cnx = Nothing
    Set rs = Nothing
End Select

Exit Sub
errsst:
    MsgBox Err.Description, vbInformation
End Sub

Private Sub tdbcNomina_ItemChange()
    txtSalAct.Text = ""
    txtPtsAct.Text = ""
End Sub
