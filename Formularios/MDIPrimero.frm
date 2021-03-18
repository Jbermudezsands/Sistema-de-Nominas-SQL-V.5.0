VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.TaskPanel.v12.0.0.Demo.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "Codejock.CommandBars.v12.0.0.Demo.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "Codejock.DockingPane.v12.0.0.Demo.ocx"
Begin VB.MDIForm MDIPrimero 
   BackColor       =   &H80000009&
   Caption         =   " Zeus Nóminas"
   ClientHeight    =   9315
   ClientLeft      =   -75
   ClientTop       =   150
   ClientWidth     =   13590
   HelpContextID   =   1
   Icon            =   "MDIPrimero.frx":0000
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "MDIPrimero.frx":0442
   NegotiateToolbars=   0   'False
   OLEDropMode     =   1  'Manual
   Picture         =   "MDIPrimero.frx":0884
   WindowState     =   2  'Maximized
   Begin XtremeTaskPanel.TaskPanel wndTaskPanel 
      Align           =   4  'Align Right
      Height          =   1035
      Left            =   11250
      TabIndex        =   16
      Top             =   855
      Visible         =   0   'False
      Width           =   2340
      _Version        =   786432
      _ExtentX        =   4128
      _ExtentY        =   1826
      _StockProps     =   64
      Animation       =   1
      ItemLayout      =   3
      HotTrackStyle   =   1
   End
   Begin MSAdodcLib.Adodc DtaSuspenciones 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
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
      Caption         =   "DtaSuspenciones"
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
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   0
      ScaleHeight     =   425.455
      ScaleMode       =   0  'User
      ScaleWidth      =   13560
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   13590
      Begin SmartButtonProject.SmartButton CmdMovimiento 
         Height          =   930
         Left            =   5520
         TabIndex        =   2
         ToolTipText     =   "Movimientos de la Nomina"
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1640
         ForeColor       =   8388608
         Caption         =   "Horas Extra"
         Picture         =   "MDIPrimero.frx":717BD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdEmpleado 
         Height          =   930
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Registro de Empleados"
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1640
         BackColor       =   -2147483644
         ForeColor       =   8388608
         Caption         =   "Empleados"
         Picture         =   "MDIPrimero.frx":7330F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdActivar 
         Height          =   930
         Left            =   3360
         TabIndex        =   4
         ToolTipText     =   "Activar Nominas"
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1640
         ForeColor       =   8388608
         Caption         =   "Act.Nom"
         Picture         =   "MDIPrimero.frx":74E61
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton Cmd13vo 
         Height          =   930
         Left            =   9840
         TabIndex        =   5
         ToolTipText     =   "Calculo del 13vo mes y Vacaciones"
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1640
         ForeColor       =   8388608
         Caption         =   "13vo-Vac"
         Picture         =   "MDIPrimero.frx":769B3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdCalcular 
         Height          =   930
         Left            =   6600
         TabIndex        =   6
         ToolTipText     =   "Calculo de la Nomina"
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1640
         ForeColor       =   8388608
         Caption         =   "Calc Nom"
         Picture         =   "MDIPrimero.frx":78505
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdDespido 
         Height          =   930
         Left            =   10920
         TabIndex        =   7
         ToolTipText     =   "Calculo de Despido y Renuncia"
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1640
         ForeColor       =   8388608
         Caption         =   "Despido"
         Picture         =   "MDIPrimero.frx":7A057
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdSubsidio 
         Height          =   930
         Left            =   7680
         TabIndex        =   8
         ToolTipText     =   "Calculo de la nomina de Subsidio"
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1640
         ForeColor       =   8388608
         Caption         =   "Subsidio"
         Picture         =   "MDIPrimero.frx":7BBA9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdUsuario 
         Height          =   930
         Left            =   12000
         TabIndex        =   9
         ToolTipText     =   "Registro de Usuarios del Sistema"
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1640
         ForeColor       =   8388608
         Caption         =   "Usuarios"
         Picture         =   "MDIPrimero.frx":7D6FB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdSalir 
         Height          =   930
         Left            =   14160
         TabIndex        =   10
         ToolTipText     =   "Boton de Salir del sistema"
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1640
         ForeColor       =   8388608
         Caption         =   "Salir"
         Picture         =   "MDIPrimero.frx":7F24D
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
      Begin SmartButtonProject.SmartButton CmdAdelanto 
         Height          =   930
         Left            =   8760
         TabIndex        =   11
         ToolTipText     =   "Adelanto 13vo mes y Vacaciones"
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1640
         ForeColor       =   8388608
         Caption         =   "Adel. 13vo"
         Picture         =   "MDIPrimero.frx":8006F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdInss 
         Height          =   930
         Left            =   1200
         TabIndex        =   12
         ToolTipText     =   "Tabla del INSS Y EL IR"
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1640
         ForeColor       =   8388608
         Caption         =   "INSS / IR"
         Picture         =   "MDIPrimero.frx":81BC1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdPeriodos 
         Height          =   930
         Left            =   2280
         TabIndex        =   13
         ToolTipText     =   "Periodo Inss"
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1640
         ForeColor       =   8388608
         Caption         =   "Periodos"
         Picture         =   "MDIPrimero.frx":83713
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdProduccion 
         Height          =   930
         Left            =   4440
         TabIndex        =   14
         ToolTipText     =   "Referencias"
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1640
         ForeColor       =   8388608
         Caption         =   "Mov Prod"
         Picture         =   "MDIPrimero.frx":85265
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
      Begin SmartButtonProject.SmartButton CmdRespaldar 
         Height          =   930
         Left            =   13080
         TabIndex        =   15
         ToolTipText     =   "Realizar Respaldos"
         Top             =   45
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   1640
         ForeColor       =   8388608
         Caption         =   "Respaldar"
         Picture         =   "MDIPrimero.frx":86DB7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   7
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1080
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DtaControles 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   6720
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
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
   Begin MSAdodcLib.Adodc DtaNAcceso 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      Top             =   6105
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
      _ExtentY        =   1085
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SistemaNominas"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaNAcceso"
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
   Begin MSAdodcLib.Adodc DtaTasa 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   5295
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
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
      Caption         =   "DtaTasa"
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
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   7095
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
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
   Begin MSAdodcLib.Adodc DtaEmpresa 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      Top             =   5625
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
      _ExtentY        =   847
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
      Caption         =   "DtaEmpresa"
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
      Left            =   2160
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   7950
      WhatsThisHelpID =   1
      Width           =   13590
      _ExtentX        =   23971
      _ExtentY        =   714
      SimpleText      =   "Programa Bajo Licencia de Juan"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Picture         =   "MDIPrimero.frx":88909
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7937
            MinWidth        =   7937
            Text            =   "Licencia: Juan"
            TextSave        =   "Licencia: Juan"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "MAYÚS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Object.Width           =   1393
            MinWidth        =   1393
            TextSave        =   "NÚM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "08:09 p.m."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "MDIPrimero.frx":88C23
   End
   Begin MSCommLib.MSComm mscReloj2 
      Left            =   5520
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm mscReloj 
      Left            =   4320
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6000
      OleObjectBlob   =   "MDIPrimero.frx":88F3D
      Top             =   3720
   End
   Begin MSAdodcLib.Adodc DtaConsulta 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   1890
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
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
   Begin MSComctlLib.ImageList imlToolbarIcons2 
      Left            =   9240
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E476A
            Key             =   ""
            Object.Tag             =   "119"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E4D04
            Key             =   ""
            Object.Tag             =   "113"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E529E
            Key             =   ""
            Object.Tag             =   "128"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E5838
            Key             =   ""
            Object.Tag             =   "115"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E5BD2
            Key             =   ""
            Object.Tag             =   "130"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E616C
            Key             =   ""
            Object.Tag             =   "160"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E6706
            Key             =   ""
            Object.Tag             =   "116"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E6CA0
            Key             =   ""
            Object.Tag             =   "300"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E723A
            Key             =   ""
            Object.Tag             =   "118"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E77D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E9B56
            Key             =   ""
            Object.Tag             =   "204"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2EA0F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2EC472
            Key             =   ""
            Object.Tag             =   "129"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2ECA0C
            Key             =   ""
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2ECFA6
            Key             =   ""
            Object.Tag             =   "205"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2ED540
            Key             =   ""
            Object.Tag             =   "1331"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2EDADA
            Key             =   ""
            Object.Tag             =   "1311"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2EF7E4
            Key             =   ""
            Object.Tag             =   "131"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2EFD7E
            Key             =   ""
            Object.Tag             =   "0"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2F0318
            Key             =   ""
            Object.Tag             =   "133"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2F08B2
            Key             =   ""
            Object.Tag             =   "132"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2F0E4C
            Key             =   ""
            Object.Tag             =   "134"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2F13E6
            Key             =   ""
            Object.Tag             =   "140111"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2F1980
            Key             =   ""
            Object.Tag             =   "139"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2F1F1A
            Key             =   ""
            Object.Tag             =   "320"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2F24B4
            Key             =   ""
            Object.Tag             =   "136"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   8400
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2F2A4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2F2B60
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2F2CB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2F2DC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2F2ED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2F2FE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc DtaIr 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   3345
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
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
   Begin MSAdodcLib.Adodc AdoReportes 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   2970
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
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
   Begin MSAdodcLib.Adodc DtaConsulta2 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   2595
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
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
   Begin MSAdodcLib.Adodc AdoConsulta 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      Top             =   375
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
      _ExtentY        =   847
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
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      Top             =   8835
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
      _ExtentY        =   847
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
      Caption         =   "DtaEmpresa"
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
   Begin MSAdodcLib.Adodc AdoConsultaEasyWay 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      Top             =   8355
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
      _ExtentY        =   847
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
      Caption         =   "DtaEmpresa"
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
   Begin MSAdodcLib.Adodc AdoDispositivos 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      Top             =   7470
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
      _ExtentY        =   847
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
      Caption         =   "AdoDispositivos"
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
   Begin MSAdodcLib.Adodc AdoConexion 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
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
      Caption         =   "AdoConexion"
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
   Begin MSAdodcLib.Adodc AdoTasaContabilidad 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   2265
      Visible         =   0   'False
      Width           =   13590
      _ExtentX        =   23971
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
      Caption         =   "AdoTasaContabilidad"
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
   Begin XtremeDockingPane.DockingPane DockingPaneManager 
      Left            =   7920
      Top             =   3600
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
      ScaleMode       =   1
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   6840
      Top             =   3360
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ScaleMode       =   2
      VisualTheme     =   2
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6720
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483648
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   36
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F30FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F3D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F499E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F55F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F6242
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F6E94
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F7AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F8738
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F938A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F9FDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2FAC2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2FB880
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2FC4D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2FD124
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2FDD76
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2FE9C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2FF61A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":30026C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":300EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":301B10
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":302762
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":3033B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":304006
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":304C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":3058AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":3064FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":30714E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":307DA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":3089F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":309644
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":30A296
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":30AEE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":30BB3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":30C78C
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":30D3DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":30E030
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PopupControl PopupControl1 
      Left            =   7200
      Top             =   4920
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
      VisualTheme     =   4
   End
   Begin VB.Menu archivo 
      Caption         =   "&Archivo"
      HelpContextID   =   1
      Visible         =   0   'False
      Begin VB.Menu Compañias 
         Caption         =   "&Niveles"
         HelpContextID   =   2
         Begin VB.Menu Niveles 
            Caption         =   "&Editar Niveles"
         End
      End
      Begin VB.Menu mnuemple 
         Caption         =   "&Empleados"
         Begin VB.Menu empleados 
            Caption         =   "&Empleados"
            HelpContextID   =   12
         End
         Begin VB.Menu mnususpen 
            Caption         =   "&Suspenciones"
         End
         Begin VB.Menu mnuhistorial 
            Caption         =   "&Historial Salarial"
         End
      End
      Begin VB.Menu tablas 
         Caption         =   "&Tablas"
         HelpContextID   =   3
         Begin VB.Menu departamento 
            Caption         =   "&Departamentos"
            HelpContextID   =   24
         End
         Begin VB.Menu Cargo 
            Caption         =   "&Cargo"
            HelpContextID   =   25
         End
         Begin VB.Menu incapacidad 
            Caption         =   "&Incapacidades"
            HelpContextID   =   26
            Begin VB.Menu TipoIncapacidad 
               Caption         =   "&Tipo Incapacidad"
               HelpContextID   =   30
            End
            Begin VB.Menu incapacidades 
               Caption         =   "&Incapacidad"
               HelpContextID   =   31
            End
         End
         Begin VB.Menu MnuIncentivos 
            Caption         =   "Tipos de Incentivos"
         End
         Begin VB.Menu MnuDeducciones 
            Caption         =   "Tipo de Deducciones"
         End
         Begin VB.Menu mnuSubsidio 
            Caption         =   "Tipos de Subsidios"
         End
         Begin VB.Menu mnucomisiones 
            Caption         =   "Tipo de Comisión"
         End
         Begin VB.Menu mnutipodestajo 
            Caption         =   "Tipo de Destajo"
         End
         Begin VB.Menu mnutipodivision 
            Caption         =   "&Divisiones de Nómina"
         End
         Begin VB.Menu TipoNomina 
            Caption         =   "T&ipo Nómina"
            HelpContextID   =   28
         End
         Begin VB.Menu Inss 
            Caption         =   "&Tablas INSS, IR"
            HelpContextID   =   29
         End
         Begin VB.Menu mnulistnomina 
            Caption         =   "Listado de Nóminas elaboradas"
         End
      End
      Begin VB.Menu abrir 
         Caption         =   "&Abrir/Cerrar Backup"
      End
      Begin VB.Menu A 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu salir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu proceso 
      Caption         =   "&Proceso"
      Visible         =   0   'False
      Begin VB.Menu mnumovnomina 
         Caption         =   "Movimientos de Nómina"
      End
      Begin VB.Menu mnuactnomina 
         Caption         =   "Activar Nómina"
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCalcNomina 
         Caption         =   "Calcular Nómina"
      End
      Begin VB.Menu mnumes13Vaca 
         Caption         =   "Calcular e Imprimir el 13vo mes Y/o Vacaciones"
      End
      Begin VB.Menu mnunomsubsidios 
         Caption         =   "Nómina de Subsidios"
      End
      Begin VB.Menu mnuraya2 
         Caption         =   "-"
      End
      Begin VB.Menu mnudesrenu 
         Caption         =   "Despidos y Renuncias"
      End
   End
   Begin VB.Menu mnureg 
      Caption         =   "Re&gistro"
      Visible         =   0   'False
      Begin VB.Menu mnuregentsal 
         Caption         =   "&Entradas y Salidas"
      End
      Begin VB.Menu MnuExtrsFaktas 
         Caption         =   "&Calcular Horas Extras o Faltas"
      End
   End
   Begin VB.Menu Opciones 
      Caption         =   "&Opciones"
      HelpContextID   =   1
      Visible         =   0   'False
      Begin VB.Menu Usuarios 
         Caption         =   "&Usuarios"
         HelpContextID   =   6
      End
      Begin VB.Menu CambiaClave 
         Caption         =   "&Registro Moneda"
      End
      Begin VB.Menu Calculadora 
         Caption         =   "&Calculadora"
      End
      Begin VB.Menu Informa 
         Caption         =   "&Informa Usuario"
      End
      Begin VB.Menu m 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu Exportar 
         Caption         =   "Exportacion de Datos"
      End
      Begin VB.Menu Importar 
         Caption         =   "Importacion de Datos"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu Barra 
         Caption         =   "&Barra de Herramientas"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnubarraestado 
         Caption         =   "Barra de &Estado"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu Reportes 
      Caption         =   "&Reportes"
      Visible         =   0   'False
      Begin VB.Menu RGenerales 
         Caption         =   "Reportes &Generales"
      End
      Begin VB.Menu mnulstreports 
         Caption         =   "Reportes &Empleados"
      End
      Begin VB.Menu RDeducciones 
         Caption         =   "Reporte &Deducciones"
      End
   End
   Begin VB.Menu mnucontrol 
      Caption         =   "&Controles"
      Visible         =   0   'False
      Begin VB.Menu mnuctrol2 
         Caption         =   "Controles &Personalizados"
      End
   End
   Begin VB.Menu ventanas 
      Caption         =   "&Ventanas"
      Visible         =   0   'False
      Begin VB.Menu MCascada 
         Caption         =   "C&ascada"
      End
      Begin VB.Menu mosaico 
         Caption         =   "&Mosaico"
      End
      Begin VB.Menu Organizar 
         Caption         =   "&Organizar Iconos"
      End
   End
   Begin VB.Menu ayuda 
      Caption         =   "&Ayuda"
      Visible         =   0   'False
      Begin VB.Menu Contendido 
         Caption         =   "&Contenido"
      End
      Begin VB.Menu ComoUsar 
         Caption         =   "&Como Usar la Ayuda"
      End
      Begin VB.Menu soporte 
         Caption         =   "&Soporte Tecnico"
      End
      Begin VB.Menu Acerca 
         Caption         =   "&Acerca del Sistema de Nominas"
      End
   End
End
Attribute VB_Name = "MDIPrimero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PreguntaSalir As Boolean
 
Public Sub CargarInterfaz()
 
    
    CommandBarsGlobalSettings.App = App
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)

      
    Dim Workspace  As TabWorkspace
    Set Workspace = CommandBars.ShowTabWorkspace(True)
    Workspace.ThemedBackColor = False
    Workspace.PaintManager.ShowIcons = False
    
'    Dim Pane1 As Pane
'    Set Pane1 = DockingPaneManager.CreatePane(1, 154, 120, DockLeftOf, Nothing)
'    Pane1.Title = "Navegador"
'    Pane1.Options = PaneNoCloseable
'    Pane1.Select
    
  
    CommandBars.Options.KeyboardCuesShow = xtpKeyboardCuesShowWindowsDefault

    CommandBars.EnableCustomization True

    DockingPaneManager.SetCommandBars CommandBars
    DockingPaneManager.ImageList = Me.imlPaneIcons
End Sub

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

On Error GoTo TipoErrs

  Dim Directorio As String
  Dim Año1 As String, Año2 As String, AÑO3 As String
  
  



      Select Case Control.Id
        Case 1300: Unload Me
        Case 1700:
                  MDIPrimero.MousePointer = 11
                   frmEmpleado.Show
                  MDIPrimero.MousePointer = 0
        Case 1701:
                 MDIPrimero.MousePointer = 11
                 FrmActivarNomina.Show
                 MDIPrimero.MousePointer = 0
        Case 1702:
                 MDIPrimero.MousePointer = 11
                 FrmProduccion.Show 1
                 MDIPrimero.MousePointer = 0
        Case 1703:
                 MDIPrimero.MousePointer = 11
                 FrmMovimientos.Show
                 MDIPrimero.MousePointer = 0
        Case 1704:
                 MDIPrimero.MousePointer = 11
                 FrmCalcularNomina.Show 1
                 MDIPrimero.MousePointer = 0
        Case 1705:
                 MDIPrimero.MousePointer = 11
                 MDIPrimero.DtaEmpresa.Refresh
                 If MDIPrimero.DtaEmpresa.Recordset("MetodoVacaciones") = "Vacaciones Semestrales" Then
                   Frm13VacaMes.Show 1
                 Else
                   Frm13Vaca.Show 1
                 End If
                 MDIPrimero.MousePointer = 0
        Case 1706:
                 MDIPrimero.MousePointer = 11
                 FrmBajas.Show 1
                 MDIPrimero.MousePointer = 0
        
        Case 1707:
                 MDIPrimero.MousePointer = 11
                 frmFecha.Show
                 MDIPrimero.MousePointer = 0
        Case 1708:
                 MDIPrimero.MousePointer = 11
                 FrmPeriodoFiscal.Show
                 MDIPrimero.MousePointer = 0
        Case 1709:
                 MDIPrimero.MousePointer = 11
                 frmPuntosSolicitudes.Show
                 MDIPrimero.MousePointer = 0
        Case 1710
                 MDIPrimero.MousePointer = 11
                 frmActividades.Show
                 MDIPrimero.MousePointer = 0

        Case 1711:
                   MDIPrimero.MousePointer = 11
                   FrmListNomina.Show 1
                   MDIPrimero.MousePointer = 0
                   
        Case 1712:
                   MDIPrimero.MousePointer = 11
                   FrmHistorial.Show
                   MDIPrimero.MousePointer = 0
        Case 1713:
                    MDIPrimero.MousePointer = 11
                    FrmListadoNominaVacaciones.Show
                    MDIPrimero.MousePointer = 0
        Case 1714:
                    MDIPrimero.MousePointer = 11
                    frmTasa2.Show 1
                    MDIPrimero.MousePointer = 0
        Case 1743:
                    MDIPrimero.MousePointer = 11
                    FrmInforme.Show
                    MDIPrimero.MousePointer = 0
        Case 1716:
                    MDIPrimero.MousePointer = 11
                    frmClaves.Show
                    MDIPrimero.MousePointer = 0
        Case 1717:
                   MDIPrimero.MousePointer = 11
                   FrmRespaldar.Show 1
                   MDIPrimero.MousePointer = 0
        Case 1718
                    MDIPrimero.MousePointer = 11
                    FrmNomSubsidio.Show
                    MDIPrimero.MousePointer = 0
        Case 1719
                    MDIPrimero.MousePointer = 11
                    frmPuntosAutorizar.Show
                    MDIPrimero.MousePointer = 0

        Case 1720
                MDIPrimero.MousePointer = 11
                frmPlaneacion.Show
                MDIPrimero.MousePointer = 0
        Case 1721
                MDIPrimero.MousePointer = 11
                frmProduccionReal.Show
                MDIPrimero.MousePointer = 0
        Case 1722
                MDIPrimero.MousePointer = 11
                frmHorasExtras.Show
                MDIPrimero.MousePointer = 0
        Case 1723
                MDIPrimero.MousePointer = 11
                FrmTipoNomina.Show
                MDIPrimero.MousePointer = 0
        Case 1724
                MDIPrimero.MousePointer = 11
                frmPuntosGrupo.Show
                MDIPrimero.MousePointer = 0
        Case 1725
                MDIPrimero.MousePointer = 11
                frmPuntos.Show
                MDIPrimero.MousePointer = 0
        Case 1801
                MDIPrimero.MousePointer = 11
                frmFinca.Show
                MDIPrimero.MousePointer = 0
        Case 1802
                MDIPrimero.MousePointer = 11
                frmPlantacion.Show
                MDIPrimero.MousePointer = 0
        Case 1803
                MDIPrimero.MousePointer = 11
                frmFincaPlantaciones.Show
                MDIPrimero.MousePointer = 0
                
        Case 1726
                MDIPrimero.MousePointer = 11
                FrmDepartamentos.Show
                MDIPrimero.MousePointer = 0
        Case 1727
                MDIPrimero.MousePointer = 11
                FrmCargo.Show
                MDIPrimero.MousePointer = 0
        Case 1728
                MDIPrimero.MousePointer = 11
                FrmTipoIncapacidad.Show
                MDIPrimero.MousePointer = 0
        Case 1729
                MDIPrimero.MousePointer = 11
                FrmIncapacidades.Show
                MDIPrimero.MousePointer = 0
        Case 1730
                MDIPrimero.MousePointer = 11
                FrmIncentivo.Show
                MDIPrimero.MousePointer = 0
        Case 1731
                MDIPrimero.MousePointer = 11
                FrmDeduccion.Show
                MDIPrimero.MousePointer = 0
        Case 1732
            MDIPrimero.MousePointer = 11
            FrmSubsidio.Show
            MDIPrimero.MousePointer = 0
        Case 1733
            MDIPrimero.MousePointer = 11
            FrmTipoComision.Show
            MDIPrimero.MousePointer = 0
        Case 1734
            MDIPrimero.MousePointer = 11
            FrmTipoDestajo.Show
            MDIPrimero.MousePointer = 0
        Case 1735
            MDIPrimero.MousePointer = 11
            FrmGrupo.Show
            MDIPrimero.MousePointer = 0
        Case 1736
            MDIPrimero.MousePointer = 11
            FrmReferencias.Show
            MDIPrimero.MousePointer = 0
        Case 1737
             MDIPrimero.MousePointer = 11
            FrmProcesos.Show
            MDIPrimero.MousePointer = 0
        Case 1738
            MDIPrimero.MousePointer = 11
            FrmPermiso.Show
            MDIPrimero.MousePointer = 0
        Case 1739
            MDIPrimero.MousePointer = 11
           FrmIncentivoMetas.Show
            MDIPrimero.MousePointer = 0
        Case 1740
            MDIPrimero.MousePointer = 11
            FrmProduccionManual.Show
            MDIPrimero.MousePointer = 0
        Case 1741
            MDIPrimero.MousePointer = 11
            FrmSuspencion.Show
            MDIPrimero.MousePointer = 0
        Case 1742
            MDIPrimero.MousePointer = 11
            Directorio = App.Path + "\Calc.exe"
            Directorio = Shell(Directorio)
            MDIPrimero.MousePointer = 0
        Case 1744
            MDIPrimero.MousePointer = 11
            FrmControles.Show
            MDIPrimero.MousePointer = 0
        Case 1745
                MDIPrimero.MousePointer = 11
                FrmInssIR.Show
                MDIPrimero.MousePointer = 0
        Case 1746
            MDIPrimero.MousePointer = 11
            FrmReportes.CmbReportes.AddItem "Listado de Cargos"
            FrmReportes.CmbReportes.AddItem "Listado de Departamentos"
            FrmReportes.CmbReportes.AddItem "Listado de Tipos de Subsidios"
            FrmReportes.CmbReportes.AddItem "Listado de Tipos de Incentivos"
            FrmReportes.CmbReportes.AddItem "Listado de Tipos de Deducciones"
            FrmReportes.CmbReportes.AddItem "Reporte Proyeccion Vacaciones"
            FrmReportes.Show 1
            MDIPrimero.MousePointer = 0
            
         Case 1747
            FrmReportes.CmbReportes.AddItem "Numeros Disponibles"
            FrmReportes.CmbReportes.AddItem "Reporte x Produccion"
            FrmReportes.CmbReportes.AddItem "Reporte x Produccion Basico"
            FrmReportes.CmbReportes.AddItem "Reporte x Produccion Linea"
            FrmReportes.CmbReportes.AddItem "Analisis Produccion"
            FrmReportes.CmbReportes.AddItem "Analisis Produccion Resumen"
            FrmReportes.CmbReportes.AddItem "Lista de Empleados Activos"
            FrmReportes.CmbReportes.AddItem "Listado Maestro de Empleados"
            FrmReportes.CmbReportes.AddItem "Salario Basico Vrs Produccion"
            FrmReportes.CmbReportes.AddItem "Adelantos 13vo y Vacaciones"
            FrmReportes.CmbReportes.AddItem "Resumen-Pago Mensual"
            FrmReportes.CmbReportes.AddItem "Reporte x Provision"
            FrmReportes.CmbReportes.AddItem "Total-Pago Mensual"
            FrmReportes.CmbReportes.AddItem "Detalle Deducciones"
            FrmReportes.CmbReportes.AddItem "Reporte Dias Acumulados"
            FrmReportes.CmbReportes.AddItem "Reporte Horas Extra"
            FrmReportes.CmbReportes.AddItem "Reporte Estimado Vacaciones"
            FrmReportes.CmbReportes.AddItem "Reporte Registro Vacaciones"
            FrmReportes.CmbReportes.AddItem "Reporte Consolidado Vacaciones"
            FrmReportes.CmbReportes.AddItem "Reporte Total Vacaciones"
            FrmReportes.CmbReportes.AddItem "Listado de Empleados FHM"
            FrmReportes.CmbReportes.AddItem "Reporte Carnet Empleados"
            FrmReportes.Show 1
            
         Case 1748
            FrmReportes.CmbReportes.AddItem "Reporte Inss"
            FrmReportes.CmbReportes.AddItem "Reporte Detalle Inss"
            FrmReportes.CmbReportes.AddItem "Reporte Inss 2"
            FrmReportes.CmbReportes.AddItem "EXPORTACION INSS"
            FrmReportes.CmbReportes.AddItem "Reporte Ir"
            FrmReportes.CmbReportes.AddItem "Reporte Detalle Ir"
            FrmReportes.CmbReportes.AddItem "Reporte IR MENSUAL"
            FrmReportes.CmbReportes.AddItem "Reporte GRAL INGRESOS"
            FrmReportes.CmbReportes.AddItem "Reporte INSS E IR MENSUAL"
            FrmReportes.CmbReportes.AddItem "Reporte Detalle Deducciones"
            FrmReportes.Show 1
            
         Case 1749
                MDIPrimero.MousePointer = 11
                FrmNominaAcumulada.Show
                MDIPrimero.MousePointer = 0
                
         Case 1750
                MDIPrimero.MousePointer = 11
                FrmReembolso.Show
                MDIPrimero.MousePointer = 0
         Case 1751
                MDIPrimero.MousePointer = 11
                FrmAdelantos13vo.Show
                MDIPrimero.MousePointer = 0
         Case 1752
                MDIPrimero.MousePointer = 11
                frmFinca.Show
                MDIPrimero.MousePointer = 0
         Case 1753
                MDIPrimero.MousePointer = 11
                frmFincaPlantaciones.Show
                MDIPrimero.MousePointer = 0
         Case 1754
                MDIPrimero.MousePointer = 11
                frmPlantacion.Show
                MDIPrimero.MousePointer = 0
         Case 1755
                MDIPrimero.MousePointer = 11
                frmAsistManual.Show
                MDIPrimero.MousePointer = 0
         Case 1756
                MDIPrimero.MousePointer = 11
                frmPermisos.Show
                MDIPrimero.MousePointer = 0
         Case 1757
                MDIPrimero.MousePointer = 11
                frmRepAsistencia.Show
                MDIPrimero.MousePointer = 0
         Case 1758
                MDIPrimero.MousePointer = 11
                 FrmAsistencias.Show
                MDIPrimero.MousePointer = 0
         Case 1759
                MDIPrimero.MousePointer = 11
                FrmCopiaMarcas.Show
                MDIPrimero.MousePointer = 0
         Case 1760
                MDIPrimero.MousePointer = 11
                FrmImportar.Show
                MDIPrimero.MousePointer = 0
         Case 1761
                MDIPrimero.MousePointer = 11
                FrmControlVacaciones.Show
                MDIPrimero.MousePointer = 0
         Case 1762
                MDIPrimero.MousePointer = 11
                FrmSolicitud.Show
                MDIPrimero.MousePointer = 0
         Case 1763
                MDIPrimero.MousePointer = 11
                FrmListadoEmpleado.Show
                MDIPrimero.MousePointer = 0
                
         Case 1779
                MDIPrimero.MousePointer = 11
                FrmRegistroJustifica.Show
                MDIPrimero.MousePointer = 0
                
         Case 1780
                MDIPrimero.MousePointer = 11
                FrmJustificacion.Show
                MDIPrimero.MousePointer = 0
                
        Case 1781
            FrmReportes.CmbReportes.AddItem "Exportar Lista Empleados"
            FrmReportes.Show 1
            
        Case 1782
                MDIPrimero.MousePointer = 11
                FrmListaBajas.Show
                MDIPrimero.MousePointer = 0
                
       Case 1783
                MDIPrimero.MousePointer = 11
                FrmDeduccionPorciento.Show
                MDIPrimero.MousePointer = 0
                
       Case 1784
                MDIPrimero.MousePointer = 11
                FrmListSolicitudes.Show
                MDIPrimero.MousePointer = 0
 End Select
        
  Exit Sub
TipoErrs:
  MDIPrimero.MousePointer = 0
  MsgBox Err.Description
        
End Sub

Public Function RibbonBar() As RibbonBar
    Set RibbonBar = CommandBars.ActiveMenuBar
End Function

Private Sub CreateRibbonBar()

    Dim TabView As RibbonTab
    Dim TabHome As RibbonTab
    Dim TabCatalogo As RibbonTab
    Dim TabEdit As RibbonTab
    Dim TabPrintPreview As RibbonTab
    Dim GroupFile As RibbonGroup
    Dim GroupClipboard As RibbonGroup
    Dim GroupEditing As RibbonGroup
    Dim GroupShowHide As RibbonGroup
    Dim GroupDocumentViews As RibbonGroup
    Dim GroupWindow As RibbonGroup
    Dim GroupPrint As RibbonGroup
    Dim GroupPageSetup As RibbonGroup
    Dim GroupZoom As RibbonGroup
    Dim GroupPreview As RibbonGroup
    Dim ControlCuentas As CommandBarButton
    Dim ControlPrint As CommandBarPopup
    Dim Control As CommandBarControl
    Dim ControlPaste As CommandBarPopup
    Dim ControlSelect As CommandBarPopup
    Dim ControlPopup As CommandBarPopup
    Dim ControlMargins As CommandBarPopup
    Dim ControlOrientation As CommandBarPopup
    Dim ControlSize As CommandBarPopup
    Dim ControlFile As CommandBarPopup
    Dim ControlAbout As CommandBarControl
    Dim item As CommandBarControl





    Dim RibbonBar As RibbonBar
    CommandBars.Options.UseSharedImageList = False
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Icono.png", 1200, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Salir.png", 1300, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Empleado.png", 1700, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ActivarNomina.png", 1701, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Produccion.png", 1702, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\HorasExtra.png", 1703, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Calcular.png", 1704, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Vacaciones.png", 1705, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Despido.png", 1706, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\PeriodoNomina.png", 1707, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\PeriodoFiscal.png", 1708, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Complemento.png", 1709, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\AdministradorTareas.png", 1710, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ListadoNominas.png", 1711, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\HistorialSalarial.png", 1712, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ListadoVacaciones.png", 1713, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\TasaCambio.png", 1714, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Configuracion.png", 1715, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Usuarios.png", 1716, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Respaldar.png", 1717, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Subsidio.png", 1718, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\SolicitudPuntos.png", 1719, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\PlanActividades.png", 1720, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\AdministradorHoras.png", 1721, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\AprobarHorasExtra.png", 1722, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\TipoNominas.png", 1723, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\GrupoPuntos.png", 1724, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Puntos.png", 1725, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Departamento.png", 1726, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Cargo.png", 1727, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\TipoIncacidad.png", 1728, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Incapacidad.png", 1729, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Incentivos.png", 1730, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Deducciones.png", 1731, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\TipoSubsidio.png", 1732, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\TipoComision.png", 1733, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Destajo.png", 1734, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\DivicionNomina.png", 1735, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Referencias.png", 1736, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Procesos.png", 1737, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Permisos.png", 1738, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\IncentivoMetas.png", 1739, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ProduccionManual.png", 1740, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Suspencion.png", 1741, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Calculator.png", 1742, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\InformacionUsuario.png", 1743, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Controles.png", 1744, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\INSS.png", 1745, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ReportesGenerales.png", 1746, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ReporteEmpleados.png", 1747, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ReporteDeducciones.png", 1748, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\NominaAcumulada.png", 1749, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Money.png", 1750, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Adelantos13vo.png", 1751, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\finca.png", 1752, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\plantacion.ico", 1753, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\finca.png", 1754, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\AsistenciaManual.png", 1755, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Money1.png", 1756, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\CalcularReporte.png", 1757, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\PersonalTurno.png", 1758, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Copy.png", 1759, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Excel2.png", 1760, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Calendario.png", 1761, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Solicitud.png", 1762, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Lista.png", 1763, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Icono.png", 1764, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Salir.png", 1765, XtremeCommandBars.XTPImageState.xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ReportesGenerales.png", 1766, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\HorasExtra.png", 1767, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Estadisticos.png", 1768, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Controles.png", 1769, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Departamentos.png", 1770, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Exportar.png", 1771, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Tarjeta.png", 1772, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Jornada.png", 1773, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\AsignacionJornada.png", 1774, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Horarios.png", 1775, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Nominas.png", 1776, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Excel2.png", 1777, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ExportarExcel5.png", 1778, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\home_page.png", 1779, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Justificacion.png", 1780, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ExportarExcel.png", 1781, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ListaLiquidacion.png", 1782, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Porciento.png", 1783, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ProduccionManual.png", 1784, xtpImageNormal

    
'
'    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'    '///////////////////////////////////CREO EL RIBBON Y LE CARGO LA IMAGEN//////////////////////////////////////////////////
'    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Set RibbonBar = CommandBars.AddRibbonBar("The Ribbon")
    RibbonBar.EnableDocking xtpFlagStretched

    Set ControlFile = RibbonBar.AddSystemButton()
    ControlFile.IconId = 1200
           Set Control = ControlFile.CommandBar.Controls.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1300, "S&alir", False, False)
    Control.BeginGroup = True
    ControlFile.CommandBar.SetIconSize 35, 35
'
'    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'    '///////////////////////////////////CREO LOS TABS//////////////////////////////////////////////////
'    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Set TabHome = RibbonBar.InsertTab(0, "&Accesos")
    TabHome.Id = 130
        Set GroupFile = TabHome.Groups.AddGroup("Procesos", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1700, "&Empleados", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1763, "&Lista Empleados", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1701, "&Activar Nominas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1702, "&Movimiento de Produccion", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1703, "&Horas Extras", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1704, "&Calcular Nominas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1705, "&Vacaciones  13vo Mes", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1706, "&Despidos", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
      Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1782, "&Lista Liquidacion", False, False)
     item.Style = xtpButtonIconAndCaptionBelow

      Set GroupFile = TabHome.Groups.AddGroup("Catalogo", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1707, "&Periodo Nominas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1708, "&Periodo Fiscal", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1709, "&Complementos Salariales", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1710, "&Administrador de Tareas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow

     Set GroupFile = TabHome.Groups.AddGroup("Historicos", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1711, "&Listado de Nominas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1712, "&Historial Salarial", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1713, "&List Nominas Vac-13vo", False, False)
     item.Style = xtpButtonIconAndCaptionBelow


     Set GroupFile = TabHome.Groups.AddGroup("Opciones", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1714, "&Tasa de Cambio", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1743, "&Informacion de Usuarios", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1716, "&Usuarios", False, False)
     item.Style = xtpButtonIconAndCaptionBelow

     Set GroupFile = TabHome.Groups.AddGroup("Ayuda", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1717, "&Respaldar", False, False)
     item.Style = xtpButtonIconAndCaptionBelow

    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////CREO EL TABS DE PROCESOS//////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////
    Set TabHome = RibbonBar.InsertTab(1, "&Procesos")
    TabHome.Id = 1500
     Set GroupFile = TabHome.Groups.AddGroup("Catalogos", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1700, "&Empleados", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1701, "&Activar Nominas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1702, "&Movimiento de Produccion", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1703, "&Horas Extras", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1704, "&Calular Nomina", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1718, "&Subsidios", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1709, "Co&mplementos Salariales", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1719, "&Autorizacion de Puntos", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1720, "&Plan de Actividades", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1721, "&Adm de Horas Laborales", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1722, "&Aprobar Horas Extras", False, False)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1750, "&Reembolso Vacaciones", False, False)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1751, "&Adelanto 13vo y Vacaciones", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
'    '/////////////////////////////////////////////////////////////////////////////////////////////////////
'    '///////////////////////////////CREO EL TABS DE CATALOGO//////////////////////////////////////////////
'    '///////////////////////////////////////////////////////////////////////////////////////////////////////
    Set TabHome = RibbonBar.InsertTab(2, "&Catalogo")
    TabHome.Id = 1500
     Set GroupFile = TabHome.Groups.AddGroup("Basicos", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1707, "&Periodo Nomina", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1708, "&Periodo Fiscal", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1723, "&Tipo Nomina", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1724, "&Grupo de Puntos", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1725, "P&untos", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1752, "Fincas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1753, "Plantaciones", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1754, "Fincas - Plantaciones", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1710, "&Administrador de Tareas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1745, "Tabla INSS/IR", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set GroupFile = TabHome.Groups.AddGroup("Generales", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1726, "&Departamentos", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1727, "&Cargos", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1728, "&Tipo Incacidad", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
      Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1729, "&Incapacidad", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1730, "&Tipo Incentivo", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1731, "&Tipo Deducciones", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1732, "&Tipo Subsidio", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1733, "&Tipo Comision", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1734, "&Tipo Destajo", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1735, "&Division de Nominas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1783, "&Deduccion Porciento", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////CREO EL TABS DE RECURSOS HUMANOS//////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////
    Set TabHome = RibbonBar.InsertTab(4, "&Recursos Humanos")
    TabHome.Id = 1500
     Set GroupFile = TabHome.Groups.AddGroup("Vacaciones", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1761, "Calendario", False, False)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1762, "Solicitud", False, False)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1779, "Registro Justifica", False, False)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1780, "Justificaciones", False, False)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1784, "Grupo Solicitud", False, False)

    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////CREO EL TABS DE PRODUCCION//////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////
    Set TabHome = RibbonBar.InsertTab(3, "&Produccion")
    TabHome.Id = 1500
     Set GroupFile = TabHome.Groups.AddGroup("Basicos", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1736, "Referencias", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1737, "Procesos", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1738, "Permisos", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1739, "Incentivos x Metas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set GroupFile = TabHome.Groups.AddGroup("Produccion", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1702, "Movimientos de Produccion", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1740, "Produccion Manual", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////CREO EL TABS DE HISTORICOS//////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////
    Set TabHome = RibbonBar.InsertTab(4, "&Historicos")
    TabHome.Id = 1500
     Set GroupFile = TabHome.Groups.AddGroup("Basicos", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1711, "Listado de Nominas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1741, "Suspenciones", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1712, "Historial Salarial", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1713, "List Nomina Vac-13vo", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1749, "Nomina Acumulada", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     RibbonBar.QuickAccessControls.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_FILE_SAVE, "Zeus Nominas V.6.28", False, False


    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////CREO EL TABS DE Asistencia//////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////
     Set TabHome = RibbonBar.InsertTab(5, "&Asistencia")
     TabHome.Id = 1500
     Set GroupFile = TabHome.Groups.AddGroup("Control de Asistencia", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1755, "&Asistencia Manual", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1756, "&Ingreso/Mod Permisos", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1757, "&Calculo Laboradas, Reportes", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1758, "&Registros Asistencias", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1759, "&Copiar Asistencia Entre Compañia", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     TabHome.Id = 1500
     
     Set GroupFile = TabHome.Groups.AddGroup("Reloj Biometricos", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1766, "&Reportes Generales", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1767, "&Reportes Asistencia", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1768, "&Reportes Estadisticos", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set GroupFile = TabHome.Groups.AddGroup("Tablas", 2)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1770, "&Departamentos", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1772, "&Tarjeta de Marcadas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1773, "&Jornadas Laborales", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1774, "&Asignacion de Jornadas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1775, "&Configuracion Horarios", False, False)
     Set GroupFile = TabHome.Groups.AddGroup("Opciones", 3)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1769, "Controles Personalizados", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1771, "Exportar Archivos TXT", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1778, "Exportar Archivos XLS", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1776, "Movimientos Nominas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1777, "Importar Archivo", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     
     
     
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////CREO EL TABS DE OPCIONES//////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////
    Set TabHome = RibbonBar.InsertTab(6, "&Opciones")
    TabHome.Id = 1500
     Set GroupFile = TabHome.Groups.AddGroup("Basicos", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1716, "Usuarios", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1714, "Tasa de Cambio", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1743, "Informacion de Usurios", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1742, "Calculadora", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1744, "Controles Personalizados", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1760, "Importar", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     
     
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////CREO EL TABS DE REPORTES//////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////
    Set TabHome = RibbonBar.InsertTab(7, "&Reportes")
    TabHome.Id = 1500
     Set GroupFile = TabHome.Groups.AddGroup("Basicos", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1746, "Reportes Generales", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1747, "Reportes de Empleados", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1748, "Reportes de Deducciones", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1781, "Exportar Excel", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
End Sub

Private Sub wndTaskPanel_GroupExpanding(ByVal Group As XtremeTaskPanel.ITaskPanelGroup, ByVal Expanding As Boolean, Cancel As Boolean)
 If Expanding = True Then
  Select Case Group.Caption
    Case "Procesos"
              wndTaskPanel.Groups(2).Expanded = False
              wndTaskPanel.Groups(3).Expanded = False
              wndTaskPanel.Groups(4).Expanded = False
              wndTaskPanel.Groups(5).Expanded = False
              wndTaskPanel.Groups(6).Expanded = False

    Case "Catalogo"
              wndTaskPanel.Groups(1).Expanded = False
              wndTaskPanel.Groups(3).Expanded = False
              wndTaskPanel.Groups(4).Expanded = False
              wndTaskPanel.Groups(5).Expanded = False
              wndTaskPanel.Groups(6).Expanded = False

    Case "Produccion"
              wndTaskPanel.Groups(1).Expanded = False
              wndTaskPanel.Groups(2).Expanded = False
              wndTaskPanel.Groups(4).Expanded = False
              wndTaskPanel.Groups(5).Expanded = False
              wndTaskPanel.Groups(6).Expanded = False

              
    Case "Historicos"
              wndTaskPanel.Groups(1).Expanded = False
              wndTaskPanel.Groups(2).Expanded = False
              wndTaskPanel.Groups(3).Expanded = False
              wndTaskPanel.Groups(5).Expanded = False
              wndTaskPanel.Groups(6).Expanded = False

    
              
              
    Case "Opciones"
              wndTaskPanel.Groups(1).Expanded = False
              wndTaskPanel.Groups(2).Expanded = False
              wndTaskPanel.Groups(3).Expanded = False
              wndTaskPanel.Groups(4).Expanded = False
              wndTaskPanel.Groups(6).Expanded = False

              
    
    Case "Reportes"
              wndTaskPanel.Groups(1).Expanded = False
              wndTaskPanel.Groups(2).Expanded = False
              wndTaskPanel.Groups(3).Expanded = False
              wndTaskPanel.Groups(4).Expanded = False
              wndTaskPanel.Groups(5).Expanded = False

              
    
 
    
  
  End Select
 End If

End Sub


Sub CreateTaskPanel()


    Dim Group As TaskPanelGroup
    Dim item As TaskPanelGroupItem
    
    Set Group = wndTaskPanel.Groups.Add(100, "Procesos") '////GRUPO1///
    Group.Tooltip = "Sistema de Nominas"
    Group.Special = True
    Group.Items.Add 1, "Empleados", xtpTaskItemTypeLink, 1
    Group.Items.Add 2, "Activar Nominas", xtpTaskItemTypeLink, 2
    Group.Items.Add 3, "Movimiento de Produccion", xtpTaskItemTypeLink, 3
    Group.Items.Add 4, "Horas Extras", xtpTaskItemTypeLink, 4
    Group.Items.Add 5, "Calcular Nomina", xtpTaskItemTypeLink, 5
    Group.Items.Add 6, "Subsidios", xtpTaskItemTypeLink, 6
    
    'UPDATE: ING. ELIAZAR POLANCO
    Group.Items.Add 7, "Complementos Salariales", xtpTaskItemTypeLink, 9
    Group.Items.Add 8, "Solicitudes de Puntos", xtpTaskItemTypeLink, 8
    Group.Items.Add 9, "Planificación de Actividades", xtpTaskItemTypeLink, 7
    Group.Items.Add 10, "Administrador de Horas Laborales", xtpTaskItemTypeLink, 10
    Group.Items.Add 11, "Aprobar Horas Extras", xtpTaskItemTypeLink, 11
    
    
    Set Group = wndTaskPanel.Groups.Add(100, "Catalogo") '///GRUPO 2 //////
    Group.Tooltip = "Sistema de Nominas"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 7, "Periodo Fiscal", xtpTaskItemTypeLink, 7
    Group.Items.Add 8, "Periodo Nomina", xtpTaskItemTypeLink, 8
    Group.Items.Add 9, "Departamento", xtpTaskItemTypeLink, 10
    Group.Items.Add 10, "Cargos", xtpTaskItemTypeLink, 11
    Group.Items.Add 12, "Tipo Incapacidad", xtpTaskItemTypeLink, 12
    Group.Items.Add 13, "Incapacidades", xtpTaskItemTypeLink, 13
    Group.Items.Add 14, "Tipo Incentivo", xtpTaskItemTypeLink, 14
    Group.Items.Add 15, "Tipo Deducciones", xtpTaskItemTypeLink, 15
    Group.Items.Add 16, "Tipo Subsidio", xtpTaskItemTypeLink, 16
    Group.Items.Add 17, "Tipo Comision", xtpTaskItemTypeLink, 17
    Group.Items.Add 18, "Tipo Destajo", xtpTaskItemTypeLink, 18
    Group.Items.Add 19, "Division Nomina", xtpTaskItemTypeLink, 19
    Group.Items.Add 20, "Tipo Nomina", xtpTaskItemTypeLink, 20
    
    'UPDATE: ING. ELIAZAR POLANCO
    Group.Items.Add 21, "Grupo de Puntos", xtpTaskItemTypeLink, 21
    Group.Items.Add 22, "Puntos", xtpTaskItemTypeLink, 22
    Group.Items.Add 23, "Administrador de Actividades", xtpTaskItemTypeLink, 23

    Set Group = wndTaskPanel.Groups.Add(100, "Produccion") '////GRUPO 3 /////
    Group.Tooltip = "Sistema de Nominas"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 9, "Referencias", xtpTaskItemTypeLink, 22
    Group.Items.Add 10, "Procesos", xtpTaskItemTypeLink, 23
    Group.Items.Add 12, "Movimientos de Produccion", xtpTaskItemTypeLink, 3
    Group.Items.Add 13, "Permisos", xtpTaskItemTypeLink, 24
    Group.Items.Add 14, "Incentivo x Metas", xtpTaskItemTypeLink, 25
    Group.Items.Add 14, "Produccion Manual", xtpTaskItemTypeLink, 25
    
    
    
    
    Set Group = wndTaskPanel.Groups.Add(100, "Historicos") '////GRUPO 4 //////
    Group.Tooltip = "Sistema de Nominas"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 9, "Listado de Nominas", xtpTaskItemTypeLink, 9
    Group.Items.Add 6, "Suspenciones", xtpTaskItemTypeLink, 26
    Group.Items.Add 7, "Historial Salarial", xtpTaskItemTypeLink, 27
    Group.Items.Add 7, "Listado Nominas de Vacaciones/13vo", xtpTaskItemTypeLink, 36
    
    


    
    Set Group = wndTaskPanel.Groups.Add(100, "Opciones")  '/////GRUPO 5 /////
    Group.Tooltip = "Procesos del Sistema Contable"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 13, "Usuarios", xtpTaskItemTypeLink, 28
    Group.Items.Add 13, "Tasa de Cambio", xtpTaskItemTypeLink, 29
    Group.Items.Add 13, "Informacion de Usuarios", xtpTaskItemTypeLink, 30
    Group.Items.Add 13, "Calculadora", xtpTaskItemTypeLink, 31
    Group.Items.Add 13, "Controles Personalizados", xtpTaskItemTypeLink, 32
    
    Set Group = wndTaskPanel.Groups.Add(100, "Reportes") '/////GRUPO 6//////
    Group.Tooltip = "Procesos del Sistema Contable"
    Group.Special = True
    Group.Expanded = False
    Group.Items.Add 13, "Reportes Generales", xtpTaskItemTypeLink, 33
    Group.Items.Add 13, "Reportes de Empleados", xtpTaskItemTypeLink, 34
    Group.Items.Add 13, "Reportes de Deducciones", xtpTaskItemTypeLink, 35
    
   
    
     
    wndTaskPanel.SetImageList Me.ImageList1
End Sub







Private Sub abrir_Click()
On Error GoTo TipoErrs
MDIPrimero.MousePointer = 11
frmBackup.Height = 5070
frmBackup.Width = 5160
frmBackup.Show
MDIPrimero.MousePointer = 0
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub Claves_Click()
On Error GoTo TipoErrs
frmClaves.Height = 3060
frmClaves.Width = 4230
frmClaves.Show
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub AgregaClaves_Click()
On Error GoTo TipoErrs
frmClaves.Show
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub Acerca_Click()
FrmAcerca.Show
End Sub

Private Sub Anotaciones_Click()
On Error GoTo TipoErrs
MDIPrimero.MousePointer = 11
FrmAnotaciones.Show
MDIPrimero.MousePointer = 0
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub ayuda_Click()
On Error GoTo TipoErrs

 
Exit Sub
TipoErrs:
 ControlErrores
 End Sub

Private Sub Barra_Click()
 ' Llama al procedimiento de la barra de herramientas, pasando una referencia
    ' a esta instancia de formulario.
If ToolBar.Visible = True Then
   Barra.Checked = False
   ToolBar.Visible = False
Else
  ToolBar.Visible = True
  Barra.Checked = True
End If
End Sub

Private Sub Calculadora_Click()
Dim ruta
On Error GoTo TipoErrs
  MDIPrimero.MousePointer = 11
  'Determina la Ruta raiz del programa.
  'Ruta = App.Path
  ruta = ruta & "C:\WINDOWS\Calc.exe"
  ruta = Shell(ruta)
 MDIPrimero.MousePointer = 0
 Exit Sub
TipoErrs:
 Prueba = 1
 ControlErrores
End Sub

Private Sub CambiaClave_Click()
 On Error GoTo TipoErrs
MDIPrimero.MousePointer = 11
 frmTasa2.Show
MDIPrimero.MousePointer = 0
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub Cargo_Click()
On Error GoTo TipoErrs
MDIPrimero.MousePointer = 11
FrmCargo.Show
MDIPrimero.MousePointer = 0
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub Cmd13vo_Click()
Frm13Vaca.Show 1
End Sub

Private Sub CmdActivar_Click()
    MDIPrimero.MousePointer = 11
      FrmActivarNomina.Show
         MDIPrimero.MousePointer = 0
End Sub

Private Sub CmdAdelanto_Click()
FrmAdelantos13vo.Show 1
End Sub

Private Sub CmdCalcular_Click()
On Error GoTo TipoErrs
FrmCalcularNomina.Show 1
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdDespido_Click()
FrmBajas.Show 1
End Sub

Private Sub CmdEmpleado_Click()

          MDIPrimero.MousePointer = 11
          frmEmpleado.Show
          MDIPrimero.MousePointer = 0
 
End Sub

Private Sub CmdInss_Click()
MDIPrimero.MousePointer = 11
         FrmInssIR.Show
         MDIPrimero.MousePointer = 0
End Sub

Private Sub CmdMovimiento_Click()
FrmMovimientos.Show
End Sub

Private Sub CmdPeriodos_Click()
frmFecha.Show 1
End Sub

Private Sub CmdProduccion_Click()
FrmProduccion.Show 1
End Sub

Private Sub CmdRespaldar_Click()
FrmRespaldar.Show 1
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSubsidio_Click()
FrmNomSubsidio.Show
End Sub

Private Sub CmdUsuario_Click()
MDIPrimero.MousePointer = 11
         frmClaves.Show
         MDIPrimero.MousePointer = 0
End Sub

Private Sub Contendido_Click()
CommonDialog1.CancelError = True
CommonDialog1.HelpCommand = &H9&
CommonDialog1.HelpFile = App.Path + "\Zeus.hlp"
CommonDialog1.ShowHelp
End Sub

Private Sub departamento_Click()
On Error GoTo TipoErrs
MDIPrimero.MousePointer = 11
FrmDepartamentos.Width = 7530
FrmDepartamentos.Show
MDIPrimero.MousePointer = 0
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub empleados_Click()
On erro GoTo TipoErrs
MDIPrimero.MousePointer = 11
frmEmpleado.Show
MDIPrimero.MousePointer = 0
Exit Sub
TipoErrs:
ControlErrores
End Sub


Private Sub Exportar_Click()
 FrmExporta.Show
End Sub

Private Sub imgCopyButton_Click()
  On Error GoTo TipoErrs
    ' Actualiza la imagen.
    imgCopyButton.Refresh
    ' Llama al procedimiento de copiar
    frmEmpleado.Show
     imgCopyButton.Refresh
Exit Sub
TipoErrs:
 ControlErrores
End Sub
Private Sub imgCopyButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Muestra la imagen del estado presionado.
    imgCopyButton.Picture = imgCopyButtonDn.Picture
End Sub

Private Sub imgCopyButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si el botón está presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su área; si no
    ' presenta el mapa de bits del estado presionado.
    Select Case Button
    Case 1
        If X <= 0 Or X > imgCopyButton.Width Or Y < 0 Or Y > imgCopyButton.Height Then
            imgCopyButton.Picture = imgCopyButtonUp.Picture
        Else
            imgCopyButton.Picture = imgCopyButtonDn.Picture
        End If
    End Select
End Sub
Private Sub imgCopyButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Muestra la imagen del estado sin presionar.
    imgCopyButton.Picture = imgCopyButtonUp.Picture
End Sub

Private Sub ImgEmpleado_Click()
   On Error GoTo TipoErrs
   ImgEmpleado.Refresh
    ' abre el formulario empleado
    frmEmpleado.Show
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub ImgEmpleado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Muestra la imagen del estado presionado.
    ImgEmpleado.Picture = imgEmpleadoDn.Picture
End Sub
Private Sub ImgEmpleado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si el botón está presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su área; si no
    ' presenta el mapa de bits del estado presionado.
    Select Case Button
    Case 1
        If X <= 0 Or X > ImgEmpleado.Width Or Y < 0 Or Y > ImgEmpleado.Height Then
            ImgEmpleado.Picture = imgEmpleadoUp.Picture
        Else
            ImgEmpleado.Picture = imgEmpleadoDn.Picture
        End If
    End Select
End Sub
Private Sub ImgEmpleado_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Muestra la imagen del estado sin presionar.
    ImgEmpleado.Picture = imgEmpleadoUp.Picture
End Sub

Private Sub imgSalir_Click()
  ' Actualiza la imagen.
    imgSalir.Refresh
   
    Unload Me
End Sub
Private Sub imgSalir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Muestra la imagen del estado presionado.
    imgSalir.Picture = imgSalirDn.Picture
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo TipoErrs
    ' Si el botón está presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su área; si no
    ' presenta el mapa de bits del estado presionado.
    Select Case Button
    Case 1
        If X <= 0 Or X > imgSalir.Width Or Y < 0 Or Y > imgSalir.Height Then
            imgSalir.Picture = imgSalirUp.Picture
        Else
            imgSalir.Picture = imgSalirDn.Picture
        End If
    End Select
Exit Sub
TipoErrs:
 ControlErrores
End Sub
Private Sub imgSalir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Muestra la imagen del estado sin presionar.
    imgSalir.Picture = imgSalirUp.Picture
End Sub


Private Sub incapacidades_Click()
On Error GoTo TipoErrs
MDIPrimero.MousePointer = 11
FrmIncapacidades.Show
MDIPrimero.MousePointer = 0
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub Informa_Click()
FrmInforme.Show
End Sub

Private Sub Inss_Click()
 On Error GoTo TipoErrs
MDIPrimero.MousePointer = 11

 FrmInssIR.Show
MDIPrimero.MousePointer = 0
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub IR_Click()
 On Error GoTo TipoErrs
 FrmInssIR.Width = 6495
 FrmInssIR.Height = 3435
 FrmInssIR.Show
 Exit Sub
TipoErrs:
 ControlErrores
 End Sub

Private Sub MCascada_Click()
' Organiza los formularios hijos en cascada.
    MDIPrimero.Arrange vbCascade
End Sub

Private Sub MDIForm_Activate()
 
 
'    MDIPrimero.HelpContextID = (1)
'    If DateTime.Now > CDate("13/12/2017") Then
'        MsgBox ("Periodo de prueba caducado")
'        KillProcess ("SistemaNominas.exe")
'    End If

End Sub
Private Sub MDIForm_Load()
'On Error GoTo TipoErr

Dim SqlSuspenciones As String
Dim VerificaTasa As Boolean
Dim Entrar As Boolean
Dim FechaIni As Date, FechaInicio As Date, FechaFinal As Date
Dim FechaFin As Date
Dim Encontrado As Boolean

 res = Bitacora(Now, NombreUsuario, "Ingreso al Sistema", "Ingresando al sistema de nominas")


'Me.CmdRespaldar.BackColor = RGB(236, 233, 216)
'Me.CmdProduccion.BackColor = RGB(236, 233, 216)
'Me.CmdPeriodos.BackColor = RGB(236, 233, 216)
'Me.CmdActivar.BackColor = RGB(236, 233, 216)
'Me.CmdEmpleado.BackColor = RGB(236, 233, 216)
'Me.CmdInss.BackColor = RGB(236, 233, 216)
'Me.CmdSalir.BackColor = RGB(236, 233, 216)
'Me.Cmd13vo.BackColor = RGB(236, 233, 216)
'Me.CmdAdelanto.BackColor = RGB(236, 233, 216)
'Me.CmdCalcular.BackColor = RGB(236, 233, 216)
'Me.CmdDespido.BackColor = RGB(236, 233, 216)
'Me.CmdMovimiento.BackColor = RGB(236, 233, 216)
'Me.CmdSubsidio.BackColor = RGB(236, 233, 216)
'Me.CmdUsuario.BackColor = RGB(236, 233, 216)
'Me.SmartMenuXP1.BackColor = RGB(236, 233, 216)
'Me.Picture1.BackColor = RGB(236, 233, 216)

Me.Picture1.BackColor = RGB(173, 199, 236)

Me.CmdRespaldar.BackColor = RGB(173, 199, 236)
Me.CmdProduccion.BackColor = RGB(173, 199, 236)
Me.CmdPeriodos.BackColor = RGB(173, 199, 236)
Me.CmdActivar.BackColor = RGB(173, 199, 236)
Me.CmdEmpleado.BackColor = RGB(173, 199, 236)
Me.CmdInss.BackColor = RGB(173, 199, 236)
Me.CmdSalir.BackColor = RGB(173, 199, 236)
Me.Cmd13vo.BackColor = RGB(173, 199, 236)
Me.CmdAdelanto.BackColor = RGB(173, 199, 236)
Me.CmdCalcular.BackColor = RGB(173, 199, 236)
Me.CmdDespido.BackColor = RGB(173, 199, 236)
Me.CmdMovimiento.BackColor = RGB(173, 199, 236)
Me.CmdSubsidio.BackColor = RGB(173, 199, 236)
Me.CmdUsuario.BackColor = RGB(173, 199, 236)
'Me.SmartMenuXP1.BackColor = RGB(173, 199, 236)


With Me.DtaControles
   .ConnectionString = Conexion
   .RecordSource = "Controles"
   .Refresh
End With

With Me.DtaEmpleados
   .ConnectionString = Conexion
   .RecordSource = "Empleado"
   .Refresh
End With

With Me.DtaNacceso
   .ConnectionString = Conexion
End With

With Me.DtaSuspenciones
   .ConnectionString = Conexion
End With

With Me.DtaConsulta
   .ConnectionString = Conexion
End With

With Me.DtaConsulta2
   .ConnectionString = Conexion
End With

With Me.DtaConsulta
   .ConnectionString = Conexion
End With

With Me.AdoReportes
   .ConnectionString = Conexion
End With

With Me.DtaIR
   .ConnectionString = Conexion
End With

With Me.DtaTasa
   .ConnectionString = Conexion
   .RecordSource = "CambioMoneda"
   .Refresh
End With

With Me.DtaEmpresa
   .ConnectionString = Conexion
   .RecordSource = "DatosEmpresa"
   .Refresh
End With


With Me.AdoTasaContabilidad
  .ConnectionString = ConexionContable
End With



 ChDir App.Path
 
Entrar = False

DtaSuspenciones.RecordSource = "SELECT CodEmpleado, Fechaini, FechaFin, Ultimo, Activo From Subsidios WHERE (Activo = - 1) OR (Ultimo = - 1)"
DtaSuspenciones.Refresh
'////Falso es igual a cero

Do While Not DtaSuspenciones.Recordset.EOF
   ''DtaSuspenciones.Recordset.Edit
   DtaSuspenciones.Recordset("Ultimo") = False
   DtaSuspenciones.Recordset.Update
DtaSuspenciones.Recordset.MoveNext
Loop

DtaSuspenciones.Refresh
Do While Not DtaSuspenciones.Recordset.EOF
    FechaIni = DtaSuspenciones.Recordset("Fechafin")
    FechaFin = Format(Now, "dd/mm/yyyy")
    
    If FechaIni <= FechaFin Then
      'actualizo las suspenciones
       ''DtaSuspenciones.Recordset.Edit
'        DtaSuspenciones.Recordset("activo") = True
        DtaSuspenciones.Recordset("ultimo") = True
       DtaSuspenciones.Recordset.Update
       'actualizo los empleados
       DtaEmpleados.Refresh
       Do While Not DtaEmpleados.Recordset.EOF
           If DtaEmpleados.Recordset("CodEmpleado") = DtaSuspenciones.Recordset("CodEmpleado") Then
              ''DtaEmpleados.Recordset.Edit
              DtaEmpleados.Recordset("Ausente") = False
              DtaEmpleados.Recordset.Update
           End If
       DtaEmpleados.Recordset.MoveNext
       Loop
       
       Entrar = True
    End If
DtaSuspenciones.Recordset.MoveNext
Loop



If Entrar Then
   MsgBox "Algunos empleados fueron reincorporados"
   FrmListSubsidios.Show
End If
DtaControles.Refresh


'VerificaTasa = Me.DtaControles.Recordset("VerificarTasa")



'/////////Compruebo la Opcion de Pedir la Tasa////////////////


'With Me.DtaEmpresa
'   .ConnectionString = Conexion
'   .RecordSource = "DatosEmpresa"
'   .Refresh
'End With
DtaEmpresa.Refresh
Titulo = DtaEmpresa.Recordset("NombreEmpresa")
SubTitulo = DtaEmpresa.Recordset("Direccion") '+ " RUC: " + DtaEmpresa.Recordset("numeroruc")
'RutaLogo = DtaEmpresa.Recordset.RutaLogo
StatusBar2.Panels(2) = "Licencia: " + Titulo


Set item = PopupControl1.AddItem(20, 15, 270, 45, Titulo)
item.TextColor = RGB(0, 61, 178)
item.Bold = True
Set item = PopupControl1.AddItem(20, 29, 400, 100, "Direc:" & DtaEmpresa.Recordset("Direccion"))
item.TextColor = RGB(0, 61, 178)
item.Bold = True
Set item = PopupControl1.AddItem(60, 60, 400, 100, "Bienvenido: " & NombreUsuario)
    item.Bold = True
    PopupControl1.VisualTheme = xtpPopupThemeOffice2003
    PopupControl1.SetSize 300, 110
    Me.PopupControl1.Show
    Me.PopupControl1.Show

    






'CreateTaskPanel

CargarInterfaz

CreateRibbonBar

RibbonBar.EnableFrameTheme


If VerificaTasa = True Then
DtaTasa.Refresh
Encontrado = False
Me.DtaTasa.RecordSource = "SELECT CambioMoneda.* From CambioMoneda WHERE (FechaDia = CONVERT(DATETIME, '" & Format(Now, "yyyy-mm-dd") & "', 102))"
Me.DtaTasa.Refresh
If Me.DtaTasa.Recordset.EOF Then
        '///////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////SI NO EXISTE LA BUSCO EN CONTABILIDAD ///////////
        '////////////////////////////////////////////////////////////////////////////////////////
        FechaInicio = DateSerial(Year(Now), Month(Now), 1)
        FechaFinal = DateSerial(Year(Now), Month(Now) + 1, 0)
        Me.AdoTasaContabilidad.RecordSource = "SELECT  * From Tasas WHERE (FechaTasas BETWEEN CONVERT(DATETIME, '" & Format(FechaInicio, "yyyy-MM-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FechaFinal, "yyyy-MM-dd") & "', 102))"
        Me.AdoTasaContabilidad.Refresh
        Do While Not Me.AdoTasaContabilidad.Recordset.EOF
           If BuscaTasaCambio(Me.AdoTasaContabilidad.Recordset("FechaTasas")) = 1 Then
               Me.DtaTasa.Recordset.AddNew
                   Me.DtaTasa.Recordset("FechaDia") = Me.AdoTasaContabilidad.Recordset("FechaTasas")
                   Me.DtaTasa.Recordset("MontoDia") = Me.AdoTasaContabilidad.Recordset("MontoCordobas")
               Me.DtaTasa.Recordset.Update
           Else
                   Me.DtaTasa.Recordset("MontoDia") = Me.AdoTasaContabilidad.Recordset("MontoCordobas")
               Me.DtaTasa.Recordset.Update
           End If
           
           
           If Format(Me.AdoTasaContabilidad.Recordset("FechaTasas"), "dd/mm/yyyy") = Format(Now, "dd/mm/yyyy") Then
                 Encontrado = True
                Tasa = Me.AdoTasaContabilidad.Recordset("MontoCordobas")
                StatusBar2.Panels(3) = "Tasa: " & Format(Tasa, "##,##0.0000")
                Encontrado = True
          
           End If
           
           
          Me.AdoTasaContabilidad.Recordset.MoveNext
        Loop
Else
    Encontrado = True
    Tasa = DtaTasa.Recordset("montodia")
    StatusBar2.Panels(3) = "Tasa: " & Format(Tasa, "##,##0.0000")
    Encontrado = True

End If
 
Else
  Encontrado = True
End If


'------------------------------------------------------------------------------------------------
'----------------------------------CARGO LAS VARIABLES DEL RELOJ --------------------------------
'------------------------------------------------------------------------------------------------
Dim IDNumber As String, IpAdress As String, ri As Long
Dim RutaConexion As String, ConexionBD As String
Dim RutaBD As String



    '////////////////////////BUSCO EL DIRECTORIO Y RUTA DE LAS BASE DE DATOS ///////////////////////////////
    RutaConexion = App.Path + "\standard\CntReloj.dll"
    ConexionBD = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaConexion & ";Persist Security Info=False"
    With Me.AdoConexion
     .ConnectionString = ConexionBD
     .RecordSource = "SELECT Servidor.* FROM Servidor"
     .Refresh
    End With
    
    If Not Me.AdoConexion.Recordset.EOF Then
      RutaBD = Me.AdoConexion.Recordset("Servidor")
      If RutaBD = "APP" Then
         RutaBD = "APP"
      Else
        RutaBD = Me.AdoConexion.Recordset("Servidor")
      End If
    End If
    
    
    If RutaBD = "APP" Then
            RutaServerReloj = App.Path + "\standard\Att2007.mdb"
            RutaServerEasy = App.Path + "\standard\Att2003.mdb"
            RutaBD = App.Path
    Else
            RutaServerReloj = RutaBD + "\standard\Att2007.mdb"
            RutaServerEasy = RutaBD + "\standard\Att2003.mdb"
    End If
          
          ConexionEasy = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaServerEasy & ";Persist Security Info=False"
          ConexionReloj = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaServerReloj & ";Persist Security Info=False"



    With Me.AdoConsulta
     .ConnectionString = ConexionReloj
    End With

'    With Me.DtaEmpresa
'     .ConnectionString = ConexionReloj
'     .RecordSource = "SELECT DatosEmpresa.* FROM DatosEmpresa"
'     .Refresh
'    End With
    
    With Me.AdoConsultaEasyWay
       .ConnectionString = ConexionEasy
    End With
    
    
    With Me.AdoDispositivos
       .ConnectionString = ConexionEasy
    End With


    '//////////////////////////////////BUSCO EL NOMBRE DE LA EMPRESA ///////////////////////////
    Me.AdoConsultaEasyWay.RecordSource = "SELECT Dept.DeptName, Dept.SupDeptid From Dept WHERE (((Dept.SupDeptid)=0))"
    Me.AdoConsultaEasyWay.Refresh
    If Not Me.AdoConsultaEasyWay.Recordset.EOF Then
'      Me.DtaEmpresa.Recordset("NombreEmpresa") = Me.AdoConsultaEasyWay.Recordset("DeptName")
'      Me.DtaEmpresa.Recordset.Update
    End If
    
    '//////////////////////////////////BUSCO SI EXISTE EL SISTEMA PARA AGREGARLO AL MENU DEL FABRICANTE ///////////////////////////
    Me.AdoConsultaEasyWay.RecordSource = "SELECT OutProg.Progid, OutProg.ProgName, OutProg.ProgPath From OutProg WHERE (((OutProg.ProgName)='REPORTES ZEUS RELOJ')) "
    Me.AdoConsultaEasyWay.Refresh
    If Me.AdoConsultaEasyWay.Recordset.EOF Then
      Me.AdoConsultaEasyWay.Recordset.AddNew
        Me.AdoConsultaEasyWay.Recordset("ProgName") = "REPORTES ZEUS RELOJ"
        Me.AdoConsultaEasyWay.Recordset("ProgPath") = RutaBD + "\standard\Zeus Reloj.exe"
      Me.AdoConsultaEasyWay.Recordset.Update
    Else
      Me.AdoConsultaEasyWay.Recordset("ProgPath") = RutaBD + "\standard\Zeus Reloj.exe"
      Me.AdoConsultaEasyWay.Recordset.Update
    
    End If




'If Not Encontrado Then
'  MsgBox "La Tasa de Hoy no ha sido grabada"
'  frmTasa2.Show 1
'End If

PreguntaSalir = True

If Format(Now, "dd/mm/yyyy") > CDate("24/12/2021") Then
  PreguntaSalir = False
  Unload Me
End If


Exit Sub
TipoErr:
If Not Err.Number = 8002 Then
 MsgBox Err.Description
End If
 
 
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo TipoErrs
'se utiliza para cuando el usario sale con la x de windows y poder validar

If PreguntaSalir = True Then
    k% = MsgBox("Desea Realmente Salir?", vbYesNo)
    If k% <> 6 Then
    Cancel = 1
    Exit Sub
    End If
End If

KillProcess ("SistemaNominas.exe")

Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub mnuactnomina_Click()
FrmActivarNomina.Show
End Sub

Private Sub mnubarraestado_Click()
 ' Llama al procedimiento de la barra de herramientas, pasando una referencia
    ' a esta instancia de formulario.
If StatusBar2.Visible = True Then
   mnubarraestado.Checked = False
   StatusBar2.Visible = False
Else
  StatusBar2.Visible = True
  mnubarraestado.Checked = True
End If
End Sub

Private Sub MnuCalcNomina_Click()
On Error GoTo TipoErrs
FrmCalcularNomina.Show
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub mnucomiproduc_Click()
FrmMovimientos.Show
End Sub

Private Sub mnucomisiones_Click()
FrmTipoComision.Show

End Sub

Private Sub mnuctrol2_Click()
FrmControles.Show
End Sub

Private Sub MnuDeducciones_Click()
FrmDeduccion.Show

End Sub

Private Sub mnudesrenu_Click()
FrmBajas.Show 1
End Sub

Private Sub mnugrabComi_Click()
FrmMovimientos.Show
End Sub

Private Sub mnuhrsextras_Click()
FrmMovimientos.Show
End Sub

Private Sub MnuExtrsFaktas_Click()
FrmCalHEFaltas.Show
End Sub

Private Sub mnuhistorial_Click()
FrmHistorial.Show
End Sub

Private Sub MnuIncentivos_Click()
FrmIncentivo.Show
End Sub

Private Sub mnulistnomina_Click()
FrmListNomina.Show
End Sub

Private Sub mnulstreports_Click()
FrmReportes.CmbReportes.AddItem "Listado de Empleados"
FrmReportes.CmbReportes.AddItem "Reporte INSS"
FrmReportes.Show
End Sub

Private Sub mnumes13Vaca_Click()
Frm13Vaca.Show
End Sub

Private Sub mnumovnomina_Click()
FrmMovimientos.Show
End Sub

Private Sub mnunomsubsidios_Click()
FrmNomSubsidio.Show
End Sub

Private Sub mnuregentsal_Click()
FrmRegistroEntSal.Show 1
End Sub

Private Sub mnuSubsidio_Click()
FrmSubsidio.Show
End Sub

Private Sub mnususpen_Click()
On erro GoTo TipoErrs
MDIPrimero.MousePointer = 11
FrmListSubsidios.Show
MDIPrimero.MousePointer = 0
Exit Sub
TipoErrs:
ControlErrores

End Sub

Private Sub mnutipodestajo_Click()
FrmTipoDestajo.Show

End Sub

Private Sub mnutipodivision_Click()
FrmGrupo.Show
End Sub

Private Sub mosaico_Click()
On Error GoTo TipoErrs
 ' Organiza los formularios hijos en mosaico.
    MDIPrimero.Arrange vbTileHorizontal
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub Niveles_Click()
 FrmEditaNiveles.Show
End Sub

Private Sub Organizar_Click()
On Error GoTo TipoErrs
' Organiza los iconos de los formularios hijos minimizados.
    MDIPrimero.Arrange vbArrangeIcons
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub PFijas_Click()
On Error GoTo TipoErrs
FrmPersecciones.Width = 7095
FrmPersecciones.Show
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub prestamos_Click()
On Error GoTo TipoErrs
MDIPrimero.MousePointer = 11
FrmPrestamos.Show
MDIPrimero.MousePointer = 0
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub RGenerales_Click()
FrmReportes.CmbReportes.AddItem "Listado Horas Extra"
FrmReportes.CmbReportes.AddItem "Listado de Cargos"
FrmReportes.CmbReportes.AddItem "Listado de Departamentos"
FrmReportes.CmbReportes.AddItem "Listado de Tipos de Subsidios"
FrmReportes.CmbReportes.AddItem "Listado de Tipos de Incentivos"
FrmReportes.CmbReportes.AddItem "Listado de Tipos de Deducciones"
FrmReportes.Show
End Sub

Private Sub Salir_Click()
Unload Me
End Sub

Private Sub SmartButton9_Click()

End Sub

Private Sub SmartButton2_Click()

End Sub

Private Sub SmartButton3_Click()

End Sub

Private Sub SmartButton4_Click()

End Sub

Private Sub SmartButton5_Click()

End Sub

Private Sub SmartButton6_Click()

End Sub

Private Sub SmartButton7_Click()

End Sub

Private Sub SmartButton8_Click()

End Sub



Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
On Error GoTo TipoErrs
  ' Utiliza la propiedad Key con la instrucción
    ' SelectCase para especificar una acción.
    Select Case Button.Index
    Case Is = 1           ' Abre archivo.
         If VerEmpleado = True Then
          MDIPrimero.MousePointer = 11
          frmEmpleado.Show
          MDIPrimero.MousePointer = 0
         End If
    Case Is = 2
         If VerRespaldo = True Then
          MDIPrimero.MousePointer = 11
          frmBackup.Height = 5070
          frmBackup.Width = 5160
          frmBackup.Show
          MDIPrimero.MousePointer = 0
         End If
    Case Is = 5
          MDIPrimero.MousePointer = 11
          FrmGrupo.Show
          MDIPrimero.MousePointer = 0
    Case Is = 6             ' Guarda archivo.
         MDIPrimero.MousePointer = 11
         FrmInssIR.Show
         MDIPrimero.MousePointer = 0
    Case Is = 7
         MDIPrimero.MousePointer = 11
         FrmTipoNomina.Show
         MDIPrimero.MousePointer = 0
    Case Is = 8
         MDIPrimero.MousePointer = 11
         FrmCalcularNomina.Show
         MDIPrimero.MousePointer = 0
    Case Is = 9
         MDIPrimero.MousePointer = 11
         FrmActivarNomina.Show
         MDIPrimero.MousePointer = 0
    Case Is = 10
         MDIPrimero.MousePointer = 11
         frmTasa2.Show
         MDIPrimero.MousePointer = 0
    Case Is = 11
         MDIPrimero.MousePointer = 11
         FrmRegistroEntSal.Show 1
         MDIPrimero.MousePointer = 0
    Case Is = 15
          Unload Me
    End Select
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub SSListBarVentas_Click()

End Sub

Private Sub TipoIncapacidad_Click()
 On Error GoTo TipoErrs
MDIPrimero.MousePointer = 11
 FrmTipoIncapacidad.Show
MDIPrimero.MousePointer = 0
Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub TipoNomina_Click()
On Error GoTo TipoErrs
FrmTipoNomina.Show
Exit Sub
TipoErrs:
 ControlErrores

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)
 'Private Sub TabStrip_Click()
 On Error GoTo TipoErrs
 Select Case TabStrip.SelectedItem.Index
    Case Is = 1           ' Abre archivo.
        Frame1.Visible = True
        Frame2.Visible = False
        Frame3.Visible = False
    Case Is = 2
       Frame1.Visible = False
       Frame2.Visible = True
       Frame3.Visible = False
    Case Is = 3
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = True
    Case Is = 4
     
    End Select
    
    Exit Sub
TipoErrs:
 ControlErrores
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
On Error GoTo TipoErrs
  ' Utiliza la propiedad Key con la instrucción
    ' SelectCase para especificar una acción.
    Select Case Button.Index
    Case Is = 1           ' Abre archivo.
         If VerEmpleado = True Then
          MDIPrimero.MousePointer = 11
          frmEmpleado.Show
          MDIPrimero.MousePointer = 0
         End If
    Case Is = 2
        MDIPrimero.MousePointer = 11
        FrmActivarNomina.Show
        MDIPrimero.MousePointer = 0
    Case Is = 5
          MDIPrimero.MousePointer = 11
          FrmGrupo.Show
          MDIPrimero.MousePointer = 0
    Case Is = 6             ' Guarda archivo.
         MDIPrimero.MousePointer = 11
         FrmInssIR.Show
         MDIPrimero.MousePointer = 0
    Case Is = 7
         MDIPrimero.MousePointer = 11
         FrmTipoNomina.Show
         MDIPrimero.MousePointer = 0
    Case Is = 8
         MDIPrimero.MousePointer = 11
         FrmCalcularNomina.Show
         MDIPrimero.MousePointer = 0
    Case Is = 9
         MDIPrimero.MousePointer = 11
         FrmActivarNomina.Show
         MDIPrimero.MousePointer = 0
    Case Is = 10
         MDIPrimero.MousePointer = 11
         frmTasa2.Show
         MDIPrimero.MousePointer = 0
    Case Is = 11
         MDIPrimero.MousePointer = 11
         FrmRegistroEntSal.Show 1
         MDIPrimero.MousePointer = 0
    Case Is = 15
          Unload Me
    End Select
Exit Sub
TipoErrs:
 ControlErrores
End Sub



Private Sub Toolbar1_ListItemClick(ByVal ItemClicked As Listbar.SSListItem)
    Select Case ItemClicked.Key
        Case "Empleado"
          MDIPrimero.MousePointer = 11
          frmEmpleado.Show
          MDIPrimero.MousePointer = 0
        Case "ActivarNomina"
            MDIPrimero.MousePointer = 11
            FrmActivarNomina.Show
            MDIPrimero.MousePointer = 0
        Case "MovimientoProduccion"
            MDIPrimero.MousePointer = 11
            FrmProduccion.Show 1
            MDIPrimero.MousePointer = 0
        Case "MovimientoNomina"
            MDIPrimero.MousePointer = 11
            FrmMovimientos.Show
            MDIPrimero.MousePointer = 0
        Case "CalcularNomina"
            MDIPrimero.MousePointer = 11
            FrmCalcularNomina.Show 1
            MDIPrimero.MousePointer = 0
        Case "Subsidios"
            MDIPrimero.MousePointer = 11
            FrmNomSubsidio.Show
            MDIPrimero.MousePointer = 0
        Case "Periodo"
            MDIPrimero.MousePointer = 11
            frmFecha.Show 1
            MDIPrimero.MousePointer = 0
        Case "Usuario"
            MDIPrimero.MousePointer = 11
            frmClaves.Show
            MDIPrimero.MousePointer = 0
        Case "Referencias"
            MDIPrimero.MousePointer = 11
            FrmReferencias.Show
            MDIPrimero.MousePointer = 11
        Case "Proceso"
             MDIPrimero.MousePointer = 11
            FrmProcesos.Show
            MDIPrimero.MousePointer = 0
        Case "Permisos"
            MDIPrimero.MousePointer = 11
            FrmPermiso.Show
            MDIPrimero.MousePointer = 0
        Case "Metas"
            MDIPrimero.MousePointer = 11
           FrmIncentivoMetas.Show
            MDIPrimero.MousePointer = 0
        Case "Inss"
            MDIPrimero.MousePointer = 11
            FrmInssIR.Show
            MDIPrimero.MousePointer = 0
        Case "TipoNomina"
            MDIPrimero.MousePointer = 11
           FrmTipoNomina.Show
            MDIPrimero.MousePointer = 0
        Case "ListadoNomina"
           MDIPrimero.MousePointer = 11
           FrmListNomina.Show 1
           MDIPrimero.MousePointer = 0
        
        Case "Adelanto13vo"
            MDIPrimero.MousePointer = 11
            FrmAdelantos13vo.Show 1
            MDIPrimero.MousePointer = 0
        Case "Despido"
            MDIPrimero.MousePointer = 11
            FrmBajas.Show 1
            MDIPrimero.MousePointer = 0
        Case "Vacaciones"
            MDIPrimero.MousePointer = 11
            Frm13Vaca.Show 1
            MDIPrimero.MousePointer = 0
        Case "Reloj"
            MDIPrimero.MousePointer = 11
            frmReloj.Show
            MDIPrimero.MousePointer = 0
        Case "Asistencia"
            MDIPrimero.MousePointer = 11
            frmRepAsistencia.Show
            MDIPrimero.MousePointer = 0
        Case "PeriodoFiscal"
            MDIPrimero.MousePointer = 11
            FrmPeriodoFiscal.Show
            MDIPrimero.MousePointer = 0
        
         Case "Configuracion"
            MDIPrimero.MousePointer = 11
            FrmConfiguracion.Show
            MDIPrimero.MousePointer = 0
        
        
    End Select
End Sub

Private Sub Usuarios_Click()
On Error GoTo TipoErrs
MDIPrimero.MousePointer = 11
frmClaves.Show
MDIPrimero.MousePointer = 0
Exit Sub
TipoErrs:
 ControlErrores
 
End Sub

Private Sub wndTaskPanel_ItemClick(ByVal item As XtremeTaskPanel.ITaskPanelGroupItem)
    Select Case item.Caption
        Case "Listado Nominas de Vacaciones/13vo"
          MDIPrimero.MousePointer = 11
          FrmListadoNominaVacaciones.Show
          MDIPrimero.MousePointer = 0
        Case "Empleados"
          MDIPrimero.MousePointer = 11
          frmEmpleado.Show
          MDIPrimero.MousePointer = 0
        Case "Activar Nominas"
            MDIPrimero.MousePointer = 11
            FrmActivarNomina.Show
            MDIPrimero.MousePointer = 0
        Case "Movimiento de Produccion"
            MDIPrimero.MousePointer = 11
            FrmProduccion.Show 1
            MDIPrimero.MousePointer = 0
        Case "Horas Extras"
            MDIPrimero.MousePointer = 11
            FrmMovimientos.Show
            MDIPrimero.MousePointer = 0
        Case "Calcular Nomina"
            MDIPrimero.MousePointer = 11
            FrmCalcularNomina.Show 1
            MDIPrimero.MousePointer = 0
        Case "Subsidios"
            MDIPrimero.MousePointer = 11
            FrmNomSubsidio.Show
            MDIPrimero.MousePointer = 0
        Case "Periodo Nomina"
            MDIPrimero.MousePointer = 11
            frmFecha.Show 1
            MDIPrimero.MousePointer = 0
        Case "Usuarios"
            MDIPrimero.MousePointer = 11
            frmClaves.Show
            MDIPrimero.MousePointer = 0
        Case "Referencias"
            MDIPrimero.MousePointer = 11
            FrmReferencias.Show
            MDIPrimero.MousePointer = 0
        Case "Procesos"
             MDIPrimero.MousePointer = 11
            FrmProcesos.Show
            MDIPrimero.MousePointer = 0
        Case "Permisos"
            MDIPrimero.MousePointer = 11
            FrmPermiso.Show
            MDIPrimero.MousePointer = 0
        Case "Incentivo x Metas"
            MDIPrimero.MousePointer = 11
           FrmIncentivoMetas.Show
            MDIPrimero.MousePointer = 0
        Case "Inss"
            MDIPrimero.MousePointer = 11
            FrmInssIR.Show
            MDIPrimero.MousePointer = 0
        Case "TipoNomina"
            MDIPrimero.MousePointer = 11
           FrmTipoNomina.Show
            MDIPrimero.MousePointer = 0
        Case "ListadoNomina"
           MDIPrimero.MousePointer = 11
           FrmListNomina.Show 1
           MDIPrimero.MousePointer = 0
        
        Case "Adelanto13vo"
            MDIPrimero.MousePointer = 11
            FrmAdelantos13vo.Show 1
            MDIPrimero.MousePointer = 0
        Case "Despido"
            MDIPrimero.MousePointer = 11
            FrmBajas.Show 1
            MDIPrimero.MousePointer = 0
        Case "Vacaciones"
            MDIPrimero.MousePointer = 11
            Frm13Vaca.Show 1
            MDIPrimero.MousePointer = 0
        Case "Reloj"
            MDIPrimero.MousePointer = 11
            frmReloj.Show
            MDIPrimero.MousePointer = 0
        Case "Asistencia"
            MDIPrimero.MousePointer = 11
            frmRepAsistencia.Show
            MDIPrimero.MousePointer = 0
        Case "Periodo Fiscal"
            MDIPrimero.MousePointer = 11
            FrmPeriodoFiscal.Show
            MDIPrimero.MousePointer = 0
        
         Case "Configuracion"
            MDIPrimero.MousePointer = 11
            FrmConfiguracion.Show
            MDIPrimero.MousePointer = 0
            
        
         Case "Departamento"
            MDIPrimero.MousePointer = 11
            FrmDepartamentos.Show
            MDIPrimero.MousePointer = 0
            
          Case "Cargos"
            MDIPrimero.MousePointer = 11
            FrmCargo.Show
            MDIPrimero.MousePointer = 0
            
          Case "Tipo Incapacidad"
            MDIPrimero.MousePointer = 11
            FrmTipoIncapacidad.Show
            MDIPrimero.MousePointer = 0
            
          Case "Incapacidades"
            MDIPrimero.MousePointer = 11
            FrmIncapacidades.Show
            MDIPrimero.MousePointer = 0
            
          Case "Tipo Incentivo"
            MDIPrimero.MousePointer = 11
            FrmIncentivo.Show
            MDIPrimero.MousePointer = 0
            
            
          Case "Tipo Deducciones"
            MDIPrimero.MousePointer = 11
            FrmDeduccion.Show
            MDIPrimero.MousePointer = 0
            
          Case "Tipo Deducciones"
            MDIPrimero.MousePointer = 11
            FrmDeduccion.Show
            MDIPrimero.MousePointer = 0
            
          Case "Tipo Subsidio"
            MDIPrimero.MousePointer = 11
            FrmSubsidio.Show
            MDIPrimero.MousePointer = 0
            
          Case "Tipo Comision"
            MDIPrimero.MousePointer = 11
            FrmTipoComision.Show
            MDIPrimero.MousePointer = 0
            
          Case "Tipo Destajo"
            MDIPrimero.MousePointer = 11
            FrmTipoDestajo.Show
            MDIPrimero.MousePointer = 0
            
          Case "Division Nomina"
            MDIPrimero.MousePointer = 11
            FrmGrupo.Show
            MDIPrimero.MousePointer = 0
            
          Case "Tipo Nomina"
            MDIPrimero.MousePointer = 11
            FrmTipoNomina.Show
            MDIPrimero.MousePointer = 0
            
          Case "Tabla INSS/IR"
            MDIPrimero.MousePointer = 11
            FrmInssIR.Show
            MDIPrimero.MousePointer = 0
            
          Case "Produccion"
            MDIPrimero.MousePointer = 11
            FrmMovimientos.Show
            MDIPrimero.MousePointer = 0
            
          Case "Listado de Nominas"
            MDIPrimero.MousePointer = 11
            FrmListNomina.Show
            MDIPrimero.MousePointer = 0
            
          Case "Suspenciones"
            MDIPrimero.MousePointer = 11
            FrmSuspencion.Show
            MDIPrimero.MousePointer = 0
            
          Case "Historial Salarial"
            MDIPrimero.MousePointer = 11
            FrmHistorial.Show
            MDIPrimero.MousePointer = 0
                        
           Case "Tasa de Cambio"
            MDIPrimero.MousePointer = 11
            frmTasa2.Show
            MDIPrimero.MousePointer = 0
            
           Case "Informacion de Usuarios"
            MDIPrimero.MousePointer = 11
            FrmInforme.Show
            MDIPrimero.MousePointer = 0
            
           Case "Calculadora"
            MDIPrimero.MousePointer = 11
            Directorio = App.Path + "\Calc.exe"
            Directorio = Shell(Directorio)
            MDIPrimero.MousePointer = 0
                        
            Case "Controles Personalizados"
            MDIPrimero.MousePointer = 11
            FrmControles.Show
            MDIPrimero.MousePointer = 0
            
            Case "Produccion Manual"
            MDIPrimero.MousePointer = 11
            FrmProduccionManual.Show
            MDIPrimero.MousePointer = 0
            
            Case "Reportes Generales"
            MDIPrimero.MousePointer = 11
            FrmReportes.CmbReportes.AddItem "Listado de Cargos"
            FrmReportes.CmbReportes.AddItem "Listado de Departamentos"
            FrmReportes.CmbReportes.AddItem "Listado de Tipos de Subsidios"
            FrmReportes.CmbReportes.AddItem "Listado de Tipos de Incentivos"
            FrmReportes.CmbReportes.AddItem "Listado de Tipos de Deducciones"
            
            FrmReportes.Show 1
            MDIPrimero.MousePointer = 0
            
            Case "Reportes de Empleados"
            FrmReportes.CmbReportes.AddItem "Numeros Disponibles"
            FrmReportes.CmbReportes.AddItem "Reporte x Produccion"
            FrmReportes.CmbReportes.AddItem "Reporte x Produccion Linea"
            FrmReportes.CmbReportes.AddItem "Analisis Produccion"
            FrmReportes.CmbReportes.AddItem "Lista de Empleados Activos"
            FrmReportes.CmbReportes.AddItem "Listado Maestro de Empleados"
            FrmReportes.CmbReportes.AddItem "Salario Basico Vrs Produccion"
            FrmReportes.CmbReportes.AddItem "Adelantos 13vo y Vacaciones"
            FrmReportes.CmbReportes.AddItem "Resumen-Pago Mensual"
            FrmReportes.CmbReportes.AddItem "Total-Pago Mensual"
            FrmReportes.CmbReportes.AddItem "Detalle Deducciones"
            FrmReportes.CmbReportes.AddItem "Reporte Devengado"
            FrmReportes.CmbReportes.AddItem "Reporte Carnet Empleados"
            FrmReportes.Show 1
            
            Case "Reportes de Deducciones"
            FrmReportes.CmbReportes.AddItem "Reporte Inss"
            FrmReportes.CmbReportes.AddItem "Reporte Detalle Inss"
            FrmReportes.CmbReportes.AddItem "Reporte Inss 2"
            FrmReportes.CmbReportes.AddItem "EXPORTACION INSS"
            FrmReportes.CmbReportes.AddItem "Reporte Ir"
            FrmReportes.CmbReportes.AddItem "Reporte Detalle Ir"
            FrmReportes.CmbReportes.AddItem "Reporte IR MENSUAL"
            FrmReportes.CmbReportes.AddItem "Reporte IR MENSUAL DETALLADO"
            FrmReportes.CmbReportes.AddItem "Reporte INSS E IR MENSUAL"
            FrmReportes.CmbReportes.AddItem "Reporte Detalle Deducciones"
            
            
            FrmReportes.Show 1
            FrmReportes.TDBCombo2.Visible = True
        
        'UPDATE: ELIAZAR POLANCO
        'CATALOGOS
        Case "Grupo de Puntos": frmPuntosGrupo.Show
        Case "Puntos": frmPuntos.Show
        Case "Administrador de Actividades": frmActividades.Show
        
        'PROCESOS
        Case "Complementos Salariales": frmPuntosSolicitudes.Show
        Case "Solicitudes de Puntos": frmPuntosAutorizar.Show
        Case "Planificación de Actividades": frmProgramacion.Show
        Case "Administrador de Horas Laborales": frmProduccionAct.Show
        Case "Aprobar Horas Extras": frmExtrasAutorizar.Show
        
    End Select


End Sub

