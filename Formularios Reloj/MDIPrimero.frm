VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "Codejock.CommandBars.v12.0.0.Demo.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#12.0#0"; "Codejock.DockingPane.v12.0.0.Demo.ocx"
Begin VB.MDIForm MDIPrimero 
   BackColor       =   &H80000009&
   Caption         =   "Zeus Reportes Control de Asistencias"
   ClientHeight    =   8055
   ClientLeft      =   -75
   ClientTop       =   150
   ClientWidth     =   12165
   HelpContextID   =   1
   Icon            =   "MDIPrimero.frx":0000
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "MDIPrimero.frx":0442
   NegotiateToolbars=   0   'False
   OLEDropMode     =   1  'Manual
   Picture         =   "MDIPrimero.frx":0884
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc AdoConsulta 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      Top             =   375
      Visible         =   0   'False
      Width           =   12165
      _ExtentX        =   21458
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
   Begin MSAdodcLib.Adodc DtaEmpresa 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      Top             =   7170
      Visible         =   0   'False
      Width           =   12165
      _ExtentX        =   21458
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
      Top             =   7650
      WhatsThisHelpID =   1
      Width           =   12165
      _ExtentX        =   21458
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
            Picture         =   "MDIPrimero.frx":7A505
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
            TextSave        =   "MAY�S"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Object.Width           =   1393
            MinWidth        =   1393
            TextSave        =   "N�M"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "08:37 a.m."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "MDIPrimero.frx":7A81F
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
      OleObjectBlob   =   "MDIPrimero.frx":7AB39
      Top             =   3720
   End
   Begin MSComctlLib.ImageList imlToolbarIcons2 
      Left            =   9120
      Top             =   1680
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
            Picture         =   "MDIPrimero.frx":2D6366
            Key             =   ""
            Object.Tag             =   "119"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D6900
            Key             =   ""
            Object.Tag             =   "113"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D6E9A
            Key             =   ""
            Object.Tag             =   "128"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D7434
            Key             =   ""
            Object.Tag             =   "115"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D77CE
            Key             =   ""
            Object.Tag             =   "130"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D7D68
            Key             =   ""
            Object.Tag             =   "160"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D8302
            Key             =   ""
            Object.Tag             =   "116"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D889C
            Key             =   ""
            Object.Tag             =   "300"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D8E36
            Key             =   ""
            Object.Tag             =   "118"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2D93D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2DB752
            Key             =   ""
            Object.Tag             =   "204"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2DBCEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2DE06E
            Key             =   ""
            Object.Tag             =   "129"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2DE608
            Key             =   ""
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2DEBA2
            Key             =   ""
            Object.Tag             =   "205"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2DF13C
            Key             =   ""
            Object.Tag             =   "1331"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2DF6D6
            Key             =   ""
            Object.Tag             =   "1311"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E13E0
            Key             =   ""
            Object.Tag             =   "131"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E197A
            Key             =   ""
            Object.Tag             =   "0"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E1F14
            Key             =   ""
            Object.Tag             =   "133"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E24AE
            Key             =   ""
            Object.Tag             =   "132"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E2A48
            Key             =   ""
            Object.Tag             =   "134"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E2FE2
            Key             =   ""
            Object.Tag             =   "140111"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E357C
            Key             =   ""
            Object.Tag             =   "139"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E3B16
            Key             =   ""
            Object.Tag             =   "320"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E40B0
            Key             =   ""
            Object.Tag             =   "136"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   8280
      Top             =   1680
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
            Picture         =   "MDIPrimero.frx":2E464A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E475C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E48AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E49C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E4AD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrimero.frx":2E4BE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdoConsultaEasyWay 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      Top             =   6690
      Visible         =   0   'False
      Width           =   12165
      _ExtentX        =   21458
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
      Top             =   6210
      Visible         =   0   'False
      Width           =   12165
      _ExtentX        =   21458
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
      Width           =   12165
      _ExtentX        =   21458
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
      Left            =   5880
      Top             =   4440
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
            Picture         =   "MDIPrimero.frx":2E4CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2E5948
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2E659A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2E71EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2E7E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2E8A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2E96E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2EA334
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2EAF86
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2EBBD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2EC82A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2ED47C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2EE0CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2EED20
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2EF972
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F05C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F1216
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F1E68
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F2ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F370C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F435E
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F4FB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F5C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F6854
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F74A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F80F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F8D4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2F999C
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2FA5EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2FB240
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2FBE92
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2FCAE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2FD736
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2FE388
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2FEFDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrimero.frx":2FFC2C
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
      Begin VB.Menu Compa�ias 
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
            Caption         =   "Tipo de Comisi�n"
         End
         Begin VB.Menu mnutipodestajo 
            Caption         =   "Tipo de Destajo"
         End
         Begin VB.Menu mnutipodivision 
            Caption         =   "&Divisiones de N�mina"
         End
         Begin VB.Menu TipoNomina 
            Caption         =   "T&ipo N�mina"
            HelpContextID   =   28
         End
         Begin VB.Menu Inss 
            Caption         =   "&Tablas INSS, IR"
            HelpContextID   =   29
         End
         Begin VB.Menu mnulistnomina 
            Caption         =   "Listado de N�minas elaboradas"
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
         Caption         =   "Movimientos de N�mina"
      End
      Begin VB.Menu mnuactnomina 
         Caption         =   "Activar N�mina"
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCalcNomina 
         Caption         =   "Calcular N�mina"
      End
      Begin VB.Menu mnumes13Vaca 
         Caption         =   "Calcular e Imprimir el 13vo mes Y/o Vacaciones"
      End
      Begin VB.Menu mnunomsubsidios 
         Caption         =   "N�mina de Subsidios"
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
  Dim Directorio As String
  Dim A�O1 As String, A�O2 As String, A�O3 As String
  
  

      Select Case Control.Id
        Case 1300: Unload Me
        Case 1703:
                  MDIPrimero.MousePointer = 11
                   FrmControles.Show 1
                  MDIPrimero.MousePointer = 0
        Case 1700:
                  MDIPrimero.MousePointer = 11
                   Quien = "Reportes Generales"
                   FrmReportes.Show
                  MDIPrimero.MousePointer = 0
         Case 1701:
                  MDIPrimero.MousePointer = 11
                   Quien = "Reportes Asistencia"
                   FrmReportes.Show
                  MDIPrimero.MousePointer = 0
         Case 1704:
                  MDIPrimero.MousePointer = 11
                   Quien = "Reportes Asistencia"
                   FrmDepartamentos.Show
                  MDIPrimero.MousePointer = 0
         Case 1705:
                  MDIPrimero.MousePointer = 11
                   Quien = "Reportes Asistencia"
                   FrmExportar.Show
                  MDIPrimero.MousePointer = 0
         Case 1706:
                  MDIPrimero.MousePointer = 11
                   Quien = "Reportes Asistencia"
                   FrmAuxiliar.Show
                  MDIPrimero.MousePointer = 0
         Case 1707:
                  MDIPrimero.MousePointer = 11
                   Quien = "Reportes Asistencia"
                   FrmJornadas.Show
                  MDIPrimero.MousePointer = 0
         Case 1708:
                  MDIPrimero.MousePointer = 11
                   Quien = "Reportes Asistencia"
                   FrmAsignacion.Show
                  MDIPrimero.MousePointer = 0
         Case 1709:
                  MDIPrimero.MousePointer = 11
                   Quien = "Reportes Asistencia"
                   FrmHorarios.Show
                  MDIPrimero.MousePointer = 0
         Case 1710:
                  MDIPrimero.MousePointer = 11
                   Quien = "Reportes Asistencia"
                   FrmNominas.Show
                  MDIPrimero.MousePointer = 0
         Case 1711:
                  MDIPrimero.MousePointer = 11
                    FrmImportacion.Show
                  MDIPrimero.MousePointer = 0
         Case 1712:
                  MDIPrimero.MousePointer = 11
                    FrmExportarExcel.Show 1
                  MDIPrimero.MousePointer = 0
        End Select
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
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ReportesGenerales.png", 1700, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\HorasExtra.png", 1701, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Estadisticos.png", 1702, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Controles.png", 1703, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Departamentos.png", 1704, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Exportar.png", 1705, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Tarjeta.png", 1706, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Jornada.png", 1707, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\AsignacionJornada.png", 1708, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Horarios.png", 1709, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Nominas.png", 1710, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\Excel2.png", 1711, xtpImageNormal
    CommandBars.Icons.LoadBitmap App.Path & "\Imagenes\ExportarExcel5.png", 1712, xtpImageNormal
    '
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
    RibbonBar.QuickAccessControls.Add XtremeCommandBars.XTPControlType.xtpControlButton, ID_FILE_SAVE, "Zeus Reloj", False, False
'
'    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'    '///////////////////////////////////CREO LOS TABS//////////////////////////////////////////////////
'    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Set TabHome = RibbonBar.InsertTab(0, "&Accesos")
    TabHome.Id = 130
        Set GroupFile = TabHome.Groups.AddGroup("Reportes", 1)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1700, "&Reportes Generales", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1701, "&Reportes Asistencia", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1702, "&Reportes Estadisticos", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
         Set GroupFile = TabHome.Groups.AddGroup("Tablas", 2)
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1704, "&Departamentos", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1706, "&Tarjeta de Marcadas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1707, "&Jornadas Laborales", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1708, "&Asignacion de Jornadas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1709, "&Configuracion Horarios", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////CREO EL TABS DE OPCIONES//////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////////////////////////
    Set TabHome = RibbonBar.InsertTab(5, "&Opciones")
    TabHome.Id = 1500
     Set GroupFile = TabHome.Groups.AddGroup("Basicos", 1)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1703, "Controles Personalizados", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1705, "Exportar Archivos TXT", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1712, "Exportar Archivos XLS", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1710, "Movimientos Nominas", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
     Set item = GroupFile.Add(XtremeCommandBars.XTPControlType.xtpControlButton, 1711, "Importar Archivo", False, False)
     item.Style = xtpButtonIconAndCaptionBelow
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
    
    Group.Items.Add 7, "Complementos Salariales", xtpTaskItemTypeLink, 9
    Group.Items.Add 8, "Solicitudes de Puntos", xtpTaskItemTypeLink, 8
    Group.Items.Add 9, "Planificaci�n de Actividades", xtpTaskItemTypeLink, 7
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
    Group.Items.Add 19, "Divicion Nomina", xtpTaskItemTypeLink, 19
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
 MsgBox Err.Description
End Sub

Private Sub Claves_Click()
On Error GoTo TipoErrs
frmClaves.Height = 3060
frmClaves.Width = 4230
frmClaves.Show
Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub AgregaClaves_Click()
On Error GoTo TipoErrs
frmClaves.Show
Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub


Private Sub Contendido_Click()
CommonDialog1.CancelError = True
CommonDialog1.HelpCommand = &H9&
CommonDialog1.HelpFile = App.Path + "\Zeus.hlp"
CommonDialog1.ShowHelp
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
 MsgBox Err.Description
End Sub
Private Sub imgCopyButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Muestra la imagen del estado presionado.
    imgCopyButton.Picture = imgCopyButtonDn.Picture
End Sub
Function LoadIcon(Path As String, cx As Long, cy As Long) As Long
    LoadIcon = LoadImage(App.hInstance, App.Path + "\" + Path, 1, cx, cy, 16)
End Function
Private Sub imgCopyButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Si el bot�n est� presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su �rea; si no
    ' presenta el mapa de bits del estado presionado.
    Select Case Button
    Case 1
        If x <= 0 Or x > imgCopyButton.Width Or y < 0 Or y > imgCopyButton.Height Then
            imgCopyButton.Picture = imgCopyButtonUp.Picture
        Else
            imgCopyButton.Picture = imgCopyButtonDn.Picture
        End If
    End Select
End Sub
Private Sub imgCopyButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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
 MsgBox Err.Description
End Sub

Private Sub ImgEmpleado_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' Muestra la imagen del estado presionado.
    ImgEmpleado.Picture = imgEmpleadoDn.Picture
End Sub
Private Sub ImgEmpleado_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Si el bot�n est� presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su �rea; si no
    ' presenta el mapa de bits del estado presionado.
    Select Case Button
    Case 1
        If x <= 0 Or x > ImgEmpleado.Width Or y < 0 Or y > ImgEmpleado.Height Then
            ImgEmpleado.Picture = imgEmpleadoUp.Picture
        Else
            ImgEmpleado.Picture = imgEmpleadoDn.Picture
        End If
    End Select
End Sub
Private Sub ImgEmpleado_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Muestra la imagen del estado sin presionar.
    ImgEmpleado.Picture = imgEmpleadoUp.Picture
End Sub

Private Sub imgSalir_Click()
  ' Actualiza la imagen.
    imgSalir.Refresh
   
    Unload Me
End Sub
Private Sub imgSalir_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Muestra la imagen del estado presionado.
    imgSalir.Picture = imgSalirDn.Picture
End Sub

Private Sub imgSalir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo TipoErrs
    ' Si el bot�n est� presionado, presenta el mapa de bits del estado sin presionar
    ' cuando el mouse se arrastra fuera de su �rea; si no
    ' presenta el mapa de bits del estado presionado.
    Select Case Button
    Case 1
        If x <= 0 Or x > imgSalir.Width Or y < 0 Or y > imgSalir.Height Then
            imgSalir.Picture = imgSalirUp.Picture
        Else
            imgSalir.Picture = imgSalirDn.Picture
        End If
    End Select
Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub
Private Sub imgSalir_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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
 MsgBox Err.Description
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
 MsgBox Err.Description
End Sub

Private Sub IR_Click()
 On Error GoTo TipoErrs
 FrmInssIR.Width = 6495
 FrmInssIR.Height = 3435
 FrmInssIR.Show
 Exit Sub
TipoErrs:
 MsgBox Err.Description
 End Sub

Private Sub MCascada_Click()
' Organiza los formularios hijos en cascada.
    MDIPrimero.Arrange vbCascade
End Sub

Private Sub MDIForm_Activate()
 
MDIPrimero.HelpContextID = (1)



End Sub
Private Sub MDIForm_Load()



Dim IDNumber As String, IpAdress As String, ri As Long
Dim RutaConexion As String, ConexionBD As String
Dim RutaBD As String



    '////////////////////////BUSCO EL DIRECTORIO Y RUTA DE LAS BASE DE DATOS ///////////////////////////////
    RutaConexion = App.Path + "\CntReloj.dll"
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
            RutaServer = App.Path + "\Att2007.mdb"
            RutaServerEasy = App.Path + "\Att2003.mdb"
            RutaBD = App.Path
          Else
            RutaServer = RutaBD + "\Att2007.mdb"
            RutaServerEasy = RutaBD + "\Att2003.mdb"
          End If
          
          ConexionEasy = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaServerEasy & ";Persist Security Info=False"
          Conexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaServer & ";Persist Security Info=False"



With Me.AdoConsulta
 .ConnectionString = Conexion
End With



With Me.DtaEmpresa
 .ConnectionString = Conexion
 .RecordSource = "SELECT DatosEmpresa.* FROM DatosEmpresa"
 .Refresh
End With

With Me.AdoConsultaEasyWay
   .ConnectionString = ConexionEasy
End With


With Me.AdoDispositivos
   .ConnectionString = ConexionEasy
End With

''//////////////////////////////////BUSCO TODOS LOS DISPOSITIVOS //////////////////////////////////////////////////
'Me.AdoDispositivos.RecordSource = "SELECT FingerClient.Clientid, FingerClient.IPaddress, FingerClient.ClientName, FingerClient.ClientNumber FROM FingerClient "
'Me.AdoDispositivos.Refresh
'Do While Not Me.AdoDispositivos.Recordset.EOF
'
'  IDNumber = Me.AdoDispositivos.Recordset("Clientid")
'  IpAdress = Me.AdoDispositivos.Recordset("IPaddress")
'
'  ri = CKT_RegisterNet(IDNumber, IpAdress) 'if from net
'  If ri = 1 Then  'if from USB
'     MsgBox ("CKT_RegisterNet OK")
'  End If
'
'  Me.AdoDispositivos.Recordset.MoveNext
'Loop

'//////////////////////////////////BUSCO EL NOMBRE DE LA EMPRESA ///////////////////////////
Me.AdoConsultaEasyWay.RecordSource = "SELECT Dept.DeptName, Dept.SupDeptid From Dept WHERE (((Dept.SupDeptid)=0))"
Me.AdoConsultaEasyWay.Refresh
If Not Me.AdoConsultaEasyWay.Recordset.EOF Then
  Me.DtaEmpresa.Recordset("NombreEmpresa") = Me.AdoConsultaEasyWay.Recordset("DeptName")
  Me.DtaEmpresa.Recordset.Update
End If

'//////////////////////////////////BUSCO SI EXISTE EL SISTEMA PARA AGREGARLO AL MENU DEL FABRICANTE ///////////////////////////
Me.AdoConsultaEasyWay.RecordSource = "SELECT OutProg.Progid, OutProg.ProgName, OutProg.ProgPath From OutProg WHERE (((OutProg.ProgName)='REPORTES ZEUS RELOJ')) "
Me.AdoConsultaEasyWay.Refresh
If Me.AdoConsultaEasyWay.Recordset.EOF Then
  Me.AdoConsultaEasyWay.Recordset.AddNew
    Me.AdoConsultaEasyWay.Recordset("ProgName") = "REPORTES ZEUS RELOJ"
    Me.AdoConsultaEasyWay.Recordset("ProgPath") = RutaBD + "\Zeus Reloj.exe"
  Me.AdoConsultaEasyWay.Recordset.Update
Else
  Me.AdoConsultaEasyWay.Recordset("ProgPath") = RutaBD + "\Zeus Reloj.exe"
  Me.AdoConsultaEasyWay.Recordset.Update

End If

DtaEmpresa.Refresh
Titulo = DtaEmpresa.Recordset("NombreEmpresa")
SubTitulo = DtaEmpresa.Recordset("Direccion") + " RUC: " + DtaEmpresa.Recordset("NumeroRuc")
''RutaLogo = DtaEmpresa.Recordset.RutaLogo
'StatusBar2.Panels(2) = "Licencia: " + Titulo


Set item = PopupControl1.AddItem(50, 15, 270, 45, Titulo)
item.TextColor = RGB(0, 61, 178)
item.Bold = True

Set item = PopupControl1.AddItem(12, 20, 12, 27, "")
item.SetIcon LoadIcon("Imagenes\Imagen.ico", 32, 32), xtpPopupItemIconNormal



Set item = PopupControl1.AddItem(50, 29, 400, 100, "Direc:" & DtaEmpresa.Recordset("Direccion"))
item.TextColor = RGB(0, 61, 178)
item.Bold = True
Set item = PopupControl1.AddItem(60, 60, 400, 100, "ZEUS RELOJ  ")
    item.Bold = True
    PopupControl1.VisualTheme = xtpPopupThemeOffice2003
    PopupControl1.SetSize 300, 110
    Me.PopupControl1.Show
    Me.PopupControl1.Show



CargarInterfaz

CreateRibbonBar

RibbonBar.EnableFrameTheme


Exit Sub
TipoErr:
If Not Err.Number = 8002 Then
 MsgBox Err.Description
End If
 
 
End Sub







