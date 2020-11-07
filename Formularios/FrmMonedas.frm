VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmMonedas 
   Caption         =   "Denominación de Billetes en Pago de Nómina"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtTot1000 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   61
      Text            =   "0.00"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Txt1000 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   60
      Text            =   "0"
      Top             =   2160
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc AdoBusca 
      Height          =   375
      Left            =   600
      Top             =   8040
      Width           =   3975
      _ExtentX        =   7011
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
   Begin VB.TextBox TxtTot200 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   56
      Text            =   "0.00"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Txt200 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   "0"
      Top             =   3120
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "FrmMonedas.frx":0000
      Top             =   5880
   End
   Begin VB.TextBox Txt500 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "0"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox TxtTot500 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   51
      Text            =   "0.00"
      Top             =   2640
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc DtaDetalleNomina 
      Height          =   495
      Left            =   480
      Top             =   7320
      Width           =   4095
      _ExtentX        =   7223
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
      Caption         =   "DtaDetalleNomina"
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
      Left            =   2160
      Top             =   6720
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.CommandButton CmdOtorga 
      Caption         =   "Otorgar"
      Height          =   255
      Left            =   2520
      TabIndex        =   50
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox TxtGranTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   49
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton CmdSalir 
      DownPicture     =   "FrmMonedas.frx":0234
      Height          =   375
      Left            =   6840
      Picture         =   "FrmMonedas.frx":1D16
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      DownPicture     =   "FrmMonedas.frx":37F8
      Height          =   375
      Left            =   4800
      Picture         =   "FrmMonedas.frx":65FA
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2520
      TabIndex        =   35
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox TxtTotD01 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "0.00"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox TxtTotD05 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "0.00"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox TxtTotD10 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "0.00"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox TxtTotD25 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "0.00"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox TxtTotD50 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "0.00"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox TxtTot1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "0.00"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox TxtTot5 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "0.00"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox TxtTot10 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "0.00"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox TxtTot20 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "0.00"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox TxtTot50 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "0.00"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox TxtTot100 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "0.00"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox TxtD01 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox TxtD05 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "0"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox TxtD10 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox TxtD25 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox TxtD50 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Txt1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox Txt5 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Txt10 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Txt20 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Txt50 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Txt100 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   3600
      Width           =   1095
   End
   Begin XtremeSuiteControls.ProgressBar PBMonedas 
      Height          =   375
      Left            =   120
      TabIndex        =   59
      Top             =   5760
      Width           =   8415
      _Version        =   786432
      _ExtentX        =   14843
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   63
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      Caption         =   "1000 *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   62
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   58
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      Caption         =   "200 *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   57
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      Caption         =   "500 *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   54
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   53
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   46
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   45
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   44
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   43
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   42
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   41
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   40
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   39
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   38
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   37
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   36
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   34
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "0.01"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "0.05"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   21
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "0.10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   20
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "0.25"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "0.50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "1 *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "5 *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "10 *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "20 *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "50 *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   " 100 *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Denominaciones de Billetes y Monedas para la Nómina"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1665
      Left            =   120
      Picture         =   "FrmMonedas.frx":8BBC
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2280
   End
End
Attribute VB_Name = "FrmMonedas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Ajuste As Double


Private Sub CmdOtorga_Click()
On Error GoTo TipoErr
Dim CantDiv As Double
Dim Total As Double


Dim Enteros As Long
Dim Decimales As Double

Total = (txtTotal.Text)
'Total = Me.txtGranTotal.Text
Enteros = Int(Total)
Decimales = Format(Total - Enteros, "#0.00000000000")

'reviso si hay billeres de 1000
If Total >= 1000 Then
    CantDiv = Int(Total / 1000)
    If CantDiv > 0 Then
       Txt1000.Text = val(Txt1000.Text) + CantDiv
       'MsgBox Total
       'MsgBox (CantDiv * 100)
       Total = Total - (CantDiv * 1000)
    Else
       Txt1000.Text = 0
    End If
End If

'reviso si hay billeres de 500
If Total >= 500 Then
    CantDiv = Int(Total / 500)
    If CantDiv > 0 Then
       Txt500.Text = val(Txt500.Text) + CantDiv
       'MsgBox Total
       'MsgBox (CantDiv * 100)
       Total = Total - (CantDiv * 500)
    Else
       Txt500.Text = 0
    End If
End If

'reviso si hay billeres de 500
If Total >= 200 Then
    CantDiv = Int(Total / 200)
    If CantDiv > 0 Then
       Txt200.Text = val(Txt200.Text) + CantDiv
       'MsgBox Total
       'MsgBox (CantDiv * 100)
       Total = Total - (CantDiv * 200)
    Else
       Txt200.Text = 0
    End If
End If


'reviso si hay billeres de 100
If Total >= 100 Then
    CantDiv = Int(Total / 100)
    If CantDiv > 0 Then
       Txt100.Text = CDbl(Txt100.Text) + CantDiv
       'MsgBox Total
       'MsgBox (CantDiv * 100)
       Total = Total - (CantDiv * 100)
    Else
       Txt100.Text = 0
    End If
End If
'reviso si hay billeres de 50
If Total >= 50 Then
    CantDiv = Int(Total / 50)
    If CantDiv > 0 Then
       Txt50.Text = CDbl(Txt50.Text) + CantDiv
       Total = Total - CantDiv * 50
    Else
       Txt50.Text = 0
    End If
End If
'reviso si hay billeres de 20
If Total >= 20 Then
    CantDiv = Int(Total / 20)
    If CantDiv > 0 Then
       Txt20.Text = CDbl(Txt20.Text) + CantDiv
       Total = Total - CantDiv * 20
    Else
       Txt20.Text = 0
    End If
End If

'reviso si hay billeres de 10
If Total >= 10 Then
    CantDiv = Int(Total / 10)
    If CantDiv > 0 Then
       Txt10.Text = CDbl(Txt10.Text) + CantDiv
       Total = Total - CantDiv * 10
    Else
      Txt10.Text = 0
    End If
End If
'reviso si hay monedas de 5
If Total >= 5 Then
    CantDiv = Int(Total / 5)
    If CantDiv > 0 Then
       Txt5.Text = CDbl(Txt5.Text) + CantDiv
       Total = Total - CantDiv * 5
    Else
      Txt5.Text = 0
    End If
End If
'reviso si hay monedas de 1
If Total >= 1 Then
    CantDiv = Int(Total / 1)
    If CantDiv > 0 Then
       Txt1.Text = CDbl(Txt1.Text) + CantDiv
       Total = Total - CantDiv * 1
    Else
       Txt1.Text = 0
    End If
End If
'conversion de decimales
Decimales = Decimales * 100


'reviso si hay decimales de 50
If Decimales >= 50 Then
    CantDiv = Int(Format(Decimales, "##.#0") / 50)
    If CantDiv > 0 Then
       TxtD50.Text = CDbl(TxtD50.Text) + CantDiv
       Decimales = Format(Decimales, "##.#0") - CantDiv * 50
    Else
       TxtD50.Text = 0
    End If
End If

'reviso si hay decimales de 25
If Decimales >= 25 Then
    CantDiv = Int(Decimales / 25)
    If CantDiv > 0 Then
       TxtD25.Text = CDbl(TxtD25.Text) + CantDiv
       Decimales = Decimales - CantDiv * 25
    Else
       TxtD25.Text = 0
    End If
End If

'reviso si hay decimales de 10
If Decimales >= 10 Then
    CantDiv = Int(Format(Decimales, "#.#0") / 10)
    If CantDiv > 0 Then
       TxtD10.Text = CDbl(TxtD10.Text) + CantDiv
       Decimales = Format(Decimales, "#.#0") - CantDiv * 10
    Else
       TxtD10.Text = 0
    End If
End If

'reviso si hay decimales de 5
If Decimales >= 5 Then
    CantDiv = Int(Format(Decimales, "#.#0") / 5)
    If CantDiv > 0 Then
       TxtD05.Text = CDbl(TxtD05.Text) + CantDiv
       Decimales = Format(Decimales, "#.#0") - CantDiv * 5
    Else
       TxtD05.Text = 0
    End If
End If

'reviso si hay decimales de 1
If Decimales >= 1 Then
    CantDiv = Int(Format(Decimales, "#.#0") / 1)   'Int(Format(Decimales, "#.#0") / 1)
    If CantDiv > 0 Then
       TxtD01.Text = CDbl(TxtD01.Text) + CantDiv
       Decimales = Format(Decimales, "#.#0") - CantDiv * 1
    Else
       TxtD01.Text = 0
    End If
End If

If Decimales > 0 Then
 If Decimales < 1 Then
    Ajuste = Ajuste + (Decimales / 100)

 End If
End If


Exit Sub
TipoErr:
ControlErrores


End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim SqlString As String
Dim FechaNomina As String
Dim rpt As Object
Dim fPreview As New FrmPreview


 SqlString = "SELECT  * From Nomina Where (NumNomina = " & NumNomina & ")"
                      
 Me.AdoBusca.RecordSource = SqlString
 Me.AdoBusca.Refresh
 If Not Me.AdoBusca.Recordset.EOF Then
  FechaNomina = "Desde: " & Format(Me.AdoBusca.Recordset("FechaNominaINI"), "Long Date") & "        Hasta: " & Format(Me.AdoBusca.Recordset("FechaNomina"), "Long Date")
 
 End If
ARDenominaciones.LblFechaNomina = FechaNomina

           fPreview.arv.ReportSource = ARDenominaciones
           fPreview.Show 1
'ARDenominaciones.Show 1
End Sub

Private Sub Form_Load()
'On Error GoTo TipoErr
Dim SqlDetalleNomina As String
Dim SubTotal As Double
Dim Total As Double




Ajuste = 0
SubTotal = 0
Total = 0

With Me.DtaDetalleNomina

   .ConnectionString = Conexion
End With

With Me.AdoBusca
   .ConnectionString = Conexion
End With


Me.Caption = Me.Caption + " " + Str(NumNomina)
Label1.Caption = Label1.Caption + " " + Str(NumNomina)
'SQlDetalleNomina = "SELECT DetalleNomina.* From DetalleNomina WHERE DetalleNomina.NumNomina=  " & NumNomina & ""
If Quien = "Nomina" Then
        SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo," & vbLf
        SQlReportes = SQlReportes & "                 Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo," & vbLf
        SQlReportes = SQlReportes & "                  Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal," & vbLf
        SQlReportes = SQlReportes & "                      Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada," & vbLf
        SQlReportes = SQlReportes & "                      DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo," & vbLf
        SQlReportes = SQlReportes & "                      Cargo.Cargo, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones," & vbLf
        SQlReportes = SQlReportes & "                      DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones, DetalleNomina.DiasVacaciones," & vbLf
        SQlReportes = SQlReportes & "                      DetalleNomina.VacacionesPagadas, DetalleNomina.BonoProduccion, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones," & vbLf
        SQlReportes = SQlReportes & "                      DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal, DetalleNomina.Mes13," & vbLf
        SQlReportes = SQlReportes & "                        DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo," & vbLf
        SQlReportes = SQlReportes & "                       Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE," & vbLf
        SQlReportes = SQlReportes & "                       DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
        SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion  AS TotalDevengado," & vbLf
        SQlReportes = SQlReportes & "                       DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir," & vbLf
        SQlReportes = SQlReportes & "                       (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +" & vbLf
        SQlReportes = SQlReportes & "                        DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas+ DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion )" & vbLf
        SQlReportes = SQlReportes & "                       - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar," & vbLf
        SQlReportes = SQlReportes & "                      DetalleNomina.TarifaHoraria,DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion,Empleado.CodEmpleado1" & vbLf
        SQlReportes = SQlReportes & " FROM         Nomina INNER JOIN" & vbLf
        SQlReportes = SQlReportes & "                       Grupo INNER JOIN" & vbLf
        SQlReportes = SQlReportes & "                       Cargo INNER JOIN" & vbLf
        SQlReportes = SQlReportes & "                       TipoNomina INNER JOIN" & vbLf
        SQlReportes = SQlReportes & "                       Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN" & vbLf
        SQlReportes = SQlReportes & "                       DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON" & vbLf
        SQlReportes = SQlReportes & "                       TipoNomina.CodTipoNomina = Nomina.CodTipoNomina And Nomina.NumNomina = DetalleNomina.NumNomina" & vbLf
        SQlReportes = SQlReportes & " WHERE     (Nomina.NumNomina = " & NumNomina & ") AND  (DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion <> 0) " & vbLf
        SQlReportes = SQlReportes & " ORDER BY Nomina.NumNomina, Empleado.CodEmpleado1" & vbLf
ElseIf Quien = "MonedasDepartamento" Then
          SQlReportes = "SELECT     Nomina.NumNomina, TipoNomina.Nomina, TipoNomina.Periodo, Nomina.CodTipoNomina, Nomina.TotalSalarioBasico, Nomina.TotalDestajo, Nomina.TotalHorasExtras, Nomina.TotalComisiones, Nomina.TotalIncentivos, Nomina.TotalDeducciones, Nomina.TotalPrestamo, Nomina.TotalMontoINSS, Nomina.TotalMontoIR, Nomina.TotalOtrosIngresos, Nomina.TotalVacaciones, Nomina.TotalINSSPatronal,  Nomina.TotalIRPatronal, Nomina.Totalmes13, Nomina.FechaNomina, Nomina.Activa, Nomina.Procesada, Nomina.Cerrada, DetalleNomina.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2 AS Nombre, Cargo.CodCargo,  Cargo.Cargo, DetalleNomina.BonoProduccion, DetalleNomina.SalarioBasico, DetalleNomina.Destajo, DetalleNomina.HorasExtras, DetalleNomina.Comisiones, DetalleNomina.DiasDescuento, DetalleNomina.Adelantos, DetalleNomina.Incentivos, DetalleNomina.Deducciones,  " & _
                      "DetalleNomina.DiasVacaciones, DetalleNomina.VacacionesPagadas, DetalleNomina.Prestamo, DetalleNomina.MontoINSS, DetalleNomina.MontoIR, DetalleNomina.Vacaciones, DetalleNomina.OtrosIngresos, DetalleNomina.DescripOtrIngre, DetalleNomina.INSSPatronal, DetalleNomina.IRPatronal,  DetalleNomina.Mes13, DetalleNomina.TotalSubsidio, Empleado.CodGrupo, Empleado.DescripOtrIngre AS Expr1, Grupo.Grupo,  Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomina.HE, DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion  AS TotalDevengado, DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones AS TotalDeducir,  " & _
                      "(DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos +  DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia + DetalleNomina.IncetivoProduccion + DetalleNomina.BonoProduccion ) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) AS NetoPagar, DetalleNomina.TarifaHoraria, DetalleNomina.HTrabajada, DetalleNomina.SeptimoDia, DetalleNomina.IncetivoProduccion, Empleado.CodEmpleado1, departamento.departamento , departamento.CodDepartamento,Nomina.FechaNominaINI  " & _
                      "FROM  Nomina INNER JOIN  Grupo INNER JOIN  Cargo INNER JOIN  TipoNomina INNER JOIN  Empleado ON TipoNomina.CodTipoNomina = Empleado.CodTipoNomina ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN  DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado ON Grupo.CodGrupo = Empleado.CodGrupo ON  TipoNomina.CodTipoNomina = Nomina.CodTipoNomina AND Nomina.NumNomina = DetalleNomina.NumNomina INNER JOIN  Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento " & _
                      "WHERE     (Nomina.NumNomina = " & NumNomina & ") AND ((DetalleNomina.SalarioBasico + DetalleNomina.Comisiones + DetalleNomina.Incentivos + DetalleNomina.HorasExtras + DetalleNomina.OtrosIngresos + DetalleNomina.Destajo + DetalleNomina.VacacionesPagadas + DetalleNomina.SeptimoDia) - (DetalleNomina.Prestamo + DetalleNomina.MontoINSS + DetalleNomina.MontoIR + DetalleNomina.Deducciones) <> 0) AND (Departamento.CodDepartamento = '" & CodigoDepartamento & "') " & _
                      "ORDER BY Empleado.CodGrupo, Empleado.CodEmpleado "



End If

DtaDetalleNomina.RecordSource = SQlReportes
DtaDetalleNomina.Refresh
DtaDetalleNomina.Recordset.MoveLast

PBMonedas.Min = 0
PBMonedas.Max = DtaDetalleNomina.Recordset.RecordCount
PBMonedas.Value = 0

DtaDetalleNomina.Refresh
Do While Not DtaDetalleNomina.Recordset.EOF
    PBMonedas.Value = PBMonedas.Value + 1
    SubTotal = Format(DtaDetalleNomina.Recordset("NetoPagar"), "##,##0.0000")
'    SubTotal = DtadetalleNomina.Recordset("SalarioBasico")
'    SubTotal = SubTotal + DtadetalleNomina.Recordset("destajo")
'    SubTotal = SubTotal + DtadetalleNomina.Recordset("IncetivoProduccion")
'    SubTotal = SubTotal + DtadetalleNomina.Recordset("HorasExtras")
'    SubTotal = SubTotal + DtadetalleNomina.Recordset("Comisiones")
'    SubTotal = SubTotal + DtadetalleNomina.Recordset("incentivos")
'    SubTotal = SubTotal + DtadetalleNomina.Recordset("OtrosIngresos")
'    SubTotal = SubTotal + DtadetalleNomina.Recordset("SeptimoDia")
'    SubTotal = SubTotal - DtadetalleNomina.Recordset("Deducciones")
'    SubTotal = SubTotal - DtadetalleNomina.Recordset("Prestamo")
'    SubTotal = SubTotal - DtadetalleNomina.Recordset("MontoInss")
'    SubTotal = SubTotal - DtadetalleNomina.Recordset("MontoIR")
'    If Not IsNull(DtaDetalleNomina.Recordset("TotalSubsidio")) Then
'      SubTotal = SubTotal + DtaDetalleNomina.Recordset("TotalSubsidio")
'
'    End If
    Me.txtTotal.Text = Format(SubTotal, "#,###,##.##0")
    CmdOtorga.Value = True
   ' MsgBox TxtTotal.Text
   
'   If CDbl(SubTotal) > 1000 Then
'     Total = CDbl(Format(Total, "#,###.##0")) + CDbl(Format(SubTotal, "#,###.##0"))
'   Else
'     Total = CDbl(Format(Total, "#,###.##0")) + CDbl(Format(SubTotal, "###.##0"))
'   End If
   
   
   Total = Total + SubTotal
   
    DtaDetalleNomina.Recordset.MoveNext
Loop


If Ajuste > 0 Then
'    Me.txtTotal.Text = Format(Ajuste, "#,###,##.#0")
'    CmdOtorga.Value = True
End If




'-------------------------TOTAL NOMINA
Me.AdoBusca.RecordSource = "SELECT SUM(SalarioBasico) AS SalarioBasico, SUM(Destajo) AS Produccion, SUM(HorasExtras) AS HorasExtra, SUM(Comisiones) AS Puntualidad, SUM(VacacionesPagadas) AS Vacaciones, SUM(SeptimoDia) AS SeptimoDia, SUM(IncetivoProduccion) AS IncentivosProduccion, SUM(Incentivos) AS Antiguedad, SUM(OtrosIngresos) AS OtrosIngresos, SUM(SalarioBasico) + SUM(Destajo) + SUM(HorasExtras) + SUM(Comisiones) + SUM(SeptimoDia) + SUM(IncetivoProduccion + Incentivos) + SUM(OtrosIngresos) + SUM(VacacionesPagadas) + SUM(BonoProduccion) AS TotalDevengado, SUM(Deducciones) AS Deducciones, SUM(Prestamo) AS Prestamo, SUM(MontoINSS) AS MontoInss, SUM(MontoIR) AS MontoIr, SUM(DiasDescuento) AS DiasDescuento, SUM(Adelantos) AS Adelantos, SUM(Deducciones) + SUM(Prestamo) + SUM(MontoINSS) + SUM(MontoIR) + SUM(DiasDescuento) + SUM(Adelantos) AS TotalDeduccines, (SUM(SalarioBasico) + SUM(Destajo) " & _
                           " + SUM(HorasExtras) + SUM(Comisiones) + SUM(SeptimoDia) + SUM(IncetivoProduccion + Incentivos) + SUM(OtrosIngresos) + SUM(BonoProduccion) + SUM(VacacionesPagadas) ) - (SUM(Deducciones) + SUM(Prestamo) + SUM(MontoINSS) + SUM(MontoIR) + SUM(DiasDescuento) + SUM(Adelantos)) AS Neto, SUM(INSSPatronal) AS InssPatronal, SUM(IRPatronal) AS IrPatronal, SUM(INATEC) AS Inatec, NumNomina, SUM(HE) AS HE, SUM(HTrabajada) AS HTrabajada, SUM(INSSPatronal) + SUM(INATEC) AS TotalObligaciones, SUM(BonoProduccion) AS BonoProduccion From DetalleNomina GROUP BY NumNomina Having (NumNomina = " & NumNomina & ")"
Me.AdoBusca.Refresh
If Not Me.AdoBusca.Recordset.EOF Then
 Total = Me.AdoBusca.Recordset("Neto")
End If

TxtGranTotal.Text = Format(Total, "##,##0.00")

Total = CDbl(Me.TxtTot1000.Text) + CDbl(Me.TxtTot500.Text) + CDbl(Me.TxtTot200.Text) + CDbl(Me.TxtTot100.Text) + CDbl(Me.TxtTot50.Text) + CDbl(Me.TxtTot20.Text) + CDbl(Me.TxtTot10.Text) + CDbl(Me.TxtTot5.Text) + CDbl(Me.TxtTot1.Text) + CDbl(Me.TxtTotD50.Text) + CDbl(Me.TxtTotD25.Text) + CDbl(Me.TxtTotD10.Text) + CDbl(Me.TxtTotD05.Text) + CDbl(Me.TxtTotD01.Text)
Ajuste = CDbl(Me.TxtGranTotal.Text) - Total

  Me.txtTotal.Text = Format(Ajuste, "#,###,##.##0")
  CmdOtorga.Value = True

Exit Sub
TipoErr:
ControlErrores

End Sub

Private Sub Text1_Change()
TxtTot500 = Str(val(Txt500.Text) * 500)
TxtTot500 = Format(val(TxtTot500.Text), "###,##0.##")
End Sub

Private Sub Txt1_Change()
TxtTot1 = Str(val(Txt1.Text) * 1)
TxtTot1 = Format(val(TxtTot1.Text), "###,##0.##")
End Sub

Private Sub Txt10_Change()
TxtTot10 = Str(val(Txt10.Text) * 10)
TxtTot10 = Format(val(TxtTot10.Text), "###,##0.##")
End Sub

Private Sub Txt1000_Change()
TxtTot1000 = Str(val(Txt1000.Text) * 1000)
TxtTot1000 = Format(val(TxtTot1000.Text), "###,##0.##")
End Sub

Private Sub Txt200_Change()
TxtTot200 = Str(val(Txt200.Text) * 200)
TxtTot200 = Format(val(TxtTot200.Text), "###,##0.##")
End Sub

Private Sub Txt500_Change()
TxtTot500 = Str(val(Txt500.Text) * 500)
TxtTot500 = Format(val(TxtTot500.Text), "###,##0.##")
End Sub



Private Sub Txt100_Change()
TxtTot100 = Str(val(Txt100.Text) * 100)
TxtTot100 = Format(val(TxtTot100.Text), "###,##0.##")
End Sub

Private Sub Txt20_Change()
TxtTot20 = Str(val(Txt20.Text) * 20)
TxtTot20 = Format(val(TxtTot20.Text), "###,##0.##")
End Sub

Private Sub Txt5_Change()
TxtTot5 = Str(val(Txt5.Text) * 5)
TxtTot5 = Format(val(TxtTot5.Text), "###,##0.##")
End Sub

Private Sub Txt50_Change()
TxtTot50 = Str(val(Txt50.Text) * 50)
TxtTot50 = Format(val(TxtTot50.Text), "###,##0.##")
End Sub

Private Sub TxtD01_Change()
TxtTotD01 = Str(val(TxtD01.Text) * 0.01)
TxtTotD01 = Format(TxtTotD01, "#0.#0")
End Sub

Private Sub TxtD05_Change()
TxtTotD05 = Str(val(TxtD05.Text) * 0.05)
TxtTotD05 = Format(TxtTotD05, "#0.#0")
End Sub

Private Sub TxtD10_Change()
TxtTotD10 = Str(val(TxtD10.Text) * 0.1)
TxtTotD10 = Format(TxtTotD10, "#0.##0")
End Sub

Private Sub TxtD25_Change()
TxtTotD25 = Str(val(TxtD25.Text) * 0.25)
TxtTotD25 = Format(TxtTotD25, "#0.##")
End Sub

Private Sub TxtD50_Change()
TxtTotD50 = Str(val(TxtD50.Text) * 0.5)
TxtTotD50 = Format(TxtTotD50, "#0.##")
End Sub

Private Sub TxtTotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim CantDiv As Integer
Dim Total As Double
Dim Enteros As Integer
Dim Decimales As Double

Total = val(txtTotal.Text)
Enteros = Int(Total)
Decimales = Format(Total - Enteros, "#0.##")

'reviso si hay billeres de 1000
If Total >= 1000 Then
    CantDiv = Int(Total / 1000)
    If CantDiv > 0 Then
       Txt1000.Text = val(Txt1000.Text) + CantDiv
       'MsgBox Total
       'MsgBox (CantDiv * 100)
       Total = Total - (CantDiv * 1000)
    Else
       Txt1000.Text = 0
    End If
End If

'reviso si hay billeres de 100
If Total >= 500 Then
    CantDiv = Int(Total / 500)
    If CantDiv > 0 Then
       Txt100.Text = CantDiv
       Total = Total - CantDiv * 500
    Else
       Txt100.Text = 0
    End If
End If

'reviso si hay billeres de 100
If Total >= 100 Then
    CantDiv = Int(Total / 100)
    If CantDiv > 0 Then
       Txt100.Text = CantDiv
       Total = Total - CantDiv * 100
    Else
       Txt100.Text = 0
    End If
End If
'reviso si hay billeres de 50
If Total >= 50 Then
    CantDiv = Int(Total / 50)
    If CantDiv > 0 Then
       Txt50.Text = CantDiv
       Total = Total - CantDiv * 50
    Else
       Txt50.Text = 0
    End If
End If
'reviso si hay billeres de 20
If Total >= 20 Then
    CantDiv = Int(Total / 20)
    If CantDiv > 0 Then
       Txt20.Text = CantDiv
       Total = Total - CantDiv * 20
    Else
       Txt20.Text = 0
    End If
End If

'reviso si hay billeres de 10
If Total >= 10 Then
    CantDiv = Int(Total / 10)
    If CantDiv > 0 Then
       Txt10.Text = CantDiv
       Total = Total - CantDiv * 10
    Else
      Txt10.Text = 0
    End If
End If
'reviso si hay monedas de 5
If Total >= 5 Then
    CantDiv = Int(Total / 5)
    If CantDiv > 0 Then
       Txt5.Text = CantDiv
       Total = Total - CantDiv * 5
    Else
      Txt5.Text = 0
    End If
End If
'reviso si hay monedas de 1
If Total >= 1 Then
    CantDiv = Int(Total / 1)
    If CantDiv > 0 Then
       Txt1.Text = CantDiv
       Total = Total - CantDiv * 1
    Else
       Txt1.Text = 0
    End If
End If
'conversion de decimales
Decimales = Decimales * 100
'reviso si hay decimales de 50
If Decimales >= 50 Then
    CantDiv = Int(Decimales / 50)
    If CantDiv > 0 Then
       TxtD50.Text = CantDiv
       Decimales = Decimales - CantDiv * 50
    Else
       TxtD50.Text = 0
    End If
End If

'reviso si hay decimales de 25
If Decimales >= 25 Then
    CantDiv = Int(Decimales / 25)
    If CantDiv > 0 Then
       TxtD25.Text = CantDiv
       Decimales = Decimales - CantDiv * 25
    Else
       TxtD25.Text = 0
    End If
End If

'reviso si hay decimales de 10
If Decimales >= 10 Then
    CantDiv = Int(Decimales / 10)
    If CantDiv > 0 Then
       TxtD10.Text = CantDiv
       Decimales = Decimales - CantDiv * 10
    Else
       TxtD10.Text = 0
    End If
End If

'reviso si hay decimales de 5
If Decimales >= 5 Then
    CantDiv = Int(Decimales / 5)
    If CantDiv > 0 Then
       TxtD05.Text = CantDiv
       Decimales = Decimales - CantDiv * 5
    Else
       TxtD05.Text = 0
    End If
End If

'reviso si hay decimales de 1
If Decimales >= 1 Then
    CantDiv = Int(Decimales / 1)
    If CantDiv > 0 Then
       TxtD01.Text = CantDiv
       Decimales = Decimales - CantDiv * 1
    Else
       TxtD01.Text = 0
    End If
End If

End If 'del ascii enter
End Sub
