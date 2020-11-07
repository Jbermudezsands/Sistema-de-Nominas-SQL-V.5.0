VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmMonedas13vo 
   Caption         =   "Denominacion de 13vo y Vacaciones"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8625
   LinkTopic       =   "Form2"
   ScaleHeight     =   6585
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txt1000 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   61
      Text            =   "0"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox TxtTot1000 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   60
      Text            =   "0.00"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Txt200 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   56
      Text            =   "0"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox TxtTot200 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   "0.00"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      DownPicture     =   "FrmMonedas13vo.frx":0000
      Height          =   375
      Left            =   5040
      Picture         =   "FrmMonedas13vo.frx":2E02
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton CmdSalir 
      DownPicture     =   "FrmMonedas13vo.frx":53C4
      Height          =   375
      Left            =   7080
      Picture         =   "FrmMonedas13vo.frx":6EA6
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox Txt100 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "0"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Txt50 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "0"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Txt20 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "0"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Txt10 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "0"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Txt5 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "0"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Txt1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox TxtD50 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox TxtD25 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox TxtD10 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox TxtD05 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "0"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox TxtD01 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox TxtTot100 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0.00"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox TxtTot50 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox TxtTot20 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox TxtTot10 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0.00"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox TxtTot5 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox TxtTot1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox TxtTotD50 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox TxtTotD25 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox TxtTotD10 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox TxtTotD05 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox TxtTotD01 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   4920
      Width           =   1095
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
      Left            =   2400
      TabIndex        =   4
      Top             =   720
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
      Left            =   6360
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton CmdOtorga 
      Caption         =   "Otorgar"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox TxtTot500 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Txt500 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0"
      Top             =   2520
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc DtaDetalleNomina 
      Height          =   495
      Left            =   1560
      Top             =   7560
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
   Begin MSAdodcLib.Adodc DtaControles 
      Height          =   495
      Left            =   1560
      Top             =   6960
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
   Begin XtremeSuiteControls.ProgressBar PBMonedas 
      Height          =   375
      Left            =   120
      TabIndex        =   59
      Top             =   5520
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
      Left            =   240
      TabIndex        =   63
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label29 
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
      Left            =   2520
      TabIndex        =   62
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label28 
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
      Left            =   240
      TabIndex        =   58
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label27 
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
      Left            =   2520
      TabIndex        =   57
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   240
      Picture         =   "FrmMonedas13vo.frx":8988
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1845
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
      Left            =   2280
      TabIndex        =   52
      Top             =   0
      Width           =   5895
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
      Left            =   240
      TabIndex        =   51
      Top             =   3480
      Width           =   975
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
      Left            =   0
      TabIndex        =   50
      Top             =   3960
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
      Left            =   0
      TabIndex        =   49
      Top             =   4440
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
      Left            =   0
      TabIndex        =   48
      Top             =   4920
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
      Left            =   4200
      TabIndex        =   47
      Top             =   2040
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
      Left            =   4200
      TabIndex        =   46
      Top             =   2520
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
      Left            =   4200
      TabIndex        =   45
      Top             =   3000
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
      Left            =   4200
      TabIndex        =   44
      Top             =   3480
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
      Left            =   4200
      TabIndex        =   43
      Top             =   3960
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
      Left            =   4200
      TabIndex        =   42
      Top             =   4440
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
      Left            =   4200
      TabIndex        =   41
      Top             =   4920
      Width           =   1215
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
      Left            =   5040
      TabIndex        =   40
      Top             =   1080
      Width           =   1215
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
      Left            =   2520
      TabIndex        =   39
      Top             =   3480
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
      Left            =   2520
      TabIndex        =   38
      Top             =   3960
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
      Left            =   2520
      TabIndex        =   37
      Top             =   4440
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
      Left            =   2520
      TabIndex        =   36
      Top             =   4920
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
      Left            =   6720
      TabIndex        =   35
      Top             =   2040
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
      Left            =   6720
      TabIndex        =   34
      Top             =   2520
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
      Left            =   6720
      TabIndex        =   33
      Top             =   3000
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
      Left            =   6720
      TabIndex        =   32
      Top             =   3480
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
      Left            =   6720
      TabIndex        =   31
      Top             =   3960
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
      Left            =   6720
      TabIndex        =   30
      Top             =   4440
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
      Left            =   6720
      TabIndex        =   29
      Top             =   4920
      Width           =   375
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
      Left            =   2520
      TabIndex        =   28
      Top             =   2520
      Width           =   375
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
      Left            =   240
      TabIndex        =   27
      Top             =   2520
      Width           =   975
   End
End
Attribute VB_Name = "FrmMonedas13vo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOtorga_Click()
On Error GoTo TipoErr
Dim CantDiv As Double
Dim Total As Double


Dim Enteros As Long
Dim Decimales As Double

Total = Format((TxtTotal.Text), "##,##0.00")

'Total = Me.txtGranTotal.Text
Enteros = Int(Total)
Decimales = Format(Total - Enteros, "#0.00000")

'reviso si hay billeres de 500
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

'reviso si hay billeres de 200
If Total >= 200 Then
    CantDiv = Int(Total / 200)
    If CantDiv > 0 Then
       Txt200.Text = val(Txt200.Text) + CantDiv
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
    CantDiv = Int(Decimales / 50)
    If CantDiv > 0 Then
       TxtD50.Text = CDbl(TxtD50.Text) + CantDiv
       Decimales = Decimales - CantDiv * 50
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
    CantDiv = Int(Decimales / 10)
    If CantDiv > 0 Then
       TxtD10.Text = CDbl(TxtD10.Text) + CantDiv
       Decimales = Decimales - CantDiv * 10
    Else
       TxtD10.Text = 0
    End If
End If

'reviso si hay decimales de 5
If Decimales >= 5 Then
    CantDiv = Int(Decimales / 5)
    If CantDiv > 0 Then
       TxtD05.Text = CDbl(TxtD05.Text) + CantDiv
       Decimales = Decimales - CantDiv * 5
    Else
       TxtD05.Text = 0
    End If
End If

'reviso si hay decimales de 1
If Decimales >= 1 Then
    CantDiv = Int(Decimales / 1)
    If CantDiv > 0 Then
       TxtD01.Text = CDbl(TxtD01.Text) + CantDiv
       Decimales = Decimales - CantDiv * 1
    Else
       TxtD01.Text = 0
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
Dim rpt As Object
Dim fPreview As New FrmPreview

Quien = "Nomina13vo"
'ARDenominaciones.Show 1

           fPreview.arv.ReportSource = ARDenominaciones
           fPreview.Show 1
End Sub

Private Sub Form_Load()
On Error GoTo TipoErr
Dim SqlDetalleNomina As String
Dim SubTotal As Double
Dim Total As Double, DiaMes As Double
With Me.DtaControles
  .ConnectionString = Conexion
  .RecordSource = "Controles"
  .Refresh
End With
DtaControles.Refresh
DiaMes = DtaControles.Recordset("DiasMes")

SubTotal = 0
Total = 0

With Me.DtaDetalleNomina
   .ConnectionString = Conexion
End With

Select Case Quien

Case "Vacaciones"

            Frm13Vaca.DtaTipoNomina.Refresh
            Do While Not Frm13Vaca.DtaTipoNomina.Recordset.EOF
                If Frm13Vaca.DtaTipoNomina.Recordset("nomina") = Frm13Vaca.DBCNominas.Text Then
                   CodTipoNomina = Frm13Vaca.DtaTipoNomina.Recordset("CodTipoNomina")
                   Exit Do
                End If
                Frm13Vaca.DtaTipoNomina.Recordset.MoveNext
            Loop
            
            NumNomina = Frm13Vaca.TxtNumNomVaca.Text
            Me.Caption = Me.Caption + " " + Str(NumNomina)
            Label1.Caption = Label1.Caption + " " + Str(NumNomina)
            
            
            If Frm13Vaca.DBCNominas.Text <> "Administracion" Then
                  
                If Frm13Vaca.ChkRestar.Value = 1 Then
                      SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir  AS MontoPagar,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir AS NetoPagar, DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                                    "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
                Else
                     SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir AS NetoPagar, DetalleNomVaca.TotalDevengado,Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                                   "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND ( DetalleNomVaca.DiasAPagar <> 0)ORDER BY Empleado.CodEmpleado1 "
                End If
            Else
            
                If Frm13Vaca.ChkRestar.Value = 1 Then
                    SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir AS MontoPagar,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir AS NetoPagar, DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                               "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1)AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
                
                Else
                      SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir AS NetoPagar, DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                                    "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
                End If
            
            End If
            
            '  SQLReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.SalarioMensual * (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento)/ 30 - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoAPagar,DetalleNomVaca.SalarioMensual * (DetalleNomVaca.DiasAPagar - DetalleNomVaca.DiasDescuento)/ 30 - DetalleNomVaca.AdelantoVacaciones AS TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
            '               "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomVaca & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "


Case Else
NumNomina = Frm13Vaca.TxtNumNom13.Text
Me.Caption = Me.Caption + " " + Str(NumNomina)
Label1.Caption = Label1.Caption + " " + Str(NumNomina)


'SQlReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, DetalleNom13Mes.SalarioMensual * DetalleNom13Mes.DiasAPagar / " & DiaMes & " - DetalleNom13Mes.Adelanto13vo AS MontoPagar, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Cargo.Cargo, DetalleNom13Mes.SalarioMensual * DetalleNom13Mes.DiasAPagar / " & DiaMes & " AS TotalDevengado, Empleado.CodEmpleado1, Empleado.TarifaHoraria, Historico.FechaContrato FROM Nom13Mes INNER JOIN Cargo INNER JOIN " & _
              "Empleado ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado " & _
              "Where (Nom13Mes.NumNom13Mes = " & NumNomina & ") ORDER BY Empleado.CodEmpleado1"
 SQlReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo AS MontoPagar, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Cargo.Cargo, DetalleNom13Mes.SalarioAPagar AS TotalDevengado, Empleado.CodEmpleado1 FROM  Nom13Mes INNER JOIN  Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes  " & _
               "Where (Nom13Mes.NumNom13Mes = " & NumNomina & ") ORDER BY Empleado.CodEmpleado1"
End Select
DtaDetalleNomina.RecordSource = SQlReportes
DtaDetalleNomina.Refresh
DtaDetalleNomina.Recordset.MoveLast

PBMonedas.Min = 0
PBMonedas.Max = DtaDetalleNomina.Recordset.RecordCount
PBMonedas.Value = 0

DtaDetalleNomina.Refresh
Do While Not DtaDetalleNomina.Recordset.EOF
    PBMonedas.Value = PBMonedas.Value + 1
    SubTotal = DtaDetalleNomina.Recordset("MontoPagar")

    Me.TxtTotal.Text = Format(SubTotal, "#####0.00")
    CmdOtorga.Value = True
    Total = CDbl(Total) + CDbl(SubTotal)
    DtaDetalleNomina.Recordset.MoveNext
Loop

TxtGranTotal.Text = Format(Total, "###,##0.00")

Exit Sub
TipoErr:
ControlErrores

End Sub

Private Sub Text2_Change()

End Sub

Private Sub Txt1_Change()
TxtTot1 = Str(val(Txt1.Text) * 1)
TxtTot1 = Format(val(TxtTot1.Text), "###,##0.00")
End Sub

Private Sub Txt10_Change()
TxtTot10 = Str(val(Txt10.Text) * 10)
TxtTot10 = Format(val(TxtTot10.Text), "###,##0.00")
End Sub

Private Sub Txt1000_Change()
TxtTot1000 = Str(val(Txt1000.Text) * 1000)
TxtTot1000 = Format(val(TxtTot1000.Text), "###,##0.00")
End Sub

Private Sub Txt200_Change()
TxtTot200 = Str(val(Txt200.Text) * 200)
TxtTot200 = Format(val(TxtTot200.Text), "###,##0.00")
End Sub

Private Sub Txt500_Change()
TxtTot500 = Str(val(Txt500.Text) * 500)
TxtTot500 = Format(val(TxtTot500.Text), "###,##0.00")
End Sub



Private Sub Txt100_Change()
TxtTot100 = Str(val(Txt100.Text) * 100)
TxtTot100 = Format(val(TxtTot100.Text), "###,##0.00")
End Sub

Private Sub Txt20_Change()
TxtTot20 = Str(val(Txt20.Text) * 20)
TxtTot20 = Format(val(TxtTot20.Text), "###,##0.00")
End Sub

Private Sub Txt5_Change()
TxtTot5 = Str(val(Txt5.Text) * 5)
TxtTot5 = Format(val(TxtTot5.Text), "###,##0.00")
End Sub

Private Sub Txt50_Change()
TxtTot50 = Str(val(Txt50.Text) * 50)
TxtTot50 = Format(val(TxtTot50.Text), "###,##0.00")
End Sub

Private Sub TxtD01_Change()
TxtTotD01 = Str(val(TxtD01.Text) * 0.01)
TxtTotD01 = Format(TxtTotD01, "#0.00")
End Sub

Private Sub TxtD05_Change()
TxtTotD05 = Str(val(TxtD05.Text) * 0.05)
TxtTotD05 = Format(TxtTotD05, "#0.00")
End Sub

Private Sub TxtD10_Change()
TxtTotD10 = Str(val(TxtD10.Text) * 0.1)
TxtTotD10 = Format(TxtTotD10, "#0.00")
End Sub

Private Sub TxtD25_Change()
TxtTotD25 = Str(val(TxtD25.Text) * 0.25)
TxtTotD25 = Format(TxtTotD25, "#0.00")
End Sub

Private Sub TxtD50_Change()
TxtTotD50 = Str(val(TxtD50.Text) * 0.5)
TxtTotD50 = Format(TxtTotD50, "#0.00")
End Sub

Private Sub TxtTotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim CantDiv As Integer
Dim Total As Double
Dim Enteros As Integer
Dim Decimales As Double

Total = val(TxtTotal.Text)
Enteros = Int(Total)
Decimales = Format(Total - Enteros, "#0.00")

'reviso si hay billeres de 1000
If Total >= 1000 Then
    CantDiv = Int(Total / 1000)
    If CantDiv > 0 Then
       Txt1000.Text = CantDiv
       Total = Total - CantDiv * 1000
    Else
       Txt1000.Text = 0
    End If
End If


'reviso si hay billeres de 500
If Total >= 500 Then
    CantDiv = Int(Total / 500)
    If CantDiv > 0 Then
       Txt500.Text = CantDiv
       Total = Total - CantDiv * 500
    Else
       Txt500.Text = 0
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

