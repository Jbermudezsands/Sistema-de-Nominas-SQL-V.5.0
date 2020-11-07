VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmActivarNomina 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activar Nominas"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   427
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   593
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   8895
      TabIndex        =   12
      Top             =   0
      Width           =   8895
      Begin VB.Image Image2 
         Height          =   1020
         Left            =   0
         Picture         =   "FrmActivarNomina.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1290
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   8880
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "ACTIVACION DE NOMINAS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   360
         Width           =   4200
      End
   End
   Begin VB.TextBox TxtAno 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodos para la Activacion"
      Height          =   2175
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   8415
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   375
         Left            =   3960
         OleObjectBlob   =   "FrmActivarNomina.frx":0AFE
         TabIndex        =   11
         Top             =   1080
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "FrmActivarNomina.frx":0B88
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmActivarNomina.frx":0C02
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin MSDataListLib.DataCombo DBComboPeriodo 
         Bindings        =   "FrmActivarNomina.frx":0C6E
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker MtxtFechaini 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   16777215
         Format          =   17104897
         CurrentDate     =   37257
      End
      Begin MSComCtl2.DTPicker MtxtFecha 
         Height          =   375
         Left            =   5880
         TabIndex        =   5
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   16777215
         Format          =   17104897
         CurrentDate     =   37257
      End
      Begin VB.Label LblFechaLarga 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   720
         TabIndex        =   6
         Top             =   1560
         Width           =   7455
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "SALIR"
      DownPicture     =   "FrmActivarNomina.frx":0C88
      Height          =   375
      Left            =   7200
      Picture         =   "FrmActivarNomina.frx":276A
      TabIndex        =   1
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton CmdActivar 
      Caption         =   "ACTIVAR"
      DownPicture     =   "FrmActivarNomina.frx":424C
      Height          =   375
      Left            =   5760
      Picture         =   "FrmActivarNomina.frx":5D2E
      TabIndex        =   0
      Top             =   5760
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc DtaPeriodos 
      Height          =   375
      Left            =   360
      Top             =   7320
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
      Caption         =   "DtaPeriodos"
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
      Left            =   240
      Top             =   7320
      Width           =   3255
      _ExtentX        =   5741
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
   Begin MSAdodcLib.Adodc DtaNomina 
      Height          =   375
      Left            =   4680
      Top             =   7320
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
      Caption         =   "DtaNomina"
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
   Begin MSAdodcLib.Adodc DtaConsecutivos 
      Height          =   375
      Left            =   4800
      Top             =   7320
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
      Caption         =   "DtaConsecutivos"
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
   Begin MSAdodcLib.Adodc DtaTipoNomina 
      Height          =   375
      Left            =   360
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
      Caption         =   "DtaTipoNomina"
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
   Begin TrueOleDBGrid70.TDBGrid DbgrTipoNominas 
      Bindings        =   "FrmActivarNomina.frx":7630
      Height          =   1935
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3413
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
      Splits(0)._SavedRecordSelectors=   0   'False
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
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=160,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "FrmActivarNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdActivar_Click()
Dim NumIni As Long, Periodo As Integer
Dim NumFin As Long, Fecha1 As Long
Dim NumAIni As Long, Fecha2 As Long
Dim NumAFin As Long, NumeroNomina As Integer
Dim FechaTemp As Date
Dim SqlNominas As String, Mes As Integer
Dim CodTipoNomina As String
Dim Año As Integer
On Error GoTo TipoErrs

   
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   If Not Me.DBComboPeriodo.Text = "" Then
    Año = Year(MtxtFechaini)
    Periodo = Me.DBComboPeriodo.Text
   Else
     Exit Sub
   End If
'//////////////////////////////BUSCO LA NOMINA ACTIVA////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////
'   Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
   Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Periodo & ") AND (año = " & Año & ")"
   Me.DtaConsulta.Refresh
   If Not DtaConsulta.Recordset.EOF Then
      Año = Me.DtaConsulta.Recordset("año")
      Periodo = Me.DtaConsulta.Recordset("Periodo")
      Fecha1 = Me.DtaConsulta.Recordset("Inicio")
      Fecha2 = Me.DtaConsulta.Recordset("Final")
'/////////////////////////BUSCO SI LA NOMINA ACTIVA ACTUAL YA ESTA CERRADA////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////
'(FechaNominaINI = " & Fecha1 & ") AND (FechaNomina = " & Fecha2 & ") AND
     
      Me.DtaConsulta.RecordSource = "SELECT NumNomina, CodTipoNomina, FechaNominaINI, FechaNomina, Activa, Procesada, Cerrada From Nomina WHERE (Cerrada = 1) AND (CodTipoNomina = '" & CodTipoNomina & "') AND (FechaNomina BETWEEN CONVERT(DATETIME, '" & Format(Me.MtxtFechaini.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME,'" & Format(Me.MtxtFecha.Value, "yyyy-mm-dd") & "', 102))"
      Me.DtaConsulta.Refresh
'/////////SI NO ESTA CERRADA, DEJO LA ACTUAL///////////////////////////////////
      If Not DtaConsulta.Recordset.EOF Then
        MsgBox "La fecha de esta Nomina ya fue Activada, Consulte a su Soperte", vbCritical
        Exit Sub
      Else
'///////////////////SI ESTA CERRADA      ////////////////////////////////////////
'//////////BUSCO LA NOMINA ANTERIOR PARA DESACTIVARLA/////////////////
'////////////////////////////////////////////////////////////////////////
'AND (Periodo = " & Periodo & ")
            Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE  (año = " & Año & ") AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1) ORDER BY Periodo"
            Me.DtaConsulta.Refresh
            Do While Not Me.DtaConsulta.Recordset.EOF
                 Me.DtaConsulta.Recordset("Actual") = 0
                 Me.DtaConsulta.Recordset.Update
              Me.DtaConsulta.Recordset.MoveNext
            Loop
            
'//////////////////////////////////////////////////////////////////////////////////////
'////////////////// BUSCO LA NOMINA SIGUIENTE PARA ACTIVARLA/////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////
        Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE  (año = " & Año & ") AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Me.DBComboPeriodo.Text & ") ORDER BY Periodo"
        Me.DtaConsulta.Refresh
        If Not DtaConsulta.Recordset.EOF Then
            Me.DtaConsulta.Recordset("Actual") = 1
            Mes = Me.DtaConsulta.Recordset("mes")
            Me.DtaConsulta.Recordset.Update
            
'//////////BUSCO EL PERIODO FISCAL SIGUIENTE PARA ACTIVARLO/////////////////
'////////////////////////////////////////////////////////////////////////
       Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From PeriodoFiscal WHERE  (año = " & Año & ") AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Me.DBComboPeriodo.Text & ") ORDER BY Periodo"
        Me.DtaConsulta.Refresh
        If Not DtaConsulta.Recordset.EOF Then
            Me.DtaConsulta.Recordset("Actual") = 1
            Mes = Me.DtaConsulta.Recordset("mes")
            Me.DtaConsulta.Recordset.Update

'//////////BUSCO EL PERIODO FISCAL ANTERIOR PARA DESACTIVARLA/////////////////
'////////////////////////////////////////////////////////////////////////
            Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From PeriodoFiscal WHERE     (año = " & Año & ") AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Periodo & ") ORDER BY Periodo"
            Me.DtaConsulta.Refresh
            If Not DtaConsulta.Recordset.EOF Then
                 Me.DtaConsulta.Recordset("Actual") = 0
                 Me.DtaConsulta.Recordset.Update
            End If
         End If
            
            
            
            
            
        Else
            Cadena = "No se puede Asignar ningun Periodo, a esta Nomina" & vbLf
            Cadena = Cadena & "Necesitar Crear un nuevo Periodo,de los contrario" & vbLf
            Cadena = Cadena & "no podra activar la Nomina"
            MsgBox Cadena, vbCritical
            Exit Sub
        
        End If
     

      End If
    Else
     Exit Sub
    End If

If Not IsDate(MtxtFecha.Value) Then
   MsgBox "La fecha de Aplicación de la Nómina no es Correcta"
   Exit Sub
End If

If Not IsDate(MtxtFechaini.Value) Then
   MsgBox "La fecha de Aplicación inicial de la Nómina no es Correcta"
   Exit Sub
End If

FechaTemp = CDate(MtxtFechaini.Value)
NumIni = FechaTemp
FechaTemp = CDate(MtxtFecha.Value)
NumFin = FechaTemp

If NumIni <= nimfin Then
   MsgBox "Error en la digitacuión de las fechas de nóminas"
   Exit Sub
End If



'SqlNominas = "SELECT Nomina.Mes,Nomina.Ano,Nomina.NumNomina, Nomina.CodTipoNomina, Nomina.FechaNominaINI, Nomina.FechaNomina From Nomina WHERE (Nomina.CodTipoNomina='" & CodTipoNomina & "') AND (Nomina.FechaNominaINI>= " & NumIni & " And Nomina.FechaNominaINI<= " & NumFin & " OR Nomina.FechaNomina>=" & NumIni & " And Nomina.FechaNomina<=" & NumFin & ")"
SqlNominas = "SELECT Mes, Ano, NumNomina, CodTipoNomina, FechaNomina, FechaNominaINI From Nomina WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (FechaNomina = CONVERT(DATETIME, '" & Format(CDate(MtxtFecha.Value), "yyyy-mm-dd") & "', 102)) AND (FechaNominaINI = CONVERT(DATETIME, '" & Format(CDate(MtxtFechaini.Value), "yyyy-mm-dd") & "', 102))"
DtaNomina.RecordSource = SqlNominas
DtaNomina.Refresh

'If Not DtaNomina.Recordset.EOF Then
'   MsgBox "Existen nóminas que incluyen este período, no se puede activar esta nómina en el período especificado"
'   Exit Sub
'End If

DtaNomina.RecordSource = "Nomina"
DtaNomina.Refresh

DtaConsecutivos.Refresh




DtaNomina.Recordset.AddNew
NumeroNomina = DtaConsecutivos.Recordset("nominas")
DtaNomina.Recordset("NumNomina") = DtaConsecutivos.Recordset("nominas")
DtaNomina.Recordset("CodTipoNomina") = DtaTipoNomina.Recordset("CodTipoNomina")
DtaNomina.Recordset("FechaNomina") = Format(CDate(MtxtFecha.Value), "DD/MM/YYYY")
DtaNomina.Recordset("FechaNominaINI") = Format(CDate(MtxtFechaini.Value), "DD/MM/YYYY")
DtaNomina.Recordset("Activa") = 1
DtaNomina.Recordset("TotalSalarioBasico") = 0
DtaNomina.Recordset("TotalDestajo") = 0
DtaNomina.Recordset("TotalHorasExtras") = 0
DtaNomina.Recordset("TotalComisiones") = 0
DtaNomina.Recordset("TotalIncentivos") = 0
DtaNomina.Recordset("TotalDeducciones") = 0
DtaNomina.Recordset("TotalPrestamo") = 0
DtaNomina.Recordset("TotalMontoInss") = 0
DtaNomina.Recordset("TotalMontoIR") = 0
DtaNomina.Recordset("TotalVacaciones") = 0
DtaNomina.Recordset("TotalINSSPatronal") = 0
DtaNomina.Recordset("TotalIRPatronal") = 0
DtaNomina.Recordset("Anulada") = 0
DtaNomina.Recordset("Cerrada") = 0
DtaNomina.Recordset("Procesada") = 0
DtaNomina.Recordset("Mes") = 0
DtaNomina.Recordset("Ano") = Año
DtaNomina.Recordset.Update

'//////////BUSCO EL PERIODO FISCAL PARA ACTIVARLO/////////////////
'////////////////////////////////////////////////////////////////////////
        Me.DtaConsulta.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From PeriodoFiscal WHERE (Año = " & Año & ") AND (CodTipoNomina = " & CodTipoNomina & ") AND (Inicio = '" & Format(CDate(MtxtFechaini.Value), "DD/MM/YYYY") & "')ORDER BY Periodo"
        Me.DtaConsulta.Refresh
        If Not DtaConsulta.Recordset.EOF Then
            Me.DtaConsulta.Recordset("Actual") = 1
            DtaConsulta.Recordset("NumNomina") = DtaConsecutivos.Recordset("nominas")
'            Mes = Me.DtaConsulta.Recordset("mes")
            Me.DtaConsulta.Recordset.Update

         End If
         
'//////////ASIGNO EL NUMERO DE NOMINA A FECHA PLANILLA/////////////////
'////////////////////////////////////////////////////////////////////////
        Me.DtaConsulta.RecordSource = "SELECT Periodo, Año, Mes, CodTipoNomina, Inicio, Final, Actual,NumNomina From Fecha_Planilla WHERE (Año = " & Año & ") AND (CodTipoNomina = " & CodTipoNomina & ") AND (Inicio = '" & Format(CDate(MtxtFechaini.Value), "DD/MM/YYYY") & "')ORDER BY Periodo"
        Me.DtaConsulta.Refresh
        If Not DtaConsulta.Recordset.EOF Then
            DtaConsulta.Recordset("NumNomina") = DtaConsecutivos.Recordset("nominas")
           Me.DtaConsulta.Recordset.Update

         End If
         
'DtaTipoNomina.Recordset.Edit
DtaTipoNomina.Recordset("Activa") = 1
DtaTipoNomina.Recordset.Update

DtaConsecutivos.Refresh
'DtaConsecutivos.Recordset.Edit
DtaConsecutivos.Recordset("nominas") = DtaConsecutivos.Recordset("nominas") + 1
DtaConsecutivos.Recordset.Update

'Periodo = Me.DBComboPeriodo.Text
'Me.DtaConsulta.RecordSource = "SELECT Periodo, año, CodTipoNomina, mes, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina ='" & CodTipoNomina & "') AND (Actual = 1)"
'Me.DtaConsulta.Refresh
'If Not DtaConsulta.Recordset.EOF Then
'     DtaConsulta.Recordset("Actual") = 0
'   Me.DtaConsulta.Recordset.Update
'End If
'
'Me.DtaConsulta.RecordSource = "SELECT mes,Periodo, año, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo <= " & Periodo & ") AND (Calculada = 0)"
'Me.DtaConsulta.Refresh
'Do While Not DtaConsulta.Recordset.EOF
'   Me.DtaConsulta.Recordset("Calculada") = 1
'   If DtaConsulta.Recordset("Periodo") = Periodo Then
'     DtaConsulta.Recordset("Actual") = 1
'     Mes = Me.DtaConsulta.Recordset("mes")
'     Periodo = Me.DtaConsulta.Recordset("Periodo")
'   End If
'   Me.DtaConsulta.Recordset.Update
'
' Me.DtaConsulta.Recordset.MoveNext
'Loop

Me.DtaConsulta.RecordSource = "SELECT Periodo,NumNomina, Mes, Ano From Nomina Where (NumNomina = " & NumeroNomina & ")"
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
   Me.DtaConsulta.Recordset("mes") = Mes
   Me.DtaConsulta.Recordset("Periodo") = Periodo
  Me.DtaConsulta.Recordset.Update
End If

Unload Me

Exit Sub
TipoErrs:
ControlErrores
Unload Me



End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DBComboPeriodo_Change()
Dim Periodo As Integer, Ano As Integer
CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
Periodo = val(Me.DBComboPeriodo.Text)

If Periodo >= 13 Then
' Exit Sub
End If

Ano = val(Me.txtAno.Text)
Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Periodo & ")AND (año = " & Ano & ")"
'InputBox "", "", Me.DtaConsulta.RecordSource
Me.DtaConsulta.Refresh

If Not DtaConsulta.Recordset.EOF Then
   Me.MtxtFecha.Value = Me.DtaConsulta.Recordset("Final")
   Me.MtxtFechaini.Value = Me.DtaConsulta.Recordset("Inicio")
   LblFechaLarga.Caption = "Pago: " + Format(MtxtFecha.Value, "Long Date")
End If
End Sub

Private Sub DbgrTipoNominas_DblClick()
On Error GoTo TipoErrs
Dim MiDiaSemana As Integer, Fecha1 As Long, Fecha2 As Long
Dim Año As Integer, Periodo As Integer

CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")

'Me.DtaPeriodos.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (Calculada = 0)"
'Me.DtaPeriodos.Refresh

MiDiaSemana = Weekday(Now, vbMonday)
'MsgBox MiDiaSemana
'MsgBox ("#" + DtaTipoNomina.Recordset("Periodo") + "#" + "#" + "Semanal Sábados" + "#")



If DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
   Select Case MiDiaSemana
          Case 1
             MtxtFecha = Format(Now + 4, "dd/mm/yyyy")
          Case 2
             MtxtFecha = Format(Now + 3, "dd/mm/yyyy")
          Case 3
             MtxtFecha = Format(Now + 2, "dd/mm/yyyy")
          Case 4
             MtxtFecha = Format(Now + 1, "dd/mm/yyyy")
          Case 5
             MtxtFecha = Format(Now + 0, "dd/mm/yyyy")
          Case 6
             MtxtFecha = Format(Now + 6, "dd/mm/yyyy")
          Case 7
             MtxtFecha = Format(Now + 5, "dd/mm/yyyy")
    End Select
    MtxtFechaini = Str(CDate(MtxtFecha) - 7)
 End If
 
If DtaTipoNomina.Recordset("Periodo") = "Semanal Viernes" Then
'////////////////////////////////BUSCO LA NOMINA ACTIVA ACTUAL/////////////
'//////////////////////////////////////////////////////////////////////////////
   Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
   Me.DtaConsulta.Refresh
   If Not DtaConsulta.Recordset.EOF Then
      Año = Me.DtaConsulta.Recordset("año")
      Me.txtAno.Text = Me.DtaConsulta.Recordset("año")
      Periodo = Me.DtaConsulta.Recordset("Periodo")
      Fecha1 = Me.DtaConsulta.Recordset("Inicio")
      Fecha2 = Me.DtaConsulta.Recordset("Final")
'////////////////////BUSCO SI LA NOMINA ACTUAL ESTA CERRADA/////////////
'////////////////////////////////////////////////////////////////////////
      Me.DtaConsulta.RecordSource = "SELECT NumNomina, CodTipoNomina, FechaNominaINI, FechaNomina, Activa, Procesada, Cerrada From Nomina WHERE (FechaNominaINI = " & Fecha1 & ") AND (FechaNomina = " & Fecha2 & ") AND (Cerrada = 1) AND (CodTipoNomina = '" & CodTipoNomina & "')AND (FechaNomina BETWEEN " & Fecha1 & " AND " & Fecha2 & ")"
      Me.DtaConsulta.Refresh
'      InputBox "", "", Me.DtaConsulta.RecordSource
'///////////////////////SI NO ESTA CERRADA,DEJO LA ACTUAL/////
'///////////////////////////////////////////////////////////////////
      If DtaConsulta.Recordset.EOF Then
        Me.MtxtFechaini.Value = Fecha1
        Me.MtxtFecha.Value = Fecha2
        Me.MtxtFecha.Enabled = False
        Me.MtxtFechaini.Enabled = False
      Else
'///////////////////////SI ESTA CERRADA BUSCO LA SIGUIENTE////////////////////
'//////////////////////////////////////////////////////////////////////////////
        Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (año = " & Año & ") AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Periodo + 1 & ") ORDER BY Periodo"
        Me.DtaConsulta.Refresh
'        InputBox "", "", Me.DtaConsulta.RecordSource
        If Not DtaConsulta.Recordset.EOF Then
            Me.MtxtFechaini.Value = Me.DtaConsulta.Recordset("Inicio")
            Me.MtxtFecha.Value = Me.DtaConsulta.Recordset("Final")
            Me.MtxtFecha.Enabled = False
            Me.MtxtFechaini.Enabled = False
        Else
            Cadena = "No se puede Asignar ningun Periodo, a esta Nomina" & vbLf
            Cadena = Cadena & "Necesitar Crear un nuevo Periodo,de los contrario" & vbLf
            Cadena = Cadena & "no podra activar la Nomina"
            MsgBox Cadena, vbCritical
        End If
      End If
   Else
     MsgBox "No esta Definido el perido actual de la nomina", vbCritical
     Exit Sub
   End If
End If

If DtaTipoNomina.Recordset("Periodo") = "Semanal Sabado" Then
'////////////////////////////////BUSCO LA NOMINA ACTIVA ACTUAL/////////////
'//////////////////////////////////////////////////////////////////////////////
   Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
   Me.DtaConsulta.Refresh
   If Not DtaConsulta.Recordset.EOF Then
      Año = Me.DtaConsulta.Recordset("año")
      Me.txtAno.Text = Me.DtaConsulta.Recordset("año")
      Periodo = Me.DtaConsulta.Recordset("Periodo")
      Fecha1 = Me.DtaConsulta.Recordset("Inicio")
      Fecha2 = Me.DtaConsulta.Recordset("Final")
'////////////////////BUSCO SI LA NOMINA ACTUAL ESTA CERRADA/////////////
'////////////////////////////////////////////////////////////////////////
      Me.DtaConsulta.RecordSource = "SELECT NumNomina, CodTipoNomina, FechaNominaINI, FechaNomina, Activa, Procesada, Cerrada From Nomina WHERE (FechaNominaINI = " & Fecha1 & ") AND (FechaNomina = " & Fecha2 & ") AND (Cerrada = 1) AND (CodTipoNomina = '" & CodTipoNomina & "')AND (FechaNomina BETWEEN " & Fecha1 & " AND " & Fecha2 & ")"
      Me.DtaConsulta.Refresh
'      InputBox "", "", Me.DtaConsulta.RecordSource
'///////////////////////SI NO ESTA CERRADA,DEJO LA ACTUAL/////
'///////////////////////////////////////////////////////////////////
      If DtaConsulta.Recordset.EOF Then
        Me.MtxtFechaini.Value = Fecha1
        Me.MtxtFecha.Value = Fecha2
        Me.MtxtFecha.Enabled = False
        Me.MtxtFechaini.Enabled = False
      Else
'///////////////////////SI ESTA CERRADA BUSCO LA SIGUIENTE////////////////////
'//////////////////////////////////////////////////////////////////////////////
        Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (año = " & Año & ") AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Periodo + 1 & ") ORDER BY Periodo"
        Me.DtaConsulta.Refresh
'        InputBox "", "", Me.DtaConsulta.RecordSource
        If Not DtaConsulta.Recordset.EOF Then
            Me.MtxtFechaini.Value = Me.DtaConsulta.Recordset("Inicio")
            Me.MtxtFecha.Value = Me.DtaConsulta.Recordset("Final")
            Me.MtxtFecha.Enabled = False
            Me.MtxtFechaini.Enabled = False
        Else
            Cadena = "No se puede Asignar ningun Periodo, a esta Nomina" & vbLf
            Cadena = Cadena & "Necesitar Crear un nuevo Periodo,de los contrario" & vbLf
            Cadena = Cadena & "no podra activar la Nomina"
            MsgBox Cadena, vbCritical
        End If
      End If
   Else
     MsgBox "No esta Definido el perido actual de la nomina", vbCritical
     Exit Sub
   End If
End If

If DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Viernes" Then
' Select Case MiDiaSemana
    '      Case 1
     '        MtxtFecha = Format(Now + 12, "dd/mm/yyyy")
      '    Case 2
       '      MtxtFecha = Format(Now + 11, "dd/mm/yyyy")
        '  Case 3
         '    MtxtFecha = Format(Now + 10, "dd/mm/yyyy")
          'Case 4
           '  MtxtFecha = Format(Now + 9, "dd/mm/yyyy")
          'Case 5
         '    MtxtFecha = Format(Now + 8, "dd/mm/yyyy")
        '  Case 6
       '      MtxtFecha = Format(Now + 7, "dd/mm/yyyy")
      '    Case 7
     '        MtxtFecha = Format(Now + 6, "dd/mm/yyyy")
    'End Select
   ' MtxtFechaini = Str(CDate(MtxtFecha) - 14)
   
       '////////////////////////////////BUSCO LA NOMINA ACTIVA ACTUAL/////////////
            '//////////////////////////////////////////////////////////////////////////////
               Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
               Me.DtaConsulta.Refresh
               If Not DtaConsulta.Recordset.EOF Then
                  Año = Me.DtaConsulta.Recordset("año")
                  Me.txtAno.Text = Me.DtaConsulta.Recordset("año")
                  Periodo = Me.DtaConsulta.Recordset("Periodo")
                  Fecha1 = Me.DtaConsulta.Recordset("Inicio")
                  Fecha2 = Me.DtaConsulta.Recordset("Final")
            '////////////////////BUSCO SI LA NOMINA ACTUAL ESTA CERRADA/////////////
            '////////////////////////////////////////////////////////////////////////
                  Me.DtaConsulta.RecordSource = "SELECT NumNomina, CodTipoNomina, FechaNominaINI, FechaNomina, Activa, Procesada, Cerrada From Nomina WHERE (FechaNominaINI = " & Fecha1 & ") AND (FechaNomina = " & Fecha2 & ") AND (Cerrada = 1) AND (CodTipoNomina = '" & CodTipoNomina & "')AND (FechaNomina BETWEEN " & Fecha1 & " AND " & Fecha2 & ")"
                  Me.DtaConsulta.Refresh
            '      InputBox "", "", Me.DtaConsulta.RecordSource
            '///////////////////////SI NO ESTA CERRADA,DEJO LA ACTUAL/////
            '///////////////////////////////////////////////////////////////////
                  If DtaConsulta.Recordset.EOF Then
                    Me.MtxtFechaini.Value = Fecha1
                    Me.MtxtFecha.Value = Fecha2
                    Me.MtxtFecha.Enabled = False
                    Me.MtxtFechaini.Enabled = False
                  Else
            '///////////////////////SI ESTA CERRADA BUSCO LA SIGUIENTE////////////////////
            '//////////////////////////////////////////////////////////////////////////////
                    Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (año = " & Año & ") AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Periodo + 1 & ") ORDER BY Periodo"
                    Me.DtaConsulta.Refresh
            '        InputBox "", "", Me.DtaConsulta.RecordSource
                    If Not DtaConsulta.Recordset.EOF Then
                        Me.MtxtFechaini.Value = Me.DtaConsulta.Recordset("Inicio")
                        Me.MtxtFecha.Value = Me.DtaConsulta.Recordset("Final")
                        Me.MtxtFecha.Enabled = False
                        Me.MtxtFechaini.Enabled = False
                    Else
                        Cadena = "No se puede Asignar ningun Periodo, a esta Nomina" & vbLf
                        Cadena = Cadena & "Necesitar Crear un nuevo Periodo,de los contrario" & vbLf
                        Cadena = Cadena & "no podra activar la Nomina"
                        MsgBox Cadena, vbCritical
                    End If
                  End If
               Else
                 MsgBox "No esta Definido el perido actual de la nomina", vbCritical
                 Exit Sub
               End If
    
ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then

        ' Select Case MiDiaSemana
        '          Case 1
        '             MtxtFecha = Format(Now + 13, "dd/mm/yyyy")
        '          Case 2
        '             MtxtFecha = Format(Now + 12, "dd/mm/yyyy")
        '          Case 3
        '             MtxtFecha = Format(Now + 11, "dd/mm/yyyy")
        '          Case 4
        '             MtxtFecha = Format(Now + 10, "dd/mm/yyyy")
        '          Case 5
        '             MtxtFecha = Format(Now + 9, "dd/mm/yyyy")
        '          Case 6
        '             MtxtFecha = Format(Now + 8, "dd/mm/yyyy")
        '          Case 7
        '             MtxtFecha = Format(Now + 7, "dd/mm/yyyy")
        '    End Select
        '    MtxtFechaini = Str(CDate(MtxtFecha) - 14)
            
            '////////////////////////////////BUSCO LA NOMINA ACTIVA ACTUAL/////////////
            '//////////////////////////////////////////////////////////////////////////////
               Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
               Me.DtaConsulta.Refresh
               If Not DtaConsulta.Recordset.EOF Then
                  Año = Me.DtaConsulta.Recordset("año")
                  Me.txtAno.Text = Me.DtaConsulta.Recordset("año")
                  Periodo = Me.DtaConsulta.Recordset("Periodo")
                  Fecha1 = Me.DtaConsulta.Recordset("Inicio")
                  Fecha2 = Me.DtaConsulta.Recordset("Final")
            '////////////////////BUSCO SI LA NOMINA ACTUAL ESTA CERRADA/////////////
            '////////////////////////////////////////////////////////////////////////
                  Me.DtaConsulta.RecordSource = "SELECT NumNomina, CodTipoNomina, FechaNominaINI, FechaNomina, Activa, Procesada, Cerrada From Nomina WHERE (FechaNominaINI = " & Fecha1 & ") AND (FechaNomina = " & Fecha2 & ") AND (Cerrada = 1) AND (CodTipoNomina = '" & CodTipoNomina & "')AND (FechaNomina BETWEEN " & Fecha1 & " AND " & Fecha2 & ")"
                  Me.DtaConsulta.Refresh
            '      InputBox "", "", Me.DtaConsulta.RecordSource
            '///////////////////////SI NO ESTA CERRADA,DEJO LA ACTUAL/////
            '///////////////////////////////////////////////////////////////////
                  If DtaConsulta.Recordset.EOF Then
                    Me.MtxtFechaini.Value = Fecha1
                    Me.MtxtFecha.Value = Fecha2
                    Me.MtxtFecha.Enabled = False
                    Me.MtxtFechaini.Enabled = False
                  Else
            '///////////////////////SI ESTA CERRADA BUSCO LA SIGUIENTE////////////////////
            '//////////////////////////////////////////////////////////////////////////////
                    Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (año = " & Año & ") AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Periodo + 1 & ") ORDER BY Periodo"
                    Me.DtaConsulta.Refresh
            '        InputBox "", "", Me.DtaConsulta.RecordSource
                    If Not DtaConsulta.Recordset.EOF Then
                        Me.MtxtFechaini.Value = Me.DtaConsulta.Recordset("Inicio")
                        Me.MtxtFecha.Value = Me.DtaConsulta.Recordset("Final")
                        Me.MtxtFecha.Enabled = False
                        Me.MtxtFechaini.Enabled = False
                    Else
                        Cadena = "No se puede Asignar ningun Periodo, a esta Nomina" & vbLf
                        Cadena = Cadena & "Necesitar Crear un nuevo Periodo,de los contrario" & vbLf
                        Cadena = Cadena & "no podra activar la Nomina"
                        MsgBox Cadena, vbCritical
                    End If
                  End If
               Else
                 MsgBox "No esta Definido el perido actual de la nomina", vbCritical
                 Exit Sub
               End If
    
    
    
    
    
    
    
    
End If

If DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
'////////////////////////////////BUSCO LA NOMINA ACTIVA ACTUAL/////////////
'//////////////////////////////////////////////////////////////////////////////
   Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
'   Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (año = " & Año & ") AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Periodo + 1 & ") ORDER BY Periodo"
   Me.DtaConsulta.Refresh
   If Not DtaConsulta.Recordset.EOF Then
      Año = Me.DtaConsulta.Recordset("año")
      Me.txtAno.Text = Me.DtaConsulta.Recordset("año")
      Periodo = Me.DtaConsulta.Recordset("Periodo")
      Fecha1 = Me.DtaConsulta.Recordset("Inicio")
      Fecha2 = Me.DtaConsulta.Recordset("Final")
'////////////////////BUSCO SI LA NOMINA ACTUAL ESTA CERRADA/////////////
'////////////////////////////////////////////////////////////////////////
      Me.DtaConsulta.RecordSource = "SELECT NumNomina, CodTipoNomina, FechaNominaINI, FechaNomina, Activa, Procesada, Cerrada From Nomina WHERE     (FechaNominaINI = " & Fecha1 & ") AND (FechaNomina = " & Fecha2 & ") AND (Cerrada = 1) AND (CodTipoNomina = '" & CodTipoNomina & "')AND (FechaNomina BETWEEN " & Fecha1 & " AND " & Fecha2 & ")"
      Me.DtaConsulta.Refresh
'///////////////////////SI NO ESTA CERRADA,DEJO LA ACTUAL/////
'///////////////////////////////////////////////////////////////////
      If Not DtaConsulta.Recordset.EOF Then
        Me.MtxtFechaini.Value = Fecha1
        Me.MtxtFecha.Value = Fecha2
        Me.MtxtFecha.Enabled = True
        Me.MtxtFechaini.Enabled = False
      Else
'///////////////////////SI ESTA CERRADA BUSCO LA SIGUIENTE////////////////////
'//////////////////////////////////////////////////////////////////////////////
        Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (año = " & Año & ") AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Periodo + 1 & ") ORDER BY Periodo"
        Me.DtaConsulta.Refresh
        If Not DtaConsulta.Recordset.EOF Then
            Me.MtxtFechaini.Value = Me.DtaConsulta.Recordset("Inicio")
            Me.MtxtFecha.Value = Me.DtaConsulta.Recordset("Final")
            Me.MtxtFecha.Enabled = True
            Me.MtxtFechaini.Enabled = False
            Periodo = Periodo + 1
        Else
            Cadena = "No se puede Asignar ningun Periodo, a esta Nomina" & vbLf
            Cadena = Cadena & "Necesitar Crear un nuevo Periodo,de los contrario" & vbLf
            Cadena = Cadena & "no podra activar la Nomina"
            MsgBox Cadena, vbCritical
        End If
      End If
  Else
     MsgBox "No Existe Ninguna Nomina Activa", vbCritical
     Exit Sub
  End If
End If
If DtaTipoNomina.Recordset("Periodo") = "Mensual" Then
'        MtxtFecha = DateSerial(Year(Now), Month(Now) + 1, 0)
'        MtxtFechaini = "1/" + Str(Month(Now)) + "/" + Str(Year(Now))
'////////////////////////////////BUSCO LA NOMINA ACTIVA ACTUAL/////////////
'//////////////////////////////////////////////////////////////////////////////
   Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
'   Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (año = " & Año & ") AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Periodo + 1 & ") ORDER BY Periodo"
   Me.DtaConsulta.Refresh
   If Not DtaConsulta.Recordset.EOF Then
      Año = Me.DtaConsulta.Recordset("año")
      Me.txtAno.Text = Me.DtaConsulta.Recordset("año")
      Periodo = Me.DtaConsulta.Recordset("Periodo")
      Fecha1 = Me.DtaConsulta.Recordset("Inicio")
      Fecha2 = Me.DtaConsulta.Recordset("Final")
'////////////////////BUSCO SI LA NOMINA ACTUAL ESTA CERRADA/////////////
'////////////////////////////////////////////////////////////////////////
      Me.DtaConsulta.RecordSource = "SELECT NumNomina, CodTipoNomina, FechaNominaINI, FechaNomina, Activa, Procesada, Cerrada From Nomina WHERE     (FechaNominaINI = " & Fecha1 & ") AND (FechaNomina = " & Fecha2 & ") AND (Cerrada = 1) AND (CodTipoNomina = '" & CodTipoNomina & "')AND (FechaNomina BETWEEN " & Fecha1 & " AND " & Fecha2 & ")"
      Me.DtaConsulta.Refresh
'///////////////////////SI NO ESTA CERRADA,DEJO LA ACTUAL/////
'///////////////////////////////////////////////////////////////////
      If Not DtaConsulta.Recordset.EOF Then
        Me.MtxtFechaini.Value = Fecha1
        Me.MtxtFecha.Value = Fecha2
        Me.MtxtFecha.Enabled = True
        Me.MtxtFechaini.Enabled = False
      Else
'///////////////////////SI ESTA CERRADA BUSCO LA SIGUIENTE////////////////////
'//////////////////////////////////////////////////////////////////////////////
        Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (año = " & Año & ") AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Periodo + 1 & ") ORDER BY Periodo"
        Me.DtaConsulta.Refresh
        If Not DtaConsulta.Recordset.EOF Then
            Me.MtxtFechaini.Value = Me.DtaConsulta.Recordset("Inicio")
            Me.MtxtFecha.Value = Me.DtaConsulta.Recordset("Final")
            Me.MtxtFecha.Enabled = True
            Me.MtxtFechaini.Enabled = False
            Periodo = Periodo + 1
        Else
            Cadena = "No se puede Asignar ningun Periodo, a esta Nomina" & vbLf
            Cadena = Cadena & "Necesitar Crear un nuevo Periodo,de los contrario" & vbLf
            Cadena = Cadena & "no podra activar la Nomina"
            MsgBox Cadena, vbCritical
        End If
      End If
  Else
     MsgBox "No Existe Ninguna Nomina Activa", vbCritical
     Exit Sub
  End If

End If


If DtaTipoNomina.Recordset("Periodo") = "Trimestral" Then
       Select Case Month(Now)
       Case 1 To 3
            MtxtFecha = "31/03/" + Year(Now)
            MtxtFechaini = "1/01/" + Year(Now)
       Case 4 To 6
            MtxtFecha = "30/06/" + Year(Now)
            MtxtFechaini = "1/04/" + Year(Now)
       Case 7 To 9
            MtxtFecha = "30/09/" + Year(Now)
            MtxtFechaini = "1/07/" + Year(Now)
       Case 10 To 12
            MtxtFecha = "31/12/" + Year(Now)
            MtxtFechaini = "1/10/" + Year(Now)
       End Select
End If



If DtaTipoNomina.Recordset("Periodo") = "Semestral" Then
       Select Case Month(Now)
       Case 1 To 6
            MtxtFecha = "30/06/" + Year(Now)
            MtxtFechaini = "1/01/" + Year(Now)
       Case 7 To 12
            MtxtFecha = "31/12/" + Year(Now)
            MtxtFechaini = "1/07/" + Year(Now)
       End Select
       
End If

Me.DtaPeriodos.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual, Calculada From Fecha_Planilla WHERE (CodTipoNomina = '" & CodTipoNomina & "') AND (Calculada = 0)AND (año = " & val(Me.txtAno.Text) & ")"
Me.DtaPeriodos.Refresh

LblFechaLarga.Caption = "Pago: " + Format(MtxtFecha.Value, "Long Date")
Me.DBComboPeriodo.Text = Periodo
Exit Sub
TipoErrs:
ControlErrores
Unload Me


End Sub

Private Sub Form_Load()
On Error GoTo TipoErrs
Dim Fecha As Date




MDIPrimero.Skin1.ApplySkin hWnd
With Me.DtaPeriodos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
 
End With


With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaConsecutivos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Consecutivos"
End With

With Me.DtaNomina
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaTipoNomina
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

 Me.DbgrTipoNominas.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DbgrTipoNominas.OddRowStyle.BackColor = &H80000005
 Me.DbgrTipoNominas.AlternatingRowStyle = True

 Me.DBComboPeriodo.ListField = "Periodo"
  
Dim SQLTipoNomina

SQLTipoNomina = "SELECT TipoNomina.CodTipoNomina, TipoNomina.Nomina, TipoNomina.Periodo, TipoNomina.UltFecha, TipoNomina.TipoPago, TipoNomina.Moneda, TipoNomina.MantValor, TipoNomina.Activa From TipoNomina WHERE (((TipoNomina.Activa)=0))"
DtaTipoNomina.RecordSource = SQLTipoNomina
DtaTipoNomina.Refresh

MtxtFechaini.Value = Now
MtxtFecha.Value = Now


'Fecha = "01/06/2014"
'If DateDiff("d", Now, Fecha) < 0 Then
'  Unload Me
'End If


Exit Sub
TipoErrs:
ControlErrores
Unload Me
End Sub

Private Sub MtxtFecha_Change()
On Error GoTo TipoErrs

LblFechaLarga.Caption = "Pago: " + Format(MtxtFecha.Value, "Long Date")

Exit Sub
TipoErrs:
ControlErrores
Unload Me
End Sub

Private Sub MtxtFecha_LostFocus()


If Not IsDate(MtxtFecha.Value) Then
MsgBox "La fecha digitada no es correcta"
End If

Exit Sub

TipoErrs:
ControlErrores
Unload Me

End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub
