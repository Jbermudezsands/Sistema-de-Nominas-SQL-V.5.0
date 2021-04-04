VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmListadoNominaVacaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LISTADO NOMINAS DE VACACIONES Y 13VO MES"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   10695
   Begin VB.CommandButton CmdSalir 
      DownPicture     =   "FrmListadoVaca.frx":0000
      Height          =   375
      Left            =   9120
      Picture         =   "FrmListadoVaca.frx":1AE2
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8160
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Nominas Vacaciones"
      TabPicture(0)   =   "FrmListadoVaca.frx":35C4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(5)=   "AdoDetalleTotal"
      Tab(0).Control(6)=   "AdoDetalleNominaVaca"
      Tab(0).Control(7)=   "DbgrdetalleNominas"
      Tab(0).Control(8)=   "Picture1"
      Tab(0).Control(9)=   "DbgrNominas"
      Tab(0).Control(10)=   "TxtInss"
      Tab(0).Control(11)=   "TxtAdelantos"
      Tab(0).Control(12)=   "TxtDiasDescuentos"
      Tab(0).Control(13)=   "TxtDiasPagar"
      Tab(0).Control(14)=   "TxtSalarioMensual"
      Tab(0).Control(15)=   "AdoNominasVaca"
      Tab(0).Control(16)=   "Frame1"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Nominas 13vo Mes"
      TabPicture(1)   =   "FrmListadoVaca.frx":35E0
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label12"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label16"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label20"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "AdoDetalleNomina13vo"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "AdoNomina13vo"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "DbgrdetalleNominas13vo"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "DbgrNominas13vo"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "TxtSalarioPagar13vo"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "TxtAdelanto13vo"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "TxtDiasPagar13vo"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "TxtSalarioMensual13vo"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Picture2"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame2"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.Frame Frame2 
         Height          =   3135
         Left            =   8640
         TabIndex        =   31
         Top             =   4560
         Width           =   1695
         Begin VB.CommandButton Command5 
            DownPicture     =   "FrmListadoVaca.frx":35FC
            Height          =   375
            Left            =   120
            Picture         =   "FrmListadoVaca.frx":50DE
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            DownPicture     =   "FrmListadoVaca.frx":6BC0
            Height          =   375
            Left            =   120
            Picture         =   "FrmListadoVaca.frx":86A2
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton Command3 
            DownPicture     =   "FrmListadoVaca.frx":A184
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            Picture         =   "FrmListadoVaca.frx":BC66
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   3000
            Visible         =   0   'False
            Width           =   1455
         End
         Begin SmartButtonProject.SmartButton SmartButton1 
            Height          =   975
            Left            =   240
            TabIndex        =   37
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1720
            Caption         =   "Historial Manual"
            Picture         =   "FrmListadoVaca.frx":D748
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
      End
      Begin VB.Frame Frame1 
         Height          =   3135
         Left            =   -66360
         TabIndex        =   25
         Top             =   4560
         Width           =   1695
         Begin VB.CheckBox ChkRestar 
            Caption         =   "Restar Inss Nomina"
            Height          =   495
            Left            =   120
            TabIndex        =   38
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CommandButton CmdAnularNomina 
            DownPicture     =   "FrmListadoVaca.frx":E022
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            Picture         =   "FrmListadoVaca.frx":FB04
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   3000
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton CmdPrNomina 
            DownPicture     =   "FrmListadoVaca.frx":115E6
            Height          =   375
            Left            =   120
            Picture         =   "FrmListadoVaca.frx":130C8
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton CmdPrColilla 
            DownPicture     =   "FrmListadoVaca.frx":14BAA
            Height          =   375
            Left            =   120
            Picture         =   "FrmListadoVaca.frx":1668C
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   240
            Width           =   1455
         End
         Begin SmartButtonProject.SmartButton SmartButton2 
            Height          =   975
            Left            =   240
            TabIndex        =   36
            Top             =   1800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1720
            Caption         =   "Historial Manual"
            Picture         =   "FrmListadoVaca.frx":1816E
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
      End
      Begin MSAdodcLib.Adodc AdoNominasVaca 
         Height          =   375
         Left            =   -73800
         Top             =   1560
         Visible         =   0   'False
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
         Caption         =   "AdoNominasVaca"
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
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C1A1&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   10215
         TabIndex        =   23
         Top             =   360
         Width           =   10215
         Begin VB.Image Image1 
            Height          =   1020
            Left            =   0
            Picture         =   "FrmListadoVaca.frx":18A48
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1650
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00800000&
            BorderWidth     =   2
            X1              =   0
            X2              =   10200
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "LISTADO DE NOMINAS 13vo Mes"
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
            Left            =   2640
            TabIndex        =   24
            Top             =   360
            Width           =   5400
         End
      End
      Begin VB.TextBox TxtSalarioMensual13vo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         DataField       =   "TotalSalarioBasico"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox TxtDiasPagar13vo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         DataField       =   "TotalOtrosIngresos"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox TxtAdelanto13vo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         DataField       =   "TotalDestajo"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox TxtSalarioPagar13vo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         DataField       =   "TotalMontoIR"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox TxtSalarioMensual 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         DataField       =   "TotalSalarioBasico"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   -66240
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox TxtDiasPagar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         DataField       =   "TotalOtrosIngresos"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   -66240
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox TxtDiasDescuentos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         DataField       =   "TotalDestajo"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   -66240
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox TxtAdelantos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF00&
         DataField       =   "TotalComisiones"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   -66240
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox TxtInss 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         DataField       =   "TotalMontoINSS"
         DataSource      =   "DtaNominas"
         Height          =   285
         Left            =   -66240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   3240
         Width           =   1335
      End
      Begin TrueOleDBGrid70.TDBGrid DbgrNominas 
         Bindings        =   "FrmListadoVaca.frx":1A58A
         Height          =   2415
         Left            =   -74880
         TabIndex        =   1
         Top             =   1920
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4260
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
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
      Begin TrueOleDBGrid70.TDBGrid DbgrNominas13vo 
         Bindings        =   "FrmListadoVaca.frx":1A5A7
         Height          =   2415
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4260
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
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
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C1A1&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -74880
         ScaleHeight     =   1095
         ScaleWidth      =   10215
         TabIndex        =   12
         Top             =   480
         Width           =   10215
         Begin VB.Image Image2 
            Height          =   1080
            Left            =   0
            Picture         =   "FrmListadoVaca.frx":1A5C3
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1845
         End
         Begin VB.Label lbltitulo 
            BackStyle       =   0  'Transparent
            Caption         =   "LISTADO DE NOMINAS VACACIONES"
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
            Left            =   2640
            TabIndex        =   13
            Top             =   360
            Width           =   5400
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00800000&
            BorderWidth     =   2
            X1              =   0
            X2              =   10200
            Y1              =   1080
            Y2              =   1080
         End
      End
      Begin TrueOleDBGrid70.TDBGrid DbgrdetalleNominas 
         Bindings        =   "FrmListadoVaca.frx":2307B
         Height          =   3135
         Left            =   -74880
         TabIndex        =   29
         Top             =   4560
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5530
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
      Begin TrueOleDBGrid70.TDBGrid DbgrdetalleNominas13vo 
         Bindings        =   "FrmListadoVaca.frx":2309E
         Height          =   3135
         Left            =   120
         TabIndex        =   35
         Top             =   4560
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5530
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
      Begin MSAdodcLib.Adodc AdoDetalleNominaVaca 
         Height          =   375
         Left            =   -70560
         Top             =   1560
         Visible         =   0   'False
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
         Caption         =   "AdoDetalleNominaVaca"
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
      Begin MSAdodcLib.Adodc AdoDetalleTotal 
         Height          =   375
         Left            =   -71640
         Top             =   7560
         Visible         =   0   'False
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
         Caption         =   "AdoDetalleNominaVaca"
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
      Begin MSAdodcLib.Adodc AdoNomina13vo 
         Height          =   375
         Left            =   1200
         Top             =   1440
         Visible         =   0   'False
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
         Caption         =   "AdoNominas13vo"
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
      Begin MSAdodcLib.Adodc AdoDetalleNomina13vo 
         Height          =   375
         Left            =   4440
         Top             =   1440
         Visible         =   0   'False
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
         Caption         =   "AdoDetalleNomina13vo"
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
      Begin VB.Label Label20 
         Caption         =   "Dias a Pagar"
         Height          =   255
         Left            =   6960
         TabIndex        =   22
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Salario a Pagar"
         Height          =   255
         Left            =   6960
         TabIndex        =   21
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Adelanto 13vo"
         Height          =   255
         Left            =   6960
         TabIndex        =   20
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Salario Mensual"
         Height          =   255
         Left            =   6960
         TabIndex        =   19
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Dias a Pagar"
         Height          =   255
         Left            =   -68160
         TabIndex        =   11
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Adelanto Vacaciones"
         Height          =   255
         Left            =   -68160
         TabIndex        =   10
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "INSS"
         Height          =   255
         Left            =   -68160
         TabIndex        =   9
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Dias Descuentos"
         Height          =   255
         Left            =   -68160
         TabIndex        =   8
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Salario Mensual"
         Height          =   255
         Left            =   -68160
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc DtaTipoNominas 
      Height          =   375
      Left            =   480
      Top             =   8160
      Visible         =   0   'False
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
      Caption         =   "DtaTipoNominas"
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
      Left            =   3600
      Top             =   8160
      Visible         =   0   'False
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
End
Attribute VB_Name = "FrmListadoNominaVacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdPrColilla_Click()
On Error GoTo TipoErrs
Dim Espacio As String, FechaIni As Date, FechaFin As Date
Dim NumNomina As Integer, CodTipoNomina As String
Espacio = " "
NumNomina = Me.AdoNominasVaca.Recordset("NumNomVaca")
CodTipoNomina = Me.AdoNominasVaca.Recordset("CodTipoNomina")
 
FechaIni = Me.AdoNominasVaca.Recordset("FechaIni")
FechaFin = Me.AdoNominasVaca.Recordset("FechaFin")

 SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                     "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomina & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "

ArepColillaVaca.LblPeriodos.Caption = "Desde   " & FechaIni & "   Hasta " & FechaFin
ArepColillaVaca.LblPeriodo.Caption = Format(FechaIni, "dd/mm/yyyy") & "   Hasta   " & Format(FechaFin, "dd/mm/yyyy")
ArepColillaVaca.lblTitulo.Caption = Titulo
ArepColillaVaca.AdoColillas.ConnectionString = ConexionReporte
ArepColillaVaca.AdoColillas.Source = SQlReportes
ArepColillaVaca.Show 1



Exit Sub
TipoErrs:
ControlErrores

End Sub

Private Sub CmdprNomina_Click()
On Error GoTo TipoErrs
Dim Espacio As String, NombreNomina As String
Espacio = " "
Dim rpt As Object
Dim fPreview As New FrmPreview

NumNomina = Me.AdoNominasVaca.Recordset("NumNomVaca")
CodTipoNomina = Me.AdoNominasVaca.Recordset("CodTipoNomina")
 
FechaIni = Me.AdoNominasVaca.Recordset("FechaIni")
FechaFin = Me.AdoNominasVaca.Recordset("FechaFin")

Me.DtaTipoNominas.Refresh
Do While Not Me.DtaTipoNominas.Recordset.EOF
If Me.DtaTipoNominas.Recordset("CodTipoNomina") = CodTipoNomina Then
   NombreNomina = Me.DtaTipoNominas.Recordset("Nomina")
   Exit Do
End If
DtaTipoNominas.Recordset.MoveNext
Loop
'
'
'      Nom13vo.LblTitulo.Caption = Titulo
'      Nom13vo.LblSubtitulo.Caption = SubTitulo
'      Nom13vo.ImgLogo.Picture = LoadPicture(RutaLogo)
'
'      Nom13vo.LblFecha.Caption = "Desde " + Format(FechaIni, "mm/dd/yyyy") + " Hasta " + Format(FechaFin, "mm/dd/yyyy")
'      Nom13vo.LblFechaHoy = Format(Now, "dddddd")
'      Nom13vo.DataControl1.ConnectionString = ConexionReporte
'
'If NombreNomina <> "Administracion" Then
'
'
'       SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
'                     "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomina & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
'
'Else
'
'    SQlReportes = "SELECT NomVaca.NumNomVaca AS NumNom13Mes, DetalleNomVaca.Inss, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
'               "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE (NomVaca.NumNomVaca = " & NumNomina & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Empleado.CodEmpleado1 "
'
'
'End If



Espacio = " "
'NumNomina = Me.TxtNumNomVaca.Text

'DtaTipoNomina.Refresh
'Do While Not DtaTipoNomina.Recordset.EOF
'If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
'   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
'   Exit Do
'End If
'DtaTipoNomina.Recordset.MoveNext
'Loop


      ArepNomVacaciones.lblTitulo.Caption = Titulo
      ArepNomVacaciones.LblSubtitulo.Caption = SubTitulo
      ArepNomVacaciones.ImgLogo.Picture = LoadPicture(RutaLogo)
      Espacio = FechaPeriodo(CDate(Me.AdoNominasVaca.Recordset("FechaIni")), CDate(Me.AdoNominasVaca.Recordset("FechaFin")), CodTipoNomina)
      ArepNomVacaciones.LblFecha.Caption = "Desde " + Format(FechaIni, "mm/dd/yyyy") + " Hasta " + Format(FechaFin, "mm/dd/yyyy")
      ArepNomVacaciones.LblFechaHoy = Format(Now, "dddddd")
      ArepNomVacaciones.DataControl1.ConnectionString = ConexionReporte
      
If NombreNomina <> "Administracion" Then
      
 If Me.ChkRestar.Value = 1 Then
       SQlReportes = "SELECT Empleado.CodEmpleado,NomVaca.NumNomVaca , DetalleNomVaca.Inss, DetalleNomVaca.Ir, Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                     "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomina & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
 Else
      SQlReportes = "SELECT Empleado.CodEmpleado,NomVaca.NumNomVaca , DetalleNomVaca.Inss, DetalleNomVaca.Ir,Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado,Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                    "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomina & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "')  AND (DetalleNomVaca.DiasAPagar <> 0) ORDER BY Empleado.CodEmpleado1 "
 End If
Else

If Me.ChkRestar.Value = 1 Then
    SQlReportes = "SELECT Empleado.CodEmpleado,NomVaca.NumNomVaca , DetalleNomVaca.Inss, DetalleNomVaca.Ir,Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN                      DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
               "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomina & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "

Else
      SQlReportes = "SELECT Empleado.CodEmpleado,NomVaca.NumNomVaca , DetalleNomVaca.Inss, DetalleNomVaca.Ir,Empleado.CodEmpleado1,Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento, DetalleNomVaca.AdelantoVacaciones AS Adelanto13vo,DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones AS MontoPagar,DetalleNomVaca.TotalDevengado, Historico.FechaContrato, Empleado.TarifaHoraria, DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss AS TotalDeducir, NomVaca.CodTipoNomina FROM  NomVaca INNER JOIN  Empleado INNER JOIN DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON " & _
                    "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado WHERE     (NomVaca.NumNomVaca = " & NumNomina & ") AND (NomVaca.CodTipoNomina = '" & CodTipoNomina & "') AND (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1 "
End If
'Nom13voMoises.DataControl1.Source = SQLReportes
'Nom13voMoises.ImgLogo.Picture = LoadPicture(RutaLogo)
'Nom13voMoises.Show 1
'Exit Sub

End If
        
      ArepNomVacaciones.Label5.Caption = "Nomina de Vacaciones"
      ArepNomVacaciones.DataControl1.Source = SQlReportes
      ArepNomVacaciones.ImgLogo.Picture = LoadPicture(RutaLogo)

'      Nom13vo.Show 1
'        Dim rpt As Object
'        Dim fPreview As New FrmPreview
        
             Set rpt = New ArepNomVacaciones
             rpt.DataControl1.ConnectionString = ConexionReporte
             rpt.DataControl1.Source = SQlReportes
             fPreview.RunReport rpt
        
        
             fPreview.Show 1
           
           
      ArepPersecionesVaca.lblTitulo.Caption = Titulo
      ArepPersecionesVaca.LblSubtitulo.Caption = SubTitulo
      ArepPersecionesVaca.ImgLogo.Picture = LoadPicture(RutaLogo)
      ArepPersecionesVaca.Label5.Caption = "Nomina Vacaciones"
      
      ArepPersecionesVaca.LblFecha.Caption = "Desde " + Format(FechaIni, "mm/dd/yyyy") + " Hasta " + Format(FechaFin, "mm/dd/yyyy")
      ArepPersecionesVaca.LblFechaHoy = Format(Now, "dddddd")
      ArepPersecionesVaca.DataControl1.ConnectionString = ConexionReporte
      
        If NombreNomina <> "Administracion" Then
              
         If Me.ChkRestar.Value = 1 Then
               SQlReportes = "SELECT  NomVaca.NumNomVaca AS NumNomVaca, SUM(DetalleNomVaca.Inss) AS Inss,SUM(DetalleNomVaca.Ir) AS Ir, MAX(Empleado.CodEmpleado1) AS CodEmpleado1,MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, SUM(DetalleNomVaca.SalarioMensual) AS SalarioMensual, SUM(DetalleNomVaca.DiasAPagar) AS DiasAPagar, SUM(DetalleNomVaca.DiasDescuento) AS DiasDescuento, SUM(DetalleNomVaca.AdelantoVacaciones) AS Adelanto13vo, SUM(DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss - DetalleNomVaca.Ir) " & _
                             "AS MontoPagar, SUM(DetalleNomVaca.TotalDevengado) AS TotalDevengado, MAX(Historico.FechaContrato) AS FechaContrato, MAX(Empleado.TarifaHoraria) AS TarifaHoraria, SUM(DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir) AS TotalDeducir, MAX(NomVaca.CodTipoNomina) AS CodTipoNomina  FROM  NomVaca INNER JOIN Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON  " & _
                             "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado GROUP BY NomVaca.NumNomVaca HAVING (SUM(DetalleNomVaca.DiasAPagar) <> 0) AND (MAX(NomVaca.CodTipoNomina) = '" & CodTipoNomina & "') AND (NomVaca.NumNomVaca =" & NumNomina & ") ORDER BY MAX(Empleado.CodEmpleado1)"
         Else
               SQlReportes = "SELECT  NomVaca.NumNomVaca AS NumNomVaca, SUM(DetalleNomVaca.Inss) AS Inss, SUM(DetalleNomVaca.Ir) AS Ir, MAX(Empleado.CodEmpleado1) AS CodEmpleado1,MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, SUM(DetalleNomVaca.SalarioMensual) AS SalarioMensual, SUM(DetalleNomVaca.DiasAPagar) AS DiasAPagar, SUM(DetalleNomVaca.DiasDescuento) AS DiasDescuento, SUM(DetalleNomVaca.AdelantoVacaciones) AS Adelanto13vo, SUM(DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss) AS MontoPagar, SUM(DetalleNomVaca.TotalDevengado) AS TotalDevengado, MAX(Historico.FechaContrato) AS FechaContrato, MAX(Empleado.TarifaHoraria) AS TarifaHoraria, SUM(DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss) AS TotalDeducir, MAX(NomVaca.CodTipoNomina) AS CodTipoNomina  FROM  NomVaca INNER JOIN Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON  " & _
                             "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado GROUP BY NomVaca.NumNomVaca HAVING (SUM(DetalleNomVaca.DiasAPagar) <> 0) AND (MAX(NomVaca.CodTipoNomina) = '" & CodTipoNomina & "') AND (NomVaca.NumNomVaca =" & NumNomina & ") ORDER BY MAX(Empleado.CodEmpleado1)"
         End If
        Else
        
        If Me.ChkRestar.Value = 1 Then
               SQlReportes = "SELECT  NomVaca.NumNomVaca AS NumNomVaca, SUM(DetalleNomVaca.Inss) AS Inss, SUM(DetalleNomVaca.Ir) AS Ir, MAX(Empleado.CodEmpleado1) AS CodEmpleado1,MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, SUM(DetalleNomVaca.SalarioMensual) AS SalarioMensual, SUM(DetalleNomVaca.DiasAPagar) AS DiasAPagar, SUM(DetalleNomVaca.DiasDescuento) AS DiasDescuento, SUM(DetalleNomVaca.AdelantoVacaciones) AS Adelanto13vo, SUM(DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones - DetalleNomVaca.Inss -DetalleNomVaca.Ir) " & _
                             " AS MontoPagar, SUM(DetalleNomVaca.TotalDevengado) AS TotalDevengado, MAX(Historico.FechaContrato) AS FechaContrato, MAX(Empleado.TarifaHoraria) AS TarifaHoraria, SUM(DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss + DetalleNomVaca.Ir) AS TotalDeducir, MAX(NomVaca.CodTipoNomina) AS CodTipoNomina  FROM  NomVaca INNER JOIN Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON  " & _
                             "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado  GROUP BY NomVaca.NumNomVaca HAVING (SUM(DetalleNomVaca.DiasAPagar) <> 0) AND (MAX(NomVaca.CodTipoNomina) = '" & CodTipoNomina & "') AND (NomVaca.NumNomVaca =" & NumNomina & ") ORDER BY MAX(Empleado.CodEmpleado1)"
        
        Else
               SQlReportes = "SELECT  NomVaca.NumNomVaca AS NumNomVaca, SUM(DetalleNomVaca.Inss) AS Inss, SUM(DetalleNomVaca.Ir) AS Ir, MAX(Empleado.CodEmpleado1) AS CodEmpleado1,MAX(Empleado.Nombre1 + N' ' + Empleado.Nombre2 + N' ' + Empleado.Apellido1 + N' ' + Empleado.Apellido2) AS Nombres, SUM(DetalleNomVaca.SalarioMensual) AS SalarioMensual, SUM(DetalleNomVaca.DiasAPagar) AS DiasAPagar, SUM(DetalleNomVaca.DiasDescuento) AS DiasDescuento, SUM(DetalleNomVaca.AdelantoVacaciones) AS Adelanto13vo, SUM(DetalleNomVaca.TotalDevengado - DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss) AS MontoPagar, SUM(DetalleNomVaca.TotalDevengado) AS TotalDevengado, MAX(Historico.FechaContrato) AS FechaContrato, MAX(Empleado.TarifaHoraria) AS TarifaHoraria, SUM(DetalleNomVaca.AdelantoVacaciones + DetalleNomVaca.Inss) AS TotalDeducir, MAX(NomVaca.CodTipoNomina) AS CodTipoNomina  FROM  NomVaca INNER JOIN Empleado INNER JOIN  DetalleNomVaca ON Empleado.CodEmpleado = DetalleNomVaca.CodEmpleado ON  " & _
                             "NomVaca.NumNomVaca = DetalleNomVaca.NumNomVaca INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado  GROUP BY NomVaca.NumNomVaca HAVING (SUM(DetalleNomVaca.DiasAPagar) <> 0) AND (MAX(NomVaca.CodTipoNomina) = '" & CodTipoNomina & "') AND (NomVaca.NumNomVaca =" & NumNomina & ") ORDER BY MAX(Empleado.CodEmpleado1)"
        End If
      End If
      
'      ArepPersecionesVaca.Label5.Caption = "Nomina de Vacaciones"
      ArepPersecionesVaca.DataControl1.Source = SQlReportes
'      ArepPersecionesVaca.ImgLogo.Picture = LoadPicture(RutaLogo)
      ArepPersecionesVaca.Show 1
        
'      Nom13vo.Label5.Caption = "Nomina de Vacaciones"
'      Nom13vo.DataControl1.Source = SQlReportes
'      Nom13vo.ImgLogo.Picture = LoadPicture(RutaLogo)
'      Nom13vo.Show 1

Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Command4_Click()
On Error GoTo TipoErrs
Dim Espacio As String
Espacio = " "

NumNomina = Me.AdoNomina13vo.Recordset("NumNom13Mes")
CodTipoNomina = Me.AdoNomina13vo.Recordset("CodTipoNomina")
FechaIni = Me.AdoNomina13vo.Recordset("FechaIni")
FechaFin = Me.AdoNomina13vo.Recordset("FechaFin")


      Nom13vo.lblTitulo.Caption = Titulo
      Nom13vo.LblSubtitulo.Caption = SubTitulo
      Nom13vo.ImgLogo.Picture = LoadPicture(RutaLogo)
      
      Nom13vo.LblFecha.Caption = "Desde " + Format(FechaIni, "mm/dd/yyyy") + " Hasta " + Format(FechaFin, "mm/dd/yyyy")
      Nom13vo.LblFechaHoy = Format(Now, "dddddd")
      Nom13vo.DataControl1.ConnectionString = ConexionReporte
      SQlReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, (DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo) AS MontoPagar, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, Cargo.Cargo, ([DetalleNom13Mes].SalarioAPagar) AS TotalDevengado, Empleado.CodEmpleado1 FROM Nom13Mes INNER JOIN ((Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNomina & ")) ORDER BY Empleado.CodEmpleado1"
      Nom13vo.DataControl1.Source = SQlReportes
      Nom13vo.ImgLogo.Picture = LoadPicture(RutaLogo)
      Nom13vo.Show 1

Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub Command5_Click()
On Error GoTo TipoErrs
Dim Espacio As String
Dim FechaIni As Date, FechaFin As Date
Dim NumNomina As Integer, CodTipoNomina As String

Espacio = " "

 
NumNomina = Me.AdoNomina13vo.Recordset("NumNom13Mes")
CodTipoNomina = Me.AdoNomina13vo.Recordset("CodTipoNomina")
FechaIni = Me.AdoNomina13vo.Recordset("FechaIni")
FechaFin = Me.AdoNomina13vo.Recordset("FechaFin")

SQlReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, (DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo) AS MontoPagar, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, Cargo.Cargo, ([DetalleNom13Mes].SalarioAPagar) AS TotalDevengado, Empleado.CodEmpleado1 FROM Nom13Mes INNER JOIN ((Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNomina & ")) ORDER BY Empleado.CodEmpleado1"
'SQLReportes = "SELECT Nom13Mes.NumNom13Mes, DetalleNom13Mes.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.DiasAPagar, DetalleNom13Mes.Adelanto13vo, (DetalleNom13Mes.SalarioAPagar - DetalleNom13Mes.Adelanto13vo) AS MontoPagar, [Nombre1]+ '" & Espacio & "'+[Nombre2]+'" & Espacio & "'+[Apellido1]+'" & Espacio & "'+ [Apellido2] AS Nombres, Cargo.Cargo, ([DetalleNom13Mes].SalarioAPagar) AS TotalDevengado, Empleado.CodEmpleado1 FROM Nom13Mes INNER JOIN ((Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) INNER JOIN DetalleNom13Mes ON Empleado.CodEmpleado = DetalleNom13Mes.CodEmpleado) ON Nom13Mes.NumNom13Mes = DetalleNom13Mes.NumNom13Mes Where (((Nom13Mes.NumNom13Mes) = " & NumNomina & ")) ORDER BY Empleado.CodEmpleado1"
ArepColilla13vo.AdoColillas.Source = SQlReportes
ArepColilla13vo.LblTipo.Caption = CodTipoNomina
ArepColilla13vo.LblPeriodos.Caption = "Desde   " & FechaIni & " Hasta    " & FechaFin
ArepColilla13vo.LblPeriodo.Caption = Format(FechaIni, "dddddd") & "   Hasta   " & Format(FechaFin, "dddddd")
ArepColilla13vo.lblTitulo.Caption = Titulo
ArepColilla13vo.AdoColillas.ConnectionString = ConexionReporte
ArepColilla13vo.Show 1
Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub DbgrNominas_Click()
Dim SqlNominas As String
Dim SqlDetalleNominas As String, SqlDetalleNominasTotal As String
Dim NumNominas As Long

NumNominas = Me.AdoNominasVaca.Recordset("NumNomVaca")

SqlDetalleNominas = "SELECT Empleado.CodEmpleado1, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento,DetalleNomVaca.AdelantoVacaciones , DetalleNomVaca.Inss FROM DetalleNomVaca INNER JOIN Empleado ON DetalleNomVaca.CodEmpleado = Empleado.CodEmpleado Where (NumNomVaca = " & NumNominas & ") Order by CodEmpleado1"

'SqlDetalleNominas = "SELECT CodEmpleado, SalarioMensual, DiasAPagar, DiasDescuento, AdelantoVacaciones, Inss From DetalleNomVaca Where (NumNomVaca = " & NumNominas & ")"
Me.AdoDetalleNominaVaca.RecordSource = SqlDetalleNominas
Me.AdoDetalleNominaVaca.Refresh

SqlDetalleNominasTotal = "SELECT MAX(CodEmpleado) AS CodEmpleado, SUM(SalarioMensual) AS SalarioMensual, SUM(DiasAPagar) AS DiasAPagar, SUM(DiasDescuento) AS DiasDescuentos, SUM(AdelantoVacaciones) AS AdelantoVacaciones, SUM(Inss) AS Inss From DetalleNomVaca Where (NumNomVaca = " & NumNominas & ")"
Me.AdoDetalleTotal.RecordSource = SqlDetalleNominasTotal
Me.AdoDetalleTotal.Refresh
If Not Me.AdoDetalleTotal.Recordset.EOF Then
   Me.TxtSalarioMensual.Text = Format(Me.AdoDetalleTotal.Recordset("SalarioMensual"), "##,##0.00")
   Me.TxtDiasPagar.Text = Format(Me.AdoDetalleTotal.Recordset("DiasAPagar"), "##,##0.00")
   Me.TxtDiasDescuentos.Text = Format(Me.AdoDetalleTotal.Recordset("DiasDescuentos"), "##,##0.00")
   Me.TxtAdelantos.Text = Format(Me.AdoDetalleTotal.Recordset("AdelantoVacaciones"), "##,##0.00")
   Me.TxtInss.Text = Format(Me.AdoDetalleTotal.Recordset("Inss"), "##,##0.00")
End If


End Sub

Private Sub DbgrNominas13vo_Click()
  Dim SqlDetalleNomina As String, NumNomina As Double
  Dim SQlConsulta As String
  
  NumNomina = Me.AdoNomina13vo.Recordset("NumNom13Mes")
  SqlDetalleNomina = "SELECT Empleado.CodEmpleado1, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.SalarioAPagar, DetalleNom13Mes.DiasAPagar,DetalleNom13Mes.Adelanto13vo FROM  DetalleNom13Mes INNER JOIN Empleado ON DetalleNom13Mes.CodEmpleado = Empleado.CodEmpleado Where (DetalleNom13Mes.NumNom13Mes = " & NumNomina & " ) ORDER BY Empleado.CodEmpleado1"
  Me.AdoDetalleNomina13vo.RecordSource = SqlDetalleNomina
  Me.AdoDetalleNomina13vo.Refresh
  
  SQlConsulta = "SELECT MAX(Empleado.CodEmpleado1) AS CodEmpleado, SUM(DetalleNom13Mes.SalarioMensual) AS SalarioMensual,SUM(DetalleNom13Mes.SalarioAPagar) AS SalarioAPagar, SUM(DetalleNom13Mes.DiasAPagar) AS DiasAPagar, SUM(DetalleNom13Mes.Adelanto13vo) AS Adelanto13vo FROM  DetalleNom13Mes INNER JOIN  Empleado ON DetalleNom13Mes.CodEmpleado = Empleado.CodEmpleado  Where (DetalleNom13Mes.NumNom13Mes = " & NumNomina & ") ORDER BY MAX(Empleado.CodEmpleado1)"
  Me.AdoConsulta.RecordSource = SQlConsulta
  Me.AdoConsulta.Refresh
  If Not Me.AdoConsulta.Recordset.EOF Then
   Me.TxtSalarioMensual13vo.Text = Format(Me.AdoConsulta.Recordset("SalarioMensual"), "##,##0.00")
   Me.TxtDiasPagar13vo.Text = Format(Me.AdoConsulta.Recordset("DiasAPagar"), "##,##0.00")
   Me.TxtAdelanto13vo.Text = Format(Me.AdoConsulta.Recordset("Adelanto13vo"), "##,##0.00")
   Me.TxtSalarioPagar13vo.Text = Format(Me.AdoConsulta.Recordset("SalarioAPagar"), "##,##0.00")
  
  End If
  
End Sub

Private Sub Form_Load()
Me.DbgrdetalleNominas.EvenRowStyle.BackColor = &HC0FFFF
 Me.DbgrdetalleNominas.OddRowStyle.BackColor = &HFFFFFF
 Me.DbgrdetalleNominas.AlternatingRowStyle = True
 
Me.DbgrNominas.EvenRowStyle.BackColor = &HC0FFFF
 Me.DbgrNominas.OddRowStyle.BackColor = &HFFFFFF
 Me.DbgrNominas.AlternatingRowStyle = True
 
 Me.DbgrdetalleNominas13vo.EvenRowStyle.BackColor = &HC0FFFF
 Me.DbgrdetalleNominas13vo.OddRowStyle.BackColor = &HFFFFFF
 Me.DbgrdetalleNominas13vo.AlternatingRowStyle = True
 
Me.DbgrNominas13vo.EvenRowStyle.BackColor = &HC0FFFF
 Me.DbgrNominas13vo.OddRowStyle.BackColor = &HFFFFFF
 Me.DbgrNominas13vo.AlternatingRowStyle = True
 
 With Me.DtaTipoNominas
   .ConnectionString = Conexion
   .RecordSource = "TipoNomina"
   .Refresh
 End With
 
  With Me.AdoConsulta
   .ConnectionString = Conexion
 End With
 
 With Me.AdoDetalleNomina13vo
   .ConnectionString = Conexion
 End With
 
 With Me.AdoNomina13vo
   .ConnectionString = Conexion
 End With
 
  With Me.AdoDetalleTotal
   .ConnectionString = Conexion
 End With
 
 With Me.AdoNominasVaca
   .ConnectionString = Conexion
 End With
 
  With Me.AdoDetalleNominaVaca
   .ConnectionString = Conexion
 End With
 
Me.AdoNominasVaca.RecordSource = "SELECT * From NomVaca"
Me.AdoNominasVaca.Refresh

Me.AdoDetalleNominaVaca.RecordSource = "SELECT Empleado.CodEmpleado1, DetalleNomVaca.SalarioMensual, DetalleNomVaca.DiasAPagar, DetalleNomVaca.DiasDescuento,DetalleNomVaca.AdelantoVacaciones , DetalleNomVaca.Inss FROM DetalleNomVaca INNER JOIN Empleado ON DetalleNomVaca.CodEmpleado = Empleado.CodEmpleado "
Me.AdoDetalleNominaVaca.Refresh

Me.AdoNomina13vo.RecordSource = "SELECT * From Nom13Mes"
Me.AdoNomina13vo.Refresh

Me.AdoDetalleNomina13vo.RecordSource = "SELECT Empleado.CodEmpleado1, DetalleNom13Mes.SalarioMensual, DetalleNom13Mes.SalarioAPagar, DetalleNom13Mes.DiasAPagar,DetalleNom13Mes.Adelanto13vo FROM DetalleNom13Mes INNER JOIN Empleado ON DetalleNom13Mes.CodEmpleado = Empleado.CodEmpleado "
Me.AdoDetalleNomina13vo.Refresh


 
 End Sub

Private Sub SmartButton1_Click()
Dim sql As String, NumeroNomina As Integer, CodTipoNomina As String

NumeroNomina = Me.AdoNomina13vo.Recordset("NumNom13Mes")
CodTipoNomina = Me.AdoNominasVaca.Recordset("CodTipoNomina")
 
FechaIni = Me.AdoNominasVaca.Recordset("FechaIni")
FechaFin = Me.AdoNominasVaca.Recordset("FechaFin")

Me.DtaTipoNominas.Refresh
Do While Not Me.DtaTipoNominas.Recordset.EOF
If Me.DtaTipoNominas.Recordset("CodTipoNomina") = CodTipoNomina Then
   NombreNomina = Me.DtaTipoNominas.Recordset("Nomina")
   Exit Do
End If
DtaTipoNominas.Recordset.MoveNext
Loop


sql = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres," & vbLf
sql = sql & "HistorialSalarioMes.Fechaini , HistorialSalarioMes.Fechafin, HistorialSalarioMes.Enero, HistorialSalarioMes.Febrero, HistorialSalarioMes.Marzo, " & vbLf
sql = sql & "HistorialSalarioMes.Abril , HistorialSalarioMes.Mayo, HistorialSalarioMes.Junio, HistorialSalarioMes.Julio, HistorialSalarioMes.Agosto, " & vbLf
sql = sql & "HistorialSalarioMes.Septiembre , HistorialSalarioMes.Octubre, HistorialSalarioMes.Noviembre, HistorialSalarioMes.Diciembre, " & vbLf
sql = sql & "HistorialSalarioMes.NumNomina " & vbLf
sql = sql & "FROM HistorialSalarioMes INNER JOIN" & vbLf
sql = sql & "Empleado ON HistorialSalarioMes.CodEmpleado = Empleado.CodEmpleado" & vbLf
sql = sql & "Where (HistorialSalarioMes.NumNomina = " & NumeroNomina & ")  ORDER BY Empleado.CodEmpleado1"

FrmSalarioHistorial.AdoSalarios.RecordSource = sql
FrmSalarioHistorial.AdoSalarios.Refresh

FrmSalarioHistorial.TxtNumNom13.Text = Me.AdoNomina13vo.Recordset("NumNom13Mes")
FrmSalarioHistorial.Dbgr13Mes.Columns(0).Width = 1100
FrmSalarioHistorial.Dbgr13Mes.Columns(1).Width = 3000
FrmSalarioHistorial.Dbgr13Mes.Columns(4).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(5).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(6).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(7).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(8).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(9).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(10).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(11).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(12).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(13).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(14).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(15).Width = 1000
FrmSalarioHistorial.Dbgr13Mes.Columns(16).Visible = False
FrmSalarioHistorial.Dbgr13Mes.Columns(2).Visible = False
FrmSalarioHistorial.Dbgr13Mes.Columns(3).Visible = False
FrmSalarioHistorial.DBCNominas.Text = NombreNomina
FrmSalarioHistorial.TxtFFIN13.Value = FechaIni
FrmSalarioHistorial.TxtFINI13.Value = FechaFin

  FrmSalarioHistorial.Dbgr13Mes.Columns(2).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(3).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(4).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(5).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(6).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(7).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(8).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(9).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(10).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(11).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(12).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(13).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(14).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(15).NumberFormat = "##,##0.00"
  FrmSalarioHistorial.Dbgr13Mes.Columns(16).NumberFormat = "##,##0.00"

FrmSalarioHistorial.SSTab1.TabEnabled(0) = False
FrmSalarioHistorial.SSTab1.Tab = 1
FrmSalarioHistorial.Show 1

End Sub

Private Sub SmartButton2_Click()
Dim sql As String, NumeroNomina As Integer
Dim NombreNomina As String

NumNomina = Me.AdoNominasVaca.Recordset("NumNomVaca")
CodTipoNomina = Me.AdoNominasVaca.Recordset("CodTipoNomina")
 
FechaIni = Me.AdoNominasVaca.Recordset("FechaIni")
FechaFin = Me.AdoNominasVaca.Recordset("FechaFin")

Me.DtaTipoNominas.Refresh
Do While Not Me.DtaTipoNominas.Recordset.EOF
If Me.DtaTipoNominas.Recordset("CodTipoNomina") = CodTipoNomina Then
   NombreNomina = Me.DtaTipoNominas.Recordset("Nomina")
   Exit Do
End If
DtaTipoNominas.Recordset.MoveNext
Loop


sql = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres," & vbLf
sql = sql & "HistorialSalarioMes.Fechaini , HistorialSalarioMes.Fechafin, HistorialSalarioMes.Enero, HistorialSalarioMes.Febrero, HistorialSalarioMes.Marzo, " & vbLf
sql = sql & "HistorialSalarioMes.Abril , HistorialSalarioMes.Mayo, HistorialSalarioMes.Junio, HistorialSalarioMes.Julio, HistorialSalarioMes.Agosto, " & vbLf
sql = sql & "HistorialSalarioMes.Septiembre , HistorialSalarioMes.Octubre, HistorialSalarioMes.Noviembre, HistorialSalarioMes.Diciembre, " & vbLf
sql = sql & "HistorialSalarioMes.NumNomina " & vbLf
sql = sql & "FROM HistorialSalarioMes INNER JOIN" & vbLf
sql = sql & "Empleado ON HistorialSalarioMes.CodEmpleado = Empleado.CodEmpleado" & vbLf
sql = sql & "Where (HistorialSalarioMes.NumNomina = " & NumNomina & ")AND (HistorialSalarioMes.Tipo = 'Vacaciones') ORDER BY Empleado.CodEmpleado1"

FrmSalarioHistorial.AdoSalarioVacaciones.RecordSource = sql
'InputBox "", "", FrmSalarioHistorial.AdoSalarioVacaciones.RecordSource
FrmSalarioHistorial.AdoSalarioVacaciones.Refresh
FrmSalarioHistorial.SSTab1.TabEnabled(1) = False
FrmSalarioHistorial.SSTab1.Tab = 0

FrmSalarioHistorial.TxtNumNom13.Text = NumNomina
FrmSalarioHistorial.DbgrVacaciones.Columns(0).Width = 1100
FrmSalarioHistorial.DbgrVacaciones.Columns(1).Width = 3000
FrmSalarioHistorial.DbgrVacaciones.Columns(4).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(5).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(6).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(7).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(8).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(9).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(10).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(11).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(12).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(13).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(14).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(15).Width = 1000
FrmSalarioHistorial.DbgrVacaciones.Columns(16).Visible = False
FrmSalarioHistorial.DbgrVacaciones.Columns(2).Visible = False
FrmSalarioHistorial.DbgrVacaciones.Columns(3).Visible = False
FrmSalarioHistorial.DbgrVacaciones.Columns(4).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(5).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(6).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(7).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(8).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(9).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(10).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(11).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(12).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(13).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(14).NumberFormat = "##,##0.00"
FrmSalarioHistorial.DbgrVacaciones.Columns(15).NumberFormat = "##,##0.00"


FrmSalarioHistorial.DBCNominas.Text = NombreNomina
FrmSalarioHistorial.TxtFFinVaca.Value = FechaFin
FrmSalarioHistorial.TxtFINIVaca.Value = FechaIni
FrmSalarioHistorial.TxtNumNomVaca.Text = NumNomina
FrmSalarioHistorial.SSTab1.Tab = 0
FrmSalarioHistorial.Show 1
End Sub

Public Function FechaPeriodo(FechaIni As Date, FechaFin As Date, CodTipoNomina As String)
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer

Mes1 = Month(FechaIni)
Año1 = Year(FechaIni)
Mes2 = Month(FechaFin)
Año2 = Year(FechaFin)

Mes1 = Format(Mes1, "0#")
Mes2 = Format(Mes2, "0#")


MDIPrimero.DtaConsulta.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año1 & ") AND (mes = '" & Mes1 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
MDIPrimero.DtaConsulta.Refresh
 If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
   FechaIniVaca = MDIPrimero.DtaConsulta.Recordset("Inicio")
 End If

MDIPrimero.DtaConsulta.RecordSource = "SELECT año, Periodo, mes, CodTipoNomina, Inicio, Final, Actual, Calculada, NumNomina From Fecha_Planilla WHERE  (año = " & Año2 & ") AND (mes = '" & Mes2 & "') AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY Periodo"
MDIPrimero.DtaConsulta.Refresh
 If Not MDIPrimero.DtaConsulta.Recordset.EOF Then
   MDIPrimero.DtaConsulta.Recordset.MoveLast
   FechaFinVaca = MDIPrimero.DtaConsulta.Recordset("Final")
 End If
End Function
