VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmActivarNomina 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Activar Nominas"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   593
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   120
      Top             =   5280
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
      Left            =   4560
      Top             =   5880
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
      Left            =   4560
      Top             =   5280
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
      Left            =   240
      Top             =   5880
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
   Begin Project1.xp_canvas xp_canvas1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8493
      Caption         =   "Activar Nominas"
      Fixed_Single    =   -1  'True
      Begin TrueOleDBGrid70.TDBGrid DbgrTipoNominas 
         Bindings        =   "FrmActivarNomina.frx":0000
         Height          =   1935
         Left            =   240
         TabIndex        =   9
         Top             =   1080
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
      Begin Project1.xptopbuttons xptopbuttons1 
         Height          =   315
         Left            =   8520
         Top             =   75
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
      End
      Begin Project1.xphelp xphelp1 
         Height          =   315
         Left            =   8160
         Top             =   75
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
      End
      Begin VB.CommandButton CmdActivar 
         DownPicture     =   "FrmActivarNomina.frx":001C
         Height          =   375
         Left            =   5760
         Picture         =   "FrmActivarNomina.frx":1AFE
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton CmdSalir 
         DownPicture     =   "FrmActivarNomina.frx":3400
         Height          =   375
         Left            =   7200
         Picture         =   "FrmActivarNomina.frx":4EE2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4080
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker MtxtFechaini 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   3120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   16777215
         Format          =   51052545
         CurrentDate     =   37257
      End
      Begin MSComCtl2.DTPicker MtxtFecha 
         Height          =   375
         Left            =   6240
         TabIndex        =   1
         Top             =   3120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   16777215
         Format          =   51052545
         CurrentDate     =   37257
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Pago Sugerida"
         Height          =   255
         Left            =   4200
         TabIndex        =   8
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Activando Nóminas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   7
         Top             =   600
         Width           =   4455
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
         Left            =   2760
         TabIndex        =   6
         Top             =   3600
         Width           =   5775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio Período"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   3240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmActivarNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdActivar_Click()
Dim NumIni As Long
Dim NumFin As Long
Dim NumAIni As Long
Dim NumAFin As Long
Dim FechaTemp As Date
Dim SqlNominas As String
Dim CodTipoNomina As String
Dim Año As Integer, Periodo As Integer
Dim Fecha1 As Long
Dim Fecha2 As Long
On Error GoTo TipoErrs

   
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
'//////////////////////////////BUSCO LA NOMINA ACTIVA////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////
   Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
   Me.DtaConsulta.Refresh
   If Not DtaConsulta.Recordset.EOF Then
      Año = Me.DtaConsulta.Recordset("año")
      Periodo = Me.DtaConsulta.Recordset("Periodo")
      Fecha1 = Me.DtaConsulta.Recordset("Inicio")
      Fecha2 = Me.DtaConsulta.Recordset("Final")
'/////////////////////////BUSCO SI LA NOMINA ACTIVA ACTUAL YA ESTA CERRADA////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////
      Me.DtaConsulta.RecordSource = "SELECT NumNomina, CodTipoNomina, FechaNominaINI, FechaNomina, Activa, Procesada, Cerrada From Nomina WHERE     (FechaNominaINI = " & Fecha1 & ") AND (FechaNomina = " & Fecha2 & ") AND (Cerrada = 1) AND (CodTipoNomina = '" & CodTipoNomina & "')AND (FechaNomina BETWEEN " & Fecha1 & " AND " & Fecha2 & ")"
      Me.DtaConsulta.Refresh
'/////////SI NO ESTA CERRADA, DEJO LA ACTUAL///////////////////////////////////
      If DtaConsulta.Recordset.EOF Then
        'MsgBox "La fecha de esta Nomina ya fue Activada, Consulte a su Soperte", vbCritical
        'Exit Sub
      Else
'///////////////////SI ESTA CERRADA      ////////////////////////////////////////
'////////////////// BUSCO LA NOMINA SIGUIENTE/////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////
        Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (año = " & Año & ") AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Periodo + 1 & ") ORDER BY Periodo"
        Me.DtaConsulta.Refresh
        If Not DtaConsulta.Recordset.EOF Then
            Me.DtaConsulta.Recordset("Actual") = 1
            Me.DtaConsulta.Recordset.Update
'//////////BUSCO LA NOMINA ANTERIOR PARA DESACTIVARLA/////////////////
'////////////////////////////////////////////////////////////////////////
            Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (año = " & Año & ") AND (CodTipoNomina = '" & CodTipoNomina & "') AND (Periodo = " & Periodo & ") ORDER BY Periodo"
            Me.DtaConsulta.Refresh
            If Not DtaConsulta.Recordset.EOF Then
                 Me.DtaConsulta.Recordset("Actual") = 0
                 Me.DtaConsulta.Recordset.Update
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



SqlNominas = "SELECT Nomina.NumNomina, Nomina.CodTipoNomina, Nomina.FechaNominaINI, Nomina.FechaNomina From Nomina WHERE (Nomina.CodTipoNomina='" & CodTipoNomina & "') AND (Nomina.FechaNominaINI>= " & NumIni & " And Nomina.FechaNominaINI<= " & NumFin & " OR Nomina.FechaNomina>=" & NumIni & " And Nomina.FechaNomina<=" & NumFin & ")"
DtaNomina.RecordSource = SqlNominas
DtaNomina.Refresh

If Not DtaNomina.Recordset.EOF Then
   MsgBox "Existen nóminas que incluyen este período, no se puede activar esta nómina en el período especificado"
   Exit Sub
End If

DtaNomina.RecordSource = "Nomina"
DtaNomina.Refresh

DtaConsecutivos.Refresh


DtaNomina.Recordset.AddNew
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
DtaNomina.Recordset.Update

'DtaTipoNomina.Recordset.Edit
DtaTipoNomina.Recordset("Activa") = 1
DtaTipoNomina.Recordset.Update

DtaConsecutivos.Refresh
'DtaConsecutivos.Recordset.Edit
DtaConsecutivos.Recordset("nominas") = DtaConsecutivos.Recordset("nominas") + 1
DtaConsecutivos.Recordset.Update

Unload Me

Exit Sub
TipoErrs:
ControlErrores
Unload Me



End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DbgrTipoNominas_DblClick()
'On Error GoTo TipoErrs
Dim MiDiaSemana As Integer, Fecha1 As Long, Fecha2 As Long
Dim Año As Integer, Periodo As Integer

MiDiaSemana = Weekday(Now, vbMonday)
'MsgBox MiDiaSemana
'MsgBox ("#" + DtaTipoNomina.Recordset("Periodo") + "#" + "#" + "Semanal Sábados" + "#")

CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")

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
      Periodo = Me.DtaConsulta.Recordset("Periodo")
      Fecha1 = Me.DtaConsulta.Recordset("Inicio")
      Fecha2 = Me.DtaConsulta.Recordset("Final")
'////////////////////BUSCO SI LA NOMINA ACTUAL ESTA CERRADA/////////////
'////////////////////////////////////////////////////////////////////////
      Me.DtaConsulta.RecordSource = "SELECT NumNomina, CodTipoNomina, FechaNominaINI, FechaNomina, Activa, Procesada, Cerrada From Nomina WHERE (FechaNominaINI = " & Fecha1 & ") AND (FechaNomina = " & Fecha2 & ") AND (Cerrada = 1) AND (CodTipoNomina = '" & CodTipoNomina & "')AND (FechaNomina BETWEEN " & Fecha1 & " AND " & Fecha2 & ")"
      Me.DtaConsulta.Refresh
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
     MsgBox "No Existe Ninguna Nomina Activa", vbCritical
     Exit Sub
   End If
End If

If DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Viernes" Then
 Select Case MiDiaSemana
          Case 1
             MtxtFecha = Format(Now + 12, "dd/mm/yyyy")
          Case 2
             MtxtFecha = Format(Now + 11, "dd/mm/yyyy")
          Case 3
             MtxtFecha = Format(Now + 10, "dd/mm/yyyy")
          Case 4
             MtxtFecha = Format(Now + 9, "dd/mm/yyyy")
          Case 5
             MtxtFecha = Format(Now + 8, "dd/mm/yyyy")
          Case 6
             MtxtFecha = Format(Now + 7, "dd/mm/yyyy")
          Case 7
             MtxtFecha = Format(Now + 6, "dd/mm/yyyy")
    End Select
    MtxtFechaini = Str(CDate(MtxtFecha) - 14)
ElseIf DtaTipoNomina.Recordset("Periodo") = "Catorcenal los Sabados" Then

 Select Case MiDiaSemana
          Case 1
             MtxtFecha = Format(Now + 13, "dd/mm/yyyy")
          Case 2
             MtxtFecha = Format(Now + 12, "dd/mm/yyyy")
          Case 3
             MtxtFecha = Format(Now + 11, "dd/mm/yyyy")
          Case 4
             MtxtFecha = Format(Now + 10, "dd/mm/yyyy")
          Case 5
             MtxtFecha = Format(Now + 9, "dd/mm/yyyy")
          Case 6
             MtxtFecha = Format(Now + 8, "dd/mm/yyyy")
          Case 7
             MtxtFecha = Format(Now + 7, "dd/mm/yyyy")
    End Select
    MtxtFechaini = Str(CDate(MtxtFecha) - 14)
End If

If DtaTipoNomina.Recordset("Periodo") = "Quincenal" Then
'////////////////////////////////BUSCO LA NOMINA ACTIVA ACTUAL/////////////
'//////////////////////////////////////////////////////////////////////////////
   Me.DtaConsulta.RecordSource = "SELECT Periodo, año, mes, CodTipoNomina, Inicio, Final, Actual From Fecha_Planilla WHERE     (CodTipoNomina = '" & CodTipoNomina & "') AND (Actual = 1)"
   Me.DtaConsulta.Refresh
   If Not DtaConsulta.Recordset.EOF Then
      Año = Me.DtaConsulta.Recordset("año")
      Periodo = Me.DtaConsulta.Recordset("Periodo")
      Fecha1 = Me.DtaConsulta.Recordset("Inicio")
      Fecha2 = Me.DtaConsulta.Recordset("Final")
'////////////////////BUSCO SI LA NOMINA ACTUAL ESTA CERRADA/////////////
'////////////////////////////////////////////////////////////////////////
      Me.DtaConsulta.RecordSource = "SELECT NumNomina, CodTipoNomina, FechaNominaINI, FechaNomina, Activa, Procesada, Cerrada From Nomina WHERE     (FechaNominaINI = " & Fecha1 & ") AND (FechaNomina = " & Fecha2 & ") AND (Cerrada = 1) AND (CodTipoNomina = '" & CodTipoNomina & "')AND (FechaNomina BETWEEN " & Fecha1 & " AND " & Fecha2 & ")"
      Me.DtaConsulta.Refresh
'///////////////////////SI NO ESTA CERRADA,DEJO LA ACTUAL/////
'///////////////////////////////////////////////////////////////////
      If DtaConsulta.Recordset.EOF Then
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
        MtxtFecha = DateSerial(Year(Now), Month(Now) + 1, 0)
        MtxtFechaini = "1/" + Str(Month(Now)) + "/" + Str(Year(Now))
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
LblFechaLarga.Caption = "Pago: " + Format(MtxtFecha.Value, "Long Date")

Exit Sub
TipoErrs:
ControlErrores
Unload Me


End Sub

Private Sub Form_Load()
On Error GoTo TipoErrs

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

 Me.DbgrTipoNominas.EvenRowStyle.BackColor = &HC0FFFF
 Me.DbgrTipoNominas.OddRowStyle.BackColor = &H80000005
 Me.DbgrTipoNominas.AlternatingRowStyle = True

Dim SQLTipoNomina

SQLTipoNomina = "SELECT TipoNomina.CodTipoNomina, TipoNomina.Nomina, TipoNomina.Periodo, TipoNomina.UltFecha, TipoNomina.TipoPago, TipoNomina.Moneda, TipoNomina.MantValor, TipoNomina.Activa From TipoNomina WHERE (((TipoNomina.Activa)=0))"
DtaTipoNomina.RecordSource = SQLTipoNomina
DtaTipoNomina.Refresh

MtxtFechaini.Value = Now
MtxtFecha.Value = Now



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
