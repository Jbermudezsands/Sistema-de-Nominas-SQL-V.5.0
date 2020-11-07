VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form FrmReembolso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reembolso de Vacaciones"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8040
   Begin TrueOleDBGrid70.TDBGrid TDBGridReembolso 
      Bindings        =   "FrmReembolso.frx":0000
      Height          =   3495
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6165
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=224,.bold=0,.fontsize=825,.italic=0"
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
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   7815
      Begin VB.TextBox TxtNumNomina 
         Height          =   375
         Left            =   5880
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TxtCodEmpleado 
         Height          =   285
         Left            =   5880
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TxtCodNomina 
         Height          =   285
         Left            =   4080
         TabIndex        =   10
         Text            =   "0"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo DBCNominas 
         Bindings        =   "FrmReembolso.frx":001B
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nomina"
         Text            =   "Listado de Nominas"
      End
      Begin XtremeSuiteControls.PushButton CmdAgregar 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Agregar"
         ForeColor       =   0
         Enabled         =   0   'False
         Appearance      =   6
         Picture         =   "FrmReembolso.frx":0037
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdQuitar 
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   960
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Quitar"
         ForeColor       =   0
         Enabled         =   0   'False
         Appearance      =   6
         Picture         =   "FrmReembolso.frx":2409
         ImageAlignment  =   0
      End
      Begin VB.Label LblDescripcion 
         Caption         =   "Nomina Vacaciones:"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   645
         Width           =   7455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Listado de Nominas"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   -120
      ScaleHeight     =   1215
      ScaleWidth      =   13095
      TabIndex        =   0
      Top             =   -120
      Width           =   13095
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Reembolso de Vacaciones"
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
         Left            =   3000
         TabIndex        =   1
         Top             =   480
         Width           =   3360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   13080
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Image Image1 
         Height          =   1320
         Left            =   0
         Picture         =   "FrmReembolso.frx":4617
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1965
      End
   End
   Begin XtremeSuiteControls.PushButton CmdSalir 
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   6360
      Width           =   1455
      _Version        =   786432
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salir"
      ForeColor       =   0
      Appearance      =   6
      Picture         =   "FrmReembolso.frx":D0CF
      ImageAlignment  =   0
   End
   Begin MSAdodcLib.Adodc DtaTipoNomina 
      Height          =   375
      Left            =   600
      Top             =   7920
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
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   600
      Top             =   8400
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
   Begin MSAdodcLib.Adodc AdoReembolso 
      Height          =   375
      Left            =   3480
      Top             =   8160
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
      Caption         =   "AdoReembolos"
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
Attribute VB_Name = "FrmReembolso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
Dim CodEmpleado As Double, NumNomina As Double, Respuesta As Double
Dim Mensaje As String

QueProducto = "CodigoEmpleado"
FrmConsulta.Show 1
Me.TxtCodEmpleado.Text = FrmConsulta.CodigoEmpleado
CodEmpleado = Me.TxtCodEmpleado.Text
NumNomina = Me.TxtNumNomina.Text


'/////////////////////////BUSCO SI EXISTE EN LA NOMINA VACACIONES //////////////////////////////
Me.DtaConsulta.RecordSource = "SELECT  * From DetalleNomVaca Where (CodEmpleado = " & CodEmpleado & ") And (NumNomVaca = " & NumNomina & ") "
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
  Mensaje = "Ya Existe este empleado, para la NOMINA VACACIONES!!! " & vbLf
  Mensaje = Mensaje & "    DESEA AGREGARLO AL REEMBOLSO?"
  Respuesta = MsgBox(Mensaje, vbYesNo, "Zeus Nominas")
  If Respuesta = 7 Then
    Exit Sub
  End If
End If


'/////////////////////////BUSCO SI EXISTE EN LOS REEMBOLSOS //////////////////////////////
Me.DtaConsulta.RecordSource = "SELECT  * From Reembolso WHERE (NumNomina = " & NumNomina & ") AND (CodEmpleado = " & CodEmpleado & ")"
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
  MsgBox "Ya Existe este empleado, para Reembolsos!!!", vbInformation, "Zeus Nominas"
  Exit Sub
End If

  Me.DtaConsulta.Recordset.AddNew
  Me.DtaConsulta.Recordset("NumNomina") = Me.TxtNumNomina.Text
  Me.DtaConsulta.Recordset("CodEmpleado") = Me.TxtCodEmpleado.Text
  Me.DtaConsulta.Recordset("Monto") = 0
  Me.DtaConsulta.Recordset.Update
  
  Me.AdoReembolso.Refresh
   Me.TDBGridReembolso.Columns(0).Width = 1100
   Me.TDBGridReembolso.Columns(0).Locked = True
   Me.TDBGridReembolso.Columns(1).Width = 1200
   Me.TDBGridReembolso.Columns(1).Locked = True
   Me.TDBGridReembolso.Columns(2).Width = 3000
   Me.TDBGridReembolso.Columns(2).Locked = True
   Me.TDBGridReembolso.Columns(4).Visible = False

End Sub

Private Sub cmdQuitar_Click()
Dim NumNomina As Double, CodEmpleado As String

  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando a el Empleado: " & CodEmpleado)
   If Respuesta = 7 Then
    Exit Sub
   End If

NumNomina = Me.TxtNumNomina.Text
CodEmpleado = Me.TDBGridReembolso.Columns(4).Text

Me.DtaConsulta.RecordSource = "SELECT  * From Reembolso WHERE (NumNomina = " & NumNomina & ") AND (CodEmpleado = " & CodEmpleado & ")"
Me.DtaConsulta.Refresh
If Not Me.DtaConsulta.Recordset.EOF Then
 Set Ejecutar = New ADODB.Connection
 Ejecutar.ConnectionString = Conexion
 Ejecutar.Open
 Ejecutar.Execute "DELETE FROM Reembolso WHERE (NumNomina = " & NumNomina & ") AND (CodEmpleado = " & CodEmpleado & ")"
 
   Me.AdoReembolso.Refresh
   Me.TDBGridReembolso.Columns(0).Width = 1100
   Me.TDBGridReembolso.Columns(0).Locked = True
   Me.TDBGridReembolso.Columns(1).Width = 1200
   Me.TDBGridReembolso.Columns(1).Locked = True
   Me.TDBGridReembolso.Columns(2).Width = 3000
   Me.TDBGridReembolso.Columns(2).Locked = True
   Me.TDBGridReembolso.Columns(4).Visible = False
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub DBCNominas_Change()
Dim Mes1 As String, Mes2 As String, Año1 As Integer, Año2 As Integer
Dim CodTipoNomina As String, NumNominaVaca As Double, FechaIni As Date, FechaFin As Date

DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("nomina") = DBCNominas.Text Then
   CodTipoNomina = DtaTipoNomina.Recordset("CodTipoNomina")
   Me.TxtCodNomina.Text = CodTipoNomina
   Exit Do
End If
DtaTipoNomina.Recordset.MoveNext
Loop

Me.cmdAgregar.Enabled = False
Me.cmdQuitar.Enabled = False

'/////////busco si existen Nominas Activas en el Sistema//////////////////
 Me.DtaConsulta.RecordSource = "SELECT NomVaca.NumNomVaca, NomVaca.FechaAplica, NomVaca.FechaIni, NomVaca.FechaFin, NomVaca.Activa From NomVaca Where (((NomVaca.Activa) = 1))AND (CodTipoNomina = '" & CodTipoNomina & "') ORDER BY NomVaca.FechaFin"
 Me.DtaConsulta.Refresh
 If Not DtaConsulta.Recordset.EOF Then
   NumNominaVaca = Me.DtaConsulta.Recordset("NumNomVaca")
   Me.TxtNumNomina.Text = NumNominaVaca
   FechaIni = Me.DtaConsulta.Recordset("FechaIni")
   FechaFin = Me.DtaConsulta.Recordset("FechaFin")
   Me.cmdAgregar.Enabled = True
   Me.cmdQuitar.Enabled = True
   
   Me.AdoReembolso.RecordSource = "SELECT Reembolso.NumNomina, Empleado.CodEmpleado1 AS CodEmpleado, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Reembolso.Monto, Reembolso.CodEmpleado AS CodigoEmpleado FROM Reembolso INNER JOIN Empleado ON Reembolso.CodEmpleado = Empleado.CodEmpleado  " & _
                                  "Where (Reembolso.NumNomina = " & NumNominaVaca & ") ORDER BY CodEmpleado"
   Me.AdoReembolso.Refresh
   Me.TDBGridReembolso.Columns(0).Width = 1100
   Me.TDBGridReembolso.Columns(0).Locked = True
   Me.TDBGridReembolso.Columns(1).Width = 1200
   Me.TDBGridReembolso.Columns(1).Locked = True
   Me.TDBGridReembolso.Columns(2).Width = 3000
   Me.TDBGridReembolso.Columns(2).Locked = True
   Me.TDBGridReembolso.Columns(4).Visible = False
   
 Else
   Me.cmdAgregar.Enabled = False
   Me.cmdQuitar.Enabled = False
   
 End If

Me.LblDescripcion.Caption = "Nomina Vaciones:  " & NumNominaVaca & "       Desde " & FechaIni & "   Hasta   " & FechaFin


End Sub

Private Sub Form_Load()
'MDIPrimero.Skin1.ApplySkin hWnd
 Me.TDBGridReembolso.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.TDBGridReembolso.OddRowStyle.BackColor = &H80000005
 Me.TDBGridReembolso.AlternatingRowStyle = True
 
 
 
 With Me.DtaTipoNomina
   .ConnectionString = Conexion
   .RecordSource = "TipoNomina"
   .Refresh
End With

 With Me.DtaConsulta
   .ConnectionString = Conexion
 End With
 
  With Me.AdoReembolso
   .ConnectionString = Conexion
 End With
 
   Me.AdoReembolso.RecordSource = "SELECT Reembolso.NumNomina, Empleado.CodEmpleado1 AS CodEmpleado, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Reembolso.Monto, Reembolso.CodEmpleado AS CodigoEmpleado FROM Reembolso INNER JOIN Empleado ON Reembolso.CodEmpleado = Empleado.CodEmpleado  " & _
                                  "Where (Reembolso.NumNomina = -100 ) ORDER BY CodEmpleado"
   Me.AdoReembolso.Refresh
   Me.TDBGridReembolso.Columns(0).Width = 1100
   Me.TDBGridReembolso.Columns(0).Locked = True
   Me.TDBGridReembolso.Columns(1).Width = 1200
   Me.TDBGridReembolso.Columns(1).Locked = True
   Me.TDBGridReembolso.Columns(2).Width = 3000
   Me.TDBGridReembolso.Columns(2).Locked = True
   Me.TDBGridReembolso.Columns(4).Visible = False
 
End Sub
