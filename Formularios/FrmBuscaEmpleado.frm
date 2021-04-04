VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmBuscaEmpleado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscando Empleado..."
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   Icon            =   "FrmBuscaEmpleado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   1080
      Top             =   5760
      Width           =   3495
      _ExtentX        =   6165
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
   Begin TrueOleDBGrid70.TDBGrid TDBGEmpleados 
      Bindings        =   "FrmBuscaEmpleado.frx":0CCA
      Height          =   1575
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2778
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=15,.bold=0,.fontsize=825,.italic=0"
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
   Begin MSDataListLib.DataCombo DBCCargo 
      Bindings        =   "FrmBuscaEmpleado.frx":0CE5
      Height          =   315
      Left            =   5400
      TabIndex        =   18
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Cargo"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DBCDepartamento 
      Bindings        =   "FrmBuscaEmpleado.frx":0CFC
      Height          =   315
      Left            =   5520
      TabIndex        =   17
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Departamento"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaDepartamento 
      Height          =   375
      Left            =   1200
      Top             =   6360
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
      Caption         =   "DtaDepartamento"
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
   Begin MSAdodcLib.Adodc DtaCargo 
      Height          =   375
      Left            =   1320
      Top             =   6960
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
      Caption         =   "DtaCargo"
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
      Height          =   375
      Left            =   1680
      Top             =   7560
      Width           =   3615
      _ExtentX        =   6376
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
   Begin vbskfree.Skinner Skinner1 
      Left            =   480
      Top             =   4200
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.TextBox TxtNumero 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      DownPicture     =   "FrmBuscaEmpleado.frx":0D1A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6000
      Picture         =   "FrmBuscaEmpleado.frx":27FC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton CmdAceptar 
      DownPicture     =   "FrmBuscaEmpleado.frx":42DE
      Height          =   375
      Left            =   6000
      Picture         =   "FrmBuscaEmpleado.frx":5DC0
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox TxtNombre2 
      Height          =   285
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox TxtApellido1 
      Height          =   285
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox TxtApellido2 
      Height          =   285
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox TxtDireccion 
      Height          =   285
      Left            =   1560
      MaxLength       =   200
      TabIndex        =   5
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox TxtNombre1 
      Height          =   285
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Resultados del Criterio"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Número"
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   120
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   120
      X2              =   7440
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label56 
      Caption         =   "Segundo Apellido:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label55 
      Caption         =   "Primer Apellido:"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label54 
      Caption         =   "Segundo Nombre:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Cargo:"
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Departamento"
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Direccion:"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Primer Nombre:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "FrmBuscaEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoTipoNomina As Integer
Private Sub CmdCerrar_Click()
Unload Me
End Sub

Private Sub CmdAceptar_Click()
Dim Criterio As String
On Error GoTo TipoErrs
'Me.DtaEmpleados.Refresh
If Not DtaEmpleados.Recordset.EOF Then
 CodEmpleado = Me.tdbgEmpleados.Columns(1).Text
 Select Case Quien
  Case "Empleado"
    frmEmpleado.DBCodigoEmpleado.Text = CodEmpleado
  Case "Historial"
    FrmHistorial.TDBCombo1.Text = CodEmpleado
    FrmHistorial.TxtCodEmpleado.Text = Me.tdbgEmpleados.Columns(0).Text
  Case "Despido"
    FrmBajas.txtCodEmpleado1.Text = CodEmpleado
  Case "NominaVaca"
    Criterio = "CodEmpleado='" & CodEmpleado & "'"
   Frm13Vaca.DtaVacaciones.Recordset.Find Criterio
  Case "Nomina13vo"
    Criterio = "CodEmpleado='" & CodEmpleado & "'"
   Frm13Vaca.Dta13voMes.Recordset.Find Criterio
  Case "Produccion"
   FrmProduccion.DBCodEmpleado.Text = CodEmpleado
   
  End Select
  Unload Me
Else
 Unload Me
End If
Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub Command2_Click()
Dim SqlEmpleados As String
Dim Numero As String
Dim Nombre1 As String
Dim Nombre2 As String
Dim Apellido1 As String
Dim Apellido2 As String
Dim Direccion As String
Dim Cargo As String
Dim departamento As String

Numero = TxtNumero.Text + "%"
Nombre1 = TxtNombre1.Text + "%"
Nombre2 = TxtNombre2.Text + "%"
Apellido1 = TxtApellido1.Text + "%"
Apellido2 = TxtApellido2.Text + "%"
Direccion = TxtDireccion.Text + "%"
Cargo = DBCCargo.Text + "%"
departamento = DBCDepartamento.Text + "%"

If Quien = "Produccion" Then
    SqlEmpleados = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo,Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion, Empleado.Activo, Empleado.CodTipoNomina,TipoNomina.TipoPago FROM Departamento INNER JOIN Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina WHERE     (Empleado.Activo = 1) AND (TipoNomina.TipoPago = N'Salario Destajo') AND (Empleado.Nombre1 LIKE N'" & Nombre1 & "') AND(Empleado.Nombre2 LIKE N'" & Nombre2 & "') AND (Empleado.Apellido1 LIKE N'" & Apellido1 & "') AND (Empleado.Apellido2 LIKE N'" & Apellido2 & "')AND (Cargo.Cargo LIKE N'" & Cargo & "') AND (Departamento.Departamento LIKE N'" & departamento & "') AND(Empleado.Direccion LIKE N'" & Direccion & "') AND (Empleado.CodEmpleado LIKE N'" & Numero & "') OR" & vbLf
    SqlEmpleados = SqlEmpleados & "(TipoNomina.TipoPago = N'Salario Fijo,Destajo y Comision') OR (TipoNomina.TipoPago = N'Salario Destajo y Comision')"
    DtaEmpleados.RecordSource = SqlEmpleados
    DtaEmpleados.Refresh

Else
    'SqlEmpleados = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion, Empleado.Activo FROM Departamento INNER JOIN (Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) ON Departamento.CodDepartamento = Empleado.CodDepartamento WHERE Empleado.CodEmpleado Like '" & Numero & "' AND Empleado.Nombre1 Like '" & Nombre1 & "' AND Empleado.Nombre2 Like '" & Nombre2 & "' AND Empleado.Apellido1 Like '" & Apellido1 & "' AND Empleado.Apellido2 Like '" & Apellido2 & "' AND Cargo.Cargo Like'" & Cargo & "' AND Departamento.Departamento Like '" & departamento & "' AND Empleado.Direccion Like '" & Direccion & "' And Empleado.Activo=1"
    SqlEmpleados = "SELECT Empleado.CodEmpleado,Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo,Empleado.CodDepartamento , departamento.departamento, Empleado.Direccion, Empleado.Activo FROM Departamento INNER JOIN Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento WHERE(Empleado.Nombre1 LIKE N'" & Nombre1 & "') AND (Empleado.Nombre2 LIKE N'" & Nombre2 & "') AND (Empleado.Apellido1 LIKE N'" & Apellido1 & "') AND (Empleado.Apellido2 LIKE N'" & Apellido2 & "') AND (Cargo.Cargo LIKE N'" & Cargo & "') AND (Departamento.Departamento LIKE N'" & departamento & "') AND (Empleado.Direccion LIKE N'" & Direccion & "') AND (Empleado.Activo = 1) AND (Empleado.CodEmpleado LIKE N'" & Numero & "') ORDER BY Empleado.CodEmpleado1"
    DtaEmpleados.RecordSource = SqlEmpleados
    DtaEmpleados.Refresh
End If
End Sub

Private Sub Form_Activate()
Numero = TxtNumero.Text + "%"
Nombre1 = TxtNombre1.Text + "%"
Nombre2 = TxtNombre2.Text + "%"
Apellido1 = TxtApellido1.Text + "%"
Apellido2 = TxtApellido2.Text + "%"
Direccion = TxtDireccion.Text + "%"
Cargo = DBCCargo.Text + "%"
departamento = DBCDepartamento.Text + "%"

If Quien = "Produccion" Then
    SqlEmpleados = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo,Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion, Empleado.Activo, Empleado.CodTipoNomina,TipoNomina.TipoPago FROM Departamento INNER JOIN Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento INNER JOIN TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina WHERE     (Empleado.Activo = 1) AND (TipoNomina.TipoPago = N'Salario Destajo') AND (Empleado.Nombre1 LIKE N'" & Nombre1 & "') AND(Empleado.Nombre2 LIKE N'" & Nombre2 & "') AND (Empleado.Apellido1 LIKE N'" & Apellido1 & "') AND (Empleado.Apellido2 LIKE N'" & Apellido2 & "')AND (Cargo.Cargo LIKE N'" & Cargo & "') AND (Departamento.Departamento LIKE N'" & departamento & "') AND(Empleado.Direccion LIKE N'" & Direccion & "') AND (Empleado.CodEmpleado LIKE N'" & Numero & "') OR" & vbLf
    SqlEmpleados = SqlEmpleados & "(TipoNomina.TipoPago = N'Salario Fijo,Destajo y Comision') OR (TipoNomina.TipoPago = N'Salario Destajo y Comision')"
    DtaEmpleados.RecordSource = SqlEmpleados
    DtaEmpleados.Refresh

Else
    'SqlEmpleados = "SELECT Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo, Empleado.CodDepartamento, Departamento.Departamento, Empleado.Direccion, Empleado.Activo FROM Departamento INNER JOIN (Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo) ON Departamento.CodDepartamento = Empleado.CodDepartamento WHERE Empleado.CodEmpleado Like '" & Numero & "' AND Empleado.Nombre1 Like '" & Nombre1 & "' AND Empleado.Nombre2 Like '" & Nombre2 & "' AND Empleado.Apellido1 Like '" & Apellido1 & "' AND Empleado.Apellido2 Like '" & Apellido2 & "' AND Cargo.Cargo Like'" & Cargo & "' AND Departamento.Departamento Like '" & departamento & "' AND Empleado.Direccion Like '" & Direccion & "' And Empleado.Activo=1"
    SqlEmpleados = "SELECT Empleado.CodEmpleado,Empleado.CodEmpleado1, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Empleado.CodCargo, Cargo.Cargo,Empleado.CodDepartamento , departamento.departamento, Empleado.Direccion, Empleado.Activo FROM Departamento INNER JOIN Cargo INNER JOIN Empleado ON Cargo.CodCargo = Empleado.CodCargo ON Departamento.CodDepartamento = Empleado.CodDepartamento WHERE(Empleado.Nombre1 LIKE N'" & Nombre1 & "') AND (Empleado.Nombre2 LIKE N'" & Nombre2 & "') AND (Empleado.Apellido1 LIKE N'" & Apellido1 & "') AND (Empleado.Apellido2 LIKE N'" & Apellido2 & "') AND (Cargo.Cargo LIKE N'" & Cargo & "') AND (Departamento.Departamento LIKE N'" & departamento & "') AND (Empleado.Direccion LIKE N'" & Direccion & "') AND (Empleado.Activo = 1) AND (Empleado.CodEmpleado LIKE N'" & Numero & "') ORDER BY Empleado.CodEmpleado1"
    DtaEmpleados.RecordSource = SqlEmpleados
    DtaEmpleados.Refresh
End If

Me.tdbgEmpleados.Columns(0).Visible = False
End Sub

Private Sub Form_Load()
 Me.tdbgEmpleados.EvenRowStyle.BackColor = &HC0FFFF
 Me.tdbgEmpleados.OddRowStyle.BackColor = &HFFFFFF
 Me.tdbgEmpleados.AlternatingRowStyle = True

With Me.DtaCargo
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Cargo"
   .Refresh
End With

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaDepartamento
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "Departamento"
   .Refresh
End With

'With Me.DtaDepartamento
'   '.DatabaseName = Ruta
'   .ConnectionString = Conexion
'End With

With Me.DtaEmpleados
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With


End Sub

Private Sub TxtApellido1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Me.Command2.Value = True
End If
End Sub

Private Sub TxtApellido2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Me.Command2.Value = True
End If
End Sub

Private Sub TxtDireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Me.Command2.Value = True
End If
End Sub

Private Sub TxtNombre1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Me.Command2.Value = True
End If
End Sub

Private Sub TxtNombre2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Me.Command2.Value = True
End If
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Me.Command2.Value = True
End If
End Sub
