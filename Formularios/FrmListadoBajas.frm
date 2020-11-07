VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmListaBajas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Bajas"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   15285
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc AdoBajas 
      Height          =   375
      Left            =   360
      Top             =   7680
      Width           =   4095
      _ExtentX        =   7223
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
      Caption         =   "Adodc1"
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
   Begin TrueOleDBGrid80.TDBGrid DbgrProducto 
      Bindings        =   "FrmListadoBajas.frx":0000
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   12515
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
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
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
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=8,.bold=0,.fontsize=825,.italic=0"
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
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   7095
      Left            =   13560
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      _Version        =   786432
      _ExtentX        =   2990
      _ExtentY        =   12515
      _StockProps     =   79
      Caption         =   "Controles"
      UseVisualStyle  =   -1  'True
      Begin MSComCtl2.DTPicker DTFechaFin 
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   3240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   74842113
         CurrentDate     =   43746
      End
      Begin MSComCtl2.DTPicker DTFechaIni 
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   74842113
         CurrentDate     =   43746
      End
      Begin XtremeSuiteControls.PushButton CmdEditar 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Procesar"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmListadoBajas.frx":0017
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdSalir 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   6600
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Salir"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmListadoBajas.frx":21C5
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   4920
         Visible         =   0   'False
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Filtrar"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmListadoBajas.frx":26C9
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Eliminar"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmListadoBajas.frx":4921
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmListadoBajas.frx":6B3C
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Abrir"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmListadoBajas.frx":707A
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Filtrar"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmListadoBajas.frx":778C
         ImageAlignment  =   0
      End
   End
End
Attribute VB_Name = "FrmListaBajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Filtrado As Boolean
Private Sub CmdEditar_Click()
 Dim Respuesta As Double

 Respuesta = MsgBox("Esta seguro de procesar los Registros?", vbYesNo, "Zeus Nominas")
 
 If Respuesta <> 6 Then
  Exit Sub
 End If

 Me.AdoBajas.Refresh
 Do While Not Me.AdoBajas.Recordset.EOF
  Me.AdoBajas.Recordset("Procesada") = 1
  Me.AdoBajas.Recordset.Update
  
  Me.AdoBajas.Recordset.MoveNext
 Loop
 
 
 Me.AdoBajas.Refresh
 
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Activate()
 Me.AdoBajas.ConnectionString = Conexion
 Me.AdoBajas.RecordSource = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Bajas.Id, Bajas.TipoBaja, Bajas.FechaBaja, Bajas.AnnosTrabajados, Bajas.MesesTrabajados, Bajas.DiasTrabajados, Bajas.SalarioMensual, Bajas.Prestamo, Bajas.Deducciones, Bajas.Calculada, Bajas.Procesada, Empleado.CuentaBanco, Historico.FechaContrato, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodEmpleado FROM Empleado INNER JOIN Bajas ON Empleado.CodEmpleado = Bajas.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado  " & _
                            "Where (Bajas.Calculada = 1) And (Bajas.Procesada = 0) And (Empleado.Activo = 1)"
 Me.AdoBajas.Refresh
 Filtrado = False
 

End Sub

Private Sub Form_Load()
 Me.DbgrProducto.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DbgrProducto.OddRowStyle.BackColor = &H80000005
 Me.DbgrProducto.AlternatingRowStyle = True
 Filtrado = False
End Sub

Private Sub PushButton1_Click()
Dim CodEmpleado As Double, CodEmpleado1 As String
 
Quien = "ListaBaja"
CodEmpleado = Me.DbgrProducto.Columns(13).Text
CodEmpleado1 = Me.DbgrProducto.Columns(0).Text

FrmBajas.txtCodEmpleado1.Text = CodEmpleado1
FrmBajas.Show


End Sub

Private Sub PushButton2_Click()
 Me.AdoBajas.ConnectionString = Conexion
 Me.AdoBajas.RecordSource = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Bajas.Id, Bajas.TipoBaja, Bajas.FechaBaja, Bajas.AnnosTrabajados, Bajas.MesesTrabajados, Bajas.DiasTrabajados, Bajas.SalarioMensual, Bajas.Prestamo, Bajas.Deducciones, Bajas.Calculada, Bajas.Procesada, Empleado.CuentaBanco, Historico.FechaContrato, Empleado.NumeroInss, Empleado.NumeroRuc, Empleado.CodEmpleado FROM Empleado INNER JOIN Bajas ON Empleado.CodEmpleado = Bajas.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado  " & _
                            "WHERE (Bajas.Calculada = 1) AND (Bajas.Procesada = 0) AND (Empleado.Activo = 1) AND (Bajas.FechaBaja BETWEEN CONVERT(DATETIME, '" & Format(Me.DTFechaIni.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTFechaFin.Value, "yyyy-mm-dd") & "', 102))"
 Me.AdoBajas.Refresh
 Filtrado = True
 
End Sub

Private Sub PushButton3_Click()
 On Error GoTo TipoErrs
 Dim Respuesta, Rsp
'Elimino el registro activo en la pantalla

  Respuesta = MsgBox("Esta seguro de Borrar el registro?", vbYesNo, "Borrando el Baja: " & Me.DbgrProducto.Columns(0).Text)
   If Respuesta = 6 Then
     
     Me.AdoBajas.Recordset.Delete

   End If
 Exit Sub
TipoErrs:
 ControlErrores
 Unload Me
End Sub

Private Sub PushButton4_Click()

ArepListaBajas.DataControl1.ConnectionString = ConexionReporte
If Filtrado = False Then
  ArepListaBajas.DataControl1.Source = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Bajas.*, Historico.FechaContrato, Departamento.Departamento, Empleado.NumeroInss FROM Empleado INNER JOIN Bajas ON Empleado.CodEmpleado = Bajas.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento Where (Bajas.Calculada = 1) And (Bajas.Procesada = 0) And (Empleado.Activo = 1) ORDER BY Nombres"
Else
ArepListaBajas.DataControl1.Source = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Bajas.*, Historico.FechaContrato, Departamento.Departamento, Empleado.NumeroInss FROM Empleado INNER JOIN Bajas ON Empleado.CodEmpleado = Bajas.CodEmpleado INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento Where (Bajas.Calculada = 1) And (Bajas.Procesada = 0) And (Empleado.Activo = 1) AND (Bajas.FechaBaja BETWEEN CONVERT(DATETIME, '" & Format(Me.DTFechaIni.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(Me.DTFechaFin.Value, "yyyy-mm-dd") & "', 102)) ORDER BY Nombres"

End If

 ArepListaBajas.Show 1
  
End Sub

Private Sub PushButton5_Click()
Dim CodEmpleado As Double, CodEmpleado1 As String
 
Quien = "ListaBaja"
CodEmpleado = Me.DbgrProducto.Columns(13).Text
CodEmpleado1 = Me.DbgrProducto.Columns(0).Text

FrmBajas.txtCodEmpleado1.Text = CodEmpleado1
FrmBajas.Show
End Sub
