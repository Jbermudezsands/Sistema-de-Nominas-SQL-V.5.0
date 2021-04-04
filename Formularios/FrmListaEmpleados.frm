VERSION 5.00
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmListadoEmpleado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Empleados"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   19335
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   7095
      Left            =   17520
      TabIndex        =   1
      Top             =   0
      Width           =   1695
      _Version        =   786432
      _ExtentX        =   2990
      _ExtentY        =   12515
      _StockProps     =   79
      Caption         =   "Controles"
      UseVisualStyle  =   -1  'True
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
         Caption         =   "Editar"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmListaEmpleados.frx":0000
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
         Picture         =   "FrmListaEmpleados.frx":057A
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Nuevo"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmListaEmpleados.frx":0A7E
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton CmdExcel 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Excel"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmListaEmpleados.frx":0FBC
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Agregar Viatico"
         ForeColor       =   0
         Appearance      =   6
         Picture         =   "FrmListaEmpleados.frx":16CE
         ImageAlignment  =   0
      End
   End
   Begin TrueOleDBGrid80.TDBGrid DbgrProducto 
      Bindings        =   "FrmListaEmpleados.frx":3A50
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17295
      _ExtentX        =   30506
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registros:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label LblRegistros 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   7200
      Width           =   1335
   End
End
Attribute VB_Name = "FrmListadoEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnx As New ADODB.Connection
Public rs As New ADODB.Recordset, rsConexion As New ADODB.Recordset
Private sql As String
Private modal As Boolean
Private getVal As Boolean
Private Id As Integer

Private Sub CmdEditar_Click()
  
Quien = "Externos"
frmEmpleado.DBCodigoEmpleado.Text = Me.DbgrProducto.Columns(0).Text
frmEmpleado.CargarDatos
'CargaEmpleados (Me.DbgrProducto.Columns(0).Text)
'frmEmpleado.DBCodigoEmpleado_Change
'frmEmpleado.Show
End Sub

Private Sub CmdExcel_Click(Index As Integer)
    Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
    'Call Formato_Excel(8, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
V = 2
H = 0
i = 1

 '///////////////////////////////////////////////////////////////////////////////////////
 '////////////////////ENCABEZADOS//////////////////////////////////////////////////////
 '///////////////////////////////////////////////////////////////////////////////////

            objExcel.ActiveSheet.Cells(1, 1) = "Codigo"
            objExcel.ActiveSheet.Cells(1, 2) = "Nombres"
            objExcel.ActiveSheet.Cells(1, 3) = "No.Cedula"
            objExcel.ActiveSheet.Cells(1, 4) = "No.Inss"
            objExcel.ActiveSheet.Cells(1, 5) = "Departamento"
            objExcel.ActiveSheet.Cells(1, 6) = "Cargo"
            objExcel.ActiveSheet.Cells(1, 7) = "Sueldo"
            objExcel.ActiveSheet.Cells(1, 8) = "Sexo"
            objExcel.ActiveSheet.Cells(1, 9) = "NumHijos"
            objExcel.ActiveSheet.Cells(1, 10) = "TarifaHoraria"
            objExcel.ActiveSheet.Cells(1, 11) = "Dolarizado"
            objExcel.ActiveSheet.Cells(1, 12) = "FechaNacimiento"
            objExcel.ActiveSheet.Cells(1, 13) = "FechaContrato"

  Do While Not rs.EOF  'esto nos sirve pa leer los datos desde
       
       
            objExcel.ActiveSheet.Cells(V, H + 1) = rs("CodEmpleado1")
            objExcel.ActiveSheet.Cells(V, H + 2) = rs("Nombres")
            objExcel.ActiveSheet.Cells(V, H + 3) = rs("NumCedula")
            objExcel.ActiveSheet.Cells(V, H + 4) = rs("NumeroInss")
            objExcel.ActiveSheet.Cells(V, H + 5) = rs("departamento")
            objExcel.ActiveSheet.Cells(V, H + 6) = rs("Cargo")
            objExcel.ActiveSheet.Cells(V, H + 7) = rs("SueldoPeriodo")
            objExcel.ActiveSheet.Cells(V, H + 8) = rs("Sexo")
            objExcel.ActiveSheet.Cells(V, H + 9) = rs("NumHijos")
            objExcel.ActiveSheet.Cells(V, H + 10) = rs("TarifaHoraria")
            objExcel.ActiveSheet.Cells(V, H + 11) = rs("Dolarizado")
            objExcel.ActiveSheet.Cells(V, H + 12) = rs("FechaNacimiento")
            objExcel.ActiveSheet.Cells(V, H + 13) = rs("FechaContrato")

            
            V = V + 1
            i = i + 1
            rs.MoveNext

 

  Loop
  
        objExcel.ActiveSheet.Columns("A").ColumnWidth = 10
        objExcel.ActiveSheet.Columns("B").ColumnWidth = 40
        objExcel.ActiveSheet.Columns("C").ColumnWidth = 18
        objExcel.ActiveSheet.Columns("D").ColumnWidth = 10
        objExcel.ActiveSheet.Columns("E").ColumnWidth = 30
        objExcel.ActiveSheet.Columns("F").ColumnWidth = 30
        objExcel.ActiveSheet.Columns("G").ColumnWidth = 17
        objExcel.ActiveSheet.Columns("G").NumberFormat = "##,##0.00"
        objExcel.ActiveSheet.Columns("H").ColumnWidth = 10
        objExcel.ActiveSheet.Columns("I").ColumnWidth = 10
        objExcel.ActiveSheet.Columns("J").ColumnWidth = 10
        objExcel.ActiveSheet.Columns("J").NumberFormat = "##,##0.00"
        objExcel.ActiveSheet.Columns("K").ColumnWidth = 10
        objExcel.ActiveSheet.Columns("L").ColumnWidth = 17
        objExcel.ActiveSheet.Columns("L").NumberFormat = "dd/mm/yyyy"
        objExcel.ActiveSheet.Columns("M").ColumnWidth = 15
        objExcel.ActiveSheet.Columns("M").NumberFormat = "dd/mm/yyyy"
        
        
        
'        objExcel.ActiveSheet.Columns("A").Font.Size = 10
'        objExcel.ActiveSheet.Columns("B").NumberFormat = "############"
'        objExcel.ActiveSheet.Columns("B").ColumnWidth = 17
'        objExcel.ActiveSheet.Columns("B").Font.Size = 10
'        objExcel.ActiveSheet.Columns("B").HorizontalAlignment = xlHAlignCenter
'        objExcel.ActiveSheet.Columns("C").ColumnWidth = 26
'        objExcel.ActiveSheet.Columns("C").Font.Size = 10
'        objExcel.ActiveSheet.Columns("C").HorizontalAlignment = xlHAlignCenter
'        objExcel.ActiveSheet.Columns("D").ColumnWidth = 10
'        objExcel.ActiveSheet.Columns("D").Font.Size = 10
'        objExcel.ActiveSheet.Columns("D").HorizontalAlignment = xlHAlignCenter
'        objExcel.ActiveSheet.Columns("E").ColumnWidth = 4
'        objExcel.ActiveSheet.Columns("E").Font.Size = 10
'        objExcel.ActiveSheet.Columns("E").HorizontalAlignment = xlHAlignCenter

 
 Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DbgrProducto_DblClick()
Quien = "Externos"
frmEmpleado.DBCodigoEmpleado.Text = Me.DbgrProducto.Columns(0).Text
frmEmpleado.CargarDatos
End Sub

Private Sub DbgrProducto_FilterChange()
On Error GoTo errTdbg
    'Gets called when an action is performed on the filter bar
    Dim col As TrueOleDBGrid80.Column
    Dim cols As TrueOleDBGrid80.Columns
    
    'On Error GoTo errHandler
    On Error Resume Next
    Set cols = Me.DbgrProducto.Columns
    Dim c As Integer
    
    c = DbgrProducto.col
    DbgrProducto.HoldFields
    sql = rs.Filter
    rs.Filter = getFilter(col, cols)
    
    DbgrProducto.col = c
    DbgrProducto.EditActive = True

    Me.LblRegistros.Caption = rs.RecordCount

Exit Sub
errTdbg:
    MsgBox Err.Description
End Sub
Private Function getFilter(col As TrueOleDBGrid80.Column, cols As TrueOleDBGrid80.Columns) As String
'Creates the SQL statement in adodc1.recordset.filter
'and only filters text currently. It must be modified to
'filter other data types.
Dim tmp As String
Dim n As Integer
Dim X As Integer

For Each col In cols
    If Trim(col.FilterText) <> "" Then
        n = n + 1
        If n > 1 Then tmp = tmp & " AND "
        Select Case rs.Fields(X).Type
        Case adVarWChar, adVarChar: tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
        Case adInteger, adNumeric: tmp = tmp & "[" & col.DataField & "] = " & col.FilterText
        Case adDBTimeStamp: tmp = tmp & "[" & col.DataField & "] = #" & col.FilterText & "#"
        Case adDouble, adNumeric: tmp = tmp & "[" & col.DataField & "] = " & col.FilterText
        Case Else: tmp = tmp & "[" & col.DataField & "] = " & col.FilterText
        End Select
    End If
    X = X + 1
Next col
getFilter = tmp

End Function

Private Sub DbgrProducto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Quien = "Externos"
frmEmpleado.DBCodigoEmpleado.Text = Me.DbgrProducto.Columns(0).Text
frmEmpleado.CargarDatos

 
End If
End Sub

Private Sub Form_Load()
 Dim SQlConsulta As String
Me.BackColor = RGB(222, 227, 247)
Me.GroupBox1.BackColor = RGB(222, 227, 247)
 Me.DbgrProducto.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DbgrProducto.OddRowStyle.BackColor = &H80000005
 Me.DbgrProducto.AlternatingRowStyle = True
 
'        SqlConsulta = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Empleado.CodEmpleado, Empleado.NumCedula , Empleado.NumeroInss, departamento.departamento, Cargo.Cargo, Empleado.Sexo FROM Empleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN Cargo ON Empleado.CodCargo = Cargo.CodCargo  Where (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1"
        SQlConsulta = "SELECT  Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Empleado.CodEmpleado, Empleado.NumCedula, Empleado.NumeroInss, Departamento.Departamento, Cargo.Cargo, Empleado.SueldoPeriodo, Empleado.Sexo, Empleado.NumHijos, Empleado.TarifaHoraria , Empleado.Dolarizado, Historico.FechaNacimiento, Historico.FechaContrato, Empleado.CuentaBanco, Empleado.NumeroInss FROM Empleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN Cargo ON Empleado.CodCargo = Cargo.CodCargo INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado Where (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1"
       
        With rs
          .CursorLocation = adUseClient
          .Open SQlConsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
             
    
        Me.DbgrProducto.DataSource = rs
        Me.DbgrProducto.Columns(0).Caption = "Codigo"
        Me.DbgrProducto.Columns(0).Width = 1000
        Me.DbgrProducto.Columns(1).Width = 4000
        Me.DbgrProducto.Columns(2).Visible = False
        Me.DbgrProducto.Columns(5).Width = 3000
        Me.DbgrProducto.Columns(6).Width = 3000
        

        Me.LblRegistros.Caption = rs.RecordCount
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set cnx = Nothing
    Set rs = Nothing
    Set rs2 = Nothing
    Set rs3 = Nothing
End Sub


Private Sub PushButton1_Click()
                  MDIPrimero.MousePointer = 11
                   frmEmpleado.Show
                  MDIPrimero.MousePointer = 0
End Sub

Private Sub PushButton3_Click()

End Sub

Private Sub PushButton2_Click(Index As Integer)
  FrmIncentivosGrupo.Show 1

End Sub
