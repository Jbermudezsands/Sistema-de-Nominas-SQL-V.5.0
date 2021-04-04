VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FrmConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Registros"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   HelpContextID   =   6
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc DtaProductos 
      Height          =   375
      Left            =   960
      Top             =   4920
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
      Caption         =   "DtaProductos"
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
      Left            =   1080
      Top             =   6120
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
   Begin vbskfree.Skinner Skinner1 
      Left            =   4080
      Top             =   1800
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin VB.CommandButton CmdPegar 
      Caption         =   "&Pegar"
      Height          =   375
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2280
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton CmdOrden 
      Caption         =   "&Orden"
      Height          =   375
      Left            =   1200
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin TrueOleDBGrid80.TDBGrid DbgrProducto 
      Bindings        =   "FrmConsulCompra.frx":0000
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   4895
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
End
Attribute VB_Name = "FrmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CuentaContable As String, CodigoEmpleado As Double, CodigoEmpleado1 As String, NumeroSolicitud As String, FechaSolicitud As Date, TipoSolicitud As String
Private cnx As New ADODB.Connection
Private rs As New ADODB.Recordset, rsConexion As New ADODB.Recordset
Private sql As String
Private modal As Boolean
Private getVal As Boolean
Private Id As Integer
Private Sub DbgrProducto_FilterChange()
    DbgrProducto.PostMsg (418)
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
        End Select
    End If
    X = X + 1
Next col

getFilter = tmp

End Function

Private Sub DbgrProducto_PostEvent(ByVal MsgId As Integer)
If MsgId = 418 Then
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
    Exit Sub
errTdbg:
        MsgBox Err.Description
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set cnx = Nothing
Set rs = Nothing
Set rs2 = Nothing
Set rs3 = Nothing
End Sub


Private Sub CmdCancelar_Click()
On Error GoTo TipoErrs
Unload Me
Exit Sub
TipoErrs:
   ControlErrores
End Sub

Private Sub CmdOrden_Click()
On Error GoTo TipoErrs
Select Case QueProducto
      Case "Produccion"
         If Orden = True Then
         SQlConsulta = "SELECT CodProceso,Descrip,CodReferencia, Ref,Precio, Unid From Procesos ORDER BY CodProceso, CodReferencia"
         DtaProductos.RecordSource = SQlConsulta
         DtaProductos.Refresh
         Respuesta = ""
         'Me.DbgrProducto.Columns(1).Caption = "Codidgo Ref"
         Me.DbgrProducto.Columns(4).Width = 1000
         Me.DbgrProducto.Columns(4).NumberFormat = "##,##0.00"
         Me.DbgrProducto.Columns(5).Width = 1000
         DbgrProducto.Columns(1).Width = 4200
            Orden = False
        Else
         SQlConsulta = "SELECT Descrip,CodProceso,CodReferencia, Ref, Precio, Unid From Procesos ORDER BY Descrip,CodProceso, CodReferencia"
         DtaProductos.RecordSource = SQlConsulta
         DtaProductos.Refresh
         Respuesta = ""
         'Me.DbgrProducto.Columns(1).Caption = "Codidgo Ref"
         Me.DbgrProducto.Columns(4).Width = 1000
         Me.DbgrProducto.Columns(4).NumberFormat = "##,##0.00"
         Me.DbgrProducto.Columns(5).Width = 1000
         DbgrProducto.Columns(0).Width = 4200
            Orden = True
         End If

      Case "CuentaContable"
      
         If Orden = True Then
         SQlConsulta = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta From Cuentas ORDER BY CodCuentas"
         DtaProductos.RecordSource = SQlConsulta
         DtaProductos.Refresh
         Respuesta = ""
            Me.DbgrProducto.Columns(0).Width = 1500
            Me.DbgrProducto.Columns(1).Width = 6000
            Me.DbgrProducto.Columns(2).Width = 1500
            Orden = False
        Else
         SQlConsulta = "SELECT DescripcionCuentas, CodCuentas, TipoCuenta From Cuentas ORDER BY DescripcionCuentas"
         DtaProductos.RecordSource = SQlConsulta
         DtaProductos.Refresh
         Respuesta = ""
            Me.DbgrProducto.Columns(0).Width = 6000
            Me.DbgrProducto.Columns(1).Width = 1500
            Me.DbgrProducto.Columns(2).Width = 1500
            Orden = True
         End If
         
    End Select
 Exit Sub
TipoErrs:
ControlErrores
End Sub

Private Sub CmdPegar_Click()
'On Error GoTo TipoErrs
Dim valor1 As String, Valor As String, PrecioCosto As Integer
Dim cantidad, Subtotal0, Subtotal1, Producto, Valores, Costo As Double
Dim CodigoP As String, Candena As String, Fecha As Long
Dim PrecioC As Double, CantidadC As Double, SubT As Double
Dim AnteriorSub As Double, Cadena As String
Dim SqlDetalle As String, Numero As String, TipoCuenta As String
  
 'Busco el numero consecutivo de la Recepcion
  Select Case QueProducto
    Case "CodigoProductoHistorico"
         CodigoEmpleado1 = Me.DbgrProducto.Columns(0).Text
         CodigoEmpleado = Me.DbgrProducto.Columns(2).Text
     
    Case "Justifica"
        CodigoEmpleado1 = Me.DbgrProducto.Columns("Classid").Text
    Case "Solicitud"
        CodigoEmpleado1 = Me.DbgrProducto.Columns("CodEmpleado").Text
        FechaSolicitud = Me.DbgrProducto.Columns("FechaSolicitud").Text
        TipoSolicitud = Me.DbgrProducto.Columns("TipoSolicitud").Text
        NumeroSolicitud = Me.DbgrProducto.Columns("NumeroSolicitud").Text
        
        'CodigoEmpleado1 = Me.DtaProductos.Recordset("CodEmpleado1")
       ' FechaSolicitud = Me.DtaProductos.Recordset("FechaSolicitud")
        'TipoSolicitud = Me.DtaProductos.Recordset("TipoSolicitud")
       ' NumeroSolicitud = Me.DtaProductos.Recordset("NumeroSolicitud")
    Case "CuentaContable"
         CuentaContable = Me.DtaProductos.Recordset("CodCuentas")
         
    Case "EmpleadosSoli"
         FrmSolicitud.DBCodigoEmpleado.Text = Me.DbgrProducto.Columns("CodEmpleado1").Text
         
    Case "Produccion"
         FrmProduccion.TDBGProduccion.Columns(0).Text = Me.DtaProductos.Recordset("CodProceso")
         FrmProduccion.TDBGProduccion.Columns(1).Text = Me.DtaProductos.Recordset("CodReferencia")
         FrmProduccion.TDBGProduccion.Columns(2).Text = Me.DtaProductos.Recordset("Ref")
         FrmProduccion.TDBGProduccion.Columns(3).Text = Me.DtaProductos.Recordset("Precio")
         FrmProduccion.TDBGProduccion.Columns(4).Text = Me.DtaProductos.Recordset("Unid")
    Case "CodigoEmpleado"
         CodigoEmpleado1 = Me.DbgrProducto.Columns(0).Text
         CodigoEmpleado = Me.DbgrProducto.Columns(2).Text
         
    Case "CodigoEmpleadoReportesIni"
         CodigoEmpleado1 = Me.DbgrProducto.Columns(0).Text
         
    Case "CodigoEmpleadoReportesFin"
         CodigoEmpleado1 = Me.DbgrProducto.Columns(0).Text
  End Select


 Unload Me
Exit Sub
TipoErrs:
MsgBox Err.Description
End Sub

Private Sub CmdX_Click()
On Error GoTo TipoErrs
Unload Me
Exit Sub
TipoErrs:
MsgBox Err.Description
End Sub

Private Sub DbgrProducto_DblClick()
On Error GoTo TipoErrs
CmdPegar.Value = True
Exit Sub
TipoErrs:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo TipoErrs

' MDIPrimero.Skin1.ApplySkin hWnd
 Me.DbgrProducto.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DbgrProducto.OddRowStyle.BackColor = &H80000005
 Me.DbgrProducto.AlternatingRowStyle = True

If QueProducto = "CuentaContable" Then
    With Me.DtaProductos
       .ConnectionString = ConexionContable
    End With
    
    With Me.DtaConsulta
       .ConnectionString = ConexionContable
    End With
    

    If cnx.State = adStateClosed Then
        cnx.ConnectionString = ConexionContable
        cnx.Open
    End If

Else
    With Me.DtaProductos
       .ConnectionString = Conexion
    End With
    
    With Me.DtaConsulta
       .ConnectionString = Conexion
    End With
    
    If cnx.State = adStateClosed Then
        cnx.ConnectionString = Conexion
        cnx.Open
    End If
End If

Origen = ""
CuentaContable = ""

Dim SQlConsulta As String
 'CmdX.Visible = True
Orden = True
Select Case QueProducto

      Case "CodigoProductoHistorico"
      
            SQlConsulta = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Empleado.CodEmpleado, Empleado.NumCedula , Empleado.NumeroInss, departamento.departamento, Cargo.Cargo, Empleado.Sexo FROM Empleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN Cargo ON Empleado.CodCargo = Cargo.CodCargo  ORDER BY Empleado.CodEmpleado1"
            DtaProductos.RecordSource = SQlConsulta
            DtaProductos.Refresh
            
            With rs
              .CursorLocation = adUseClient
              .Open SQlConsulta, Conexion, adOpenDynamic, adLockOptimistic
            End With
                 
        
            Me.DbgrProducto.DataSource = rs
            
            Me.DbgrProducto.Columns(0).Width = 1500
            Me.DbgrProducto.Columns(1).Width = 6000
            Me.DbgrProducto.Columns(2).Visible = False
            
            Case "CodigoEmpleadoReportesIni"
            SQlConsulta = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Empleado.CodEmpleado, Empleado.NumCedula , Empleado.NumeroInss, departamento.departamento, Cargo.Cargo, Empleado.Sexo FROM Empleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN Cargo ON Empleado.CodCargo = Cargo.CodCargo  Where (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1"
            DtaProductos.RecordSource = SQlConsulta
            DtaProductos.Refresh
            
            With rs
              .CursorLocation = adUseClient
              .Open SQlConsulta, Conexion, adOpenDynamic, adLockOptimistic
            End With
                 
        
            Me.DbgrProducto.DataSource = rs
            
            Me.DbgrProducto.Columns(0).Width = 1500
            Me.DbgrProducto.Columns(1).Width = 6000
            Me.DbgrProducto.Columns(2).Visible = False






       Case "Justifica"
       
       'SQlConsulta = "SELECT     SolicitudVacaciones.FechaSolicitud AS FechaSolicitud, SolicitudVacaciones.TipoSolicitud, SolicitudVacaciones.CodigoEmpleado as CodEmpleado,   Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, MAX(SolicitudVacaciones.NumeroSolicitud) AS NumeroSolicitud  FROM         SolicitudVacaciones INNER JOIN  Empleado ON SolicitudVacaciones.CodigoEmpleado = Empleado.CodEmpleado1  GROUP BY SolicitudVacaciones.FechaSolicitud, SolicitudVacaciones.TipoSolicitud, SolicitudVacaciones.CodigoEmpleado,  Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2  ORDER BY NumeroSolicitud"
       
       SQlConsulta = "SELECT Classid, Classname From LeaveClass ORDER BY Classid"
       DtaProductos.RecordSource = SQlConsulta
       DtaProductos.Refresh
       
       With rs
          .CursorLocation = adUseClient
          .Open SQlConsulta, Conexion, adOpenDynamic, adLockOptimistic
       End With
             
    
       Me.DbgrProducto.DataSource = rs
       
       Me.DbgrProducto.Columns(0).Width = 1500
       Me.DbgrProducto.Columns(1).Width = 1500

       


       Case "Solicitud"
       
       'SQlConsulta = "SELECT     SolicitudVacaciones.FechaSolicitud AS FechaSolicitud, SolicitudVacaciones.TipoSolicitud, SolicitudVacaciones.CodigoEmpleado as CodEmpleado,   Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, MAX(SolicitudVacaciones.NumeroSolicitud) AS NumeroSolicitud  FROM         SolicitudVacaciones INNER JOIN  Empleado ON SolicitudVacaciones.CodigoEmpleado = Empleado.CodEmpleado1  GROUP BY SolicitudVacaciones.FechaSolicitud, SolicitudVacaciones.TipoSolicitud, SolicitudVacaciones.CodigoEmpleado,  Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2  ORDER BY NumeroSolicitud"
       
       SQlConsulta = "SELECT     SolicitudVacaciones.FechaSolicitud, SolicitudVacaciones.TipoSolicitud, Empleado.CodEmpleado1 as CodEmpleado,    Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, SolicitudVacaciones.NumeroSolicitud  FROM         SolicitudVacaciones INNER JOIN    Empleado ON SolicitudVacaciones.CodigoEmpleado = Empleado.CodEmpleado1 ORDER BY SolicitudVacaciones.NumeroSolicitud DESC"
       DtaProductos.RecordSource = SQlConsulta
       DtaProductos.Refresh
       
       With rs
          .CursorLocation = adUseClient
          .Open SQlConsulta, Conexion, adOpenDynamic, adLockOptimistic
       End With
             
    
       Me.DbgrProducto.DataSource = rs
       
       Me.DbgrProducto.Columns(0).Width = 1500
       Me.DbgrProducto.Columns(1).Width = 1500
       Me.DbgrProducto.Columns(2).Width = 1500
       
       Case "EmpleadosSoli"
       
        SQlConsulta = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Empleado.CodEmpleado, Empleado.NumCedula , Empleado.NumeroInss, departamento.departamento, Cargo.Cargo, Empleado.Sexo FROM Empleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN Cargo ON Empleado.CodCargo = Cargo.CodCargo  Where (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1"
        DtaProductos.RecordSource = SQlConsulta
        DtaProductos.Refresh
        
        With rs
          .CursorLocation = adUseClient
          .Open SQlConsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
             
    
        Me.DbgrProducto.DataSource = rs
        
        Me.DbgrProducto.Columns(0).Width = 1500
        Me.DbgrProducto.Columns(1).Width = 6000
        Me.DbgrProducto.Columns(2).Visible = False
       

       Case "CuentaContable"
       SQlConsulta = "SELECT DescripcionCuentas, CodCuentas,TipoCuenta From Cuentas ORDER BY CodCuentas"
       DtaProductos.RecordSource = SQlConsulta
       DtaProductos.Refresh
       
       With rs
          .CursorLocation = adUseClient
          .Open SQlConsulta, ConexionContable, adOpenDynamic, adLockOptimistic
       End With
             
    
       Me.DbgrProducto.DataSource = rs
       
       Me.DbgrProducto.Columns(0).Width = 6000
       Me.DbgrProducto.Columns(1).Width = 1500
       Me.DbgrProducto.Columns(2).Width = 1500
       
       Case "Produccion"
         SQlConsulta = "SELECT Descrip,CodProceso,CodReferencia, Ref,Precio, Unid From Procesos ORDER BY Descrip,CodProceso, CodReferencia"
         DtaProductos.RecordSource = SQlConsulta
         DtaProductos.Refresh
         
         With rs
          .CursorLocation = adUseClient
          .Open SQlConsulta, Conexion, adOpenDynamic, adLockOptimistic
         End With
             
    
         Me.DbgrProducto.DataSource = rs
         
         Respuesta = ""
         'Me.DbgrProducto.Columns(1).Caption = "Codidgo Ref"
         Me.DbgrProducto.Columns(4).Width = 1000
         Me.DbgrProducto.Columns(4).NumberFormat = "##,##0.00"
         Me.DbgrProducto.Columns(5).Width = 1000
         DbgrProducto.Columns(0).Width = 4200
         
        
        Case "CodigoEmpleado"
        SQlConsulta = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Empleado.CodEmpleado, Empleado.NumCedula , Empleado.NumeroInss, departamento.departamento, Cargo.Cargo, Empleado.Sexo FROM Empleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN Cargo ON Empleado.CodCargo = Cargo.CodCargo  Where (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1"
        DtaProductos.RecordSource = SQlConsulta
        DtaProductos.Refresh
        
        With rs
          .CursorLocation = adUseClient
          .Open SQlConsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
             
    
        Me.DbgrProducto.DataSource = rs
        
        Me.DbgrProducto.Columns(0).Width = 1500
        Me.DbgrProducto.Columns(1).Width = 6000
        Me.DbgrProducto.Columns(2).Visible = False
        
        Case "CodigoEmpleadoReportesIni"
        SQlConsulta = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Empleado.CodEmpleado, Empleado.NumCedula , Empleado.NumeroInss, departamento.departamento, Cargo.Cargo, Empleado.Sexo FROM Empleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN Cargo ON Empleado.CodCargo = Cargo.CodCargo  Where (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1"
        DtaProductos.RecordSource = SQlConsulta
        DtaProductos.Refresh
        
        With rs
          .CursorLocation = adUseClient
          .Open SQlConsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
             
    
        Me.DbgrProducto.DataSource = rs
        
        Me.DbgrProducto.Columns(0).Width = 1500
        Me.DbgrProducto.Columns(1).Width = 6000
        Me.DbgrProducto.Columns(2).Visible = False
        
         Case "CodigoEmpleadoReportesFin"
        SQlConsulta = "SELECT Empleado.CodEmpleado1, Empleado.Nombre1 + ' ' + Empleado.Nombre2 + ' ' + Empleado.Apellido1 + ' ' + Empleado.Apellido2 AS Nombres, Empleado.CodEmpleado, Empleado.NumCedula , Empleado.NumeroInss, departamento.departamento, Cargo.Cargo, Empleado.Sexo FROM Empleado INNER JOIN Departamento ON Empleado.CodDepartamento = Departamento.CodDepartamento INNER JOIN Cargo ON Empleado.CodCargo = Cargo.CodCargo  Where (Empleado.Activo = 1) ORDER BY Empleado.CodEmpleado1"
        DtaProductos.RecordSource = SQlConsulta
        DtaProductos.Refresh
        
        With rs
          .CursorLocation = adUseClient
          .Open SQlConsulta, Conexion, adOpenDynamic, adLockOptimistic
        End With
             
    
        Me.DbgrProducto.DataSource = rs
        
        Me.DbgrProducto.Columns(0).Width = 1500
        Me.DbgrProducto.Columns(1).Width = 6000
        Me.DbgrProducto.Columns(2).Visible = False
 
       End Select
   Me.DbgrProducto.MarqueeStyle = dbgHighlightCell

Exit Sub
TipoErrs:
MsgBox Err.Description
Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo TipoErrs
Campo = False
Exit Sub
TipoErrs:
MsgBox Err.Description
End Sub




'---------------------------CODIGO RETIRADO -------------------
'Private Sub DbgrProducto_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo TipoErrs
' Dim consulta
' Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
' If Shift = 1 Then
'  If Not KeyCode = 16 Then
'   Select Case KeyCode
'    Case 57
'     Lectura = "("
'    Case 219
'     Lectura = "-"
'    Case 48
'     Lectura = ")"
'   End Select
'  Else
'   Exit Sub
'  End If
' Else
'  'Si no Preciona Shift leo la tecla
'    Lectura = KeyCode
' '////////Chequeo si utiliza las Flechas direccioneales//////////
'   If Lectura = 40 Then
'    Exit Sub
'   End If
'   If Lectura = 38 Then
'    Exit Sub
'   End If
'   If Lectura = 39 Then
'    Exit Sub
'   End If
'   If Lectura = 37 Then
'    Exit Sub
'   End If
' '///////////Fin de la Busqueda Direccional////////////////////////
'
'  If Lectura = 8 Then
'   If Not Respuesta = "" Then
'    Respuesta = Mid$(Respuesta, 1, Len(Respuesta) - 1)
'    Lectura = ""
'    If Not Origen = "" Then
'     Origen = Mid$(Origen, 1, Len(Origen) - 1)
'    End If
'   End If
'  Else
'
'   LeeTecla
'  End If
' End If
'
'
'
' If KeyCode = 13 Then
'  Me.CmdPegar.Value = True
' Else
'   Select Case QueProducto
'        Case "Produccion"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT CodProceso,Descrip,CodReferencia, Ref, Precio, Unid From Procesos WHERE     (CodProceso LIKE '" & Respuesta & "%') ORDER BY CodProceso, CodReferencia"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT CodProceso,Descrip, CodReferencia, Ref, Precio, Unid From Procesos WHERE     (CodProceso LIKE '" & Respuesta & "%') ORDER BY CodProceso, CodReferencia"
'              DtaProductos.Refresh
'           End If
'            Me.DbgrProducto.Columns(4).Width = 1000
'            Me.DbgrProducto.Columns(4).NumberFormat = "##,##0.00"
'            Me.DbgrProducto.Columns(5).Width = 1000
'            Me.DbgrProducto.Columns(0).Width = 4200
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'         Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT Descrip,CodProceso,CodReferencia, Ref, Precio, Unid From Procesos WHERE     (Descrip LIKE '" & Respuesta & "%') ORDER BY Descrip,CodProceso, CodReferencia"
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'                Respuesta = "" & Respuesta & Lectura & ""
'                DtaProductos.RecordSource = "SELECT Descrip,CodProceso,CodReferencia, Ref, Precio, Unid From Procesos WHERE     (Descrip LIKE '" & Respuesta & "%') ORDER BY Descrip,CodProceso, CodReferencia"
'                DtaProductos.Refresh
'              End If
'                Me.DbgrProducto.Columns(4).Width = 1000
'                Me.DbgrProducto.Columns(4).NumberFormat = "##,##0.00"
'                Me.DbgrProducto.Columns(5).Width = 1000
'                Me.DbgrProducto.Columns(0).Width = 4200
'                Orden = True
'
'          End If
'
'
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'
'
'         Case "CuentaContable"
'         If Orden = False Then
'           If Respuesta = "" Or Lectura = "%" Then
'              Respuesta = Lectura
'
'              DtaProductos.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta From Cuentas WHERE (CodCuentas LIKE '" & Respuesta & "%') ORDER BY CodCuentas"
'              DtaProductos.Refresh
'
'              If Lectura = "%" Then
'                 Respuesta = ""
'                 Lectura = ""
'                 Origen = ""
'              End If
'           Else
'              Respuesta = "" & Respuesta & Lectura & ""
'              DtaProductos.RecordSource = "SELECT CodCuentas, DescripcionCuentas, TipoCuenta From Cuentas WHERE (CodCuentas LIKE '" & Respuesta & "%') ORDER BY CodCuentas"
'              DtaProductos.Refresh
'           End If
'            Me.DbgrProducto.Columns(0).Width = 1500
'            Me.DbgrProducto.Columns(1).Width = 6000
'            Me.DbgrProducto.Columns(2).Width = 1500
'              Orden = False
'            'Si el orden  Cambia a codigo entra a esta opcion
'         Else
'              If Respuesta = "" Or Lectura = "%" Then
'                 Respuesta = Lectura
'                 DtaProductos.RecordSource = "SELECT DescripcionCuentas, CodCuentas, TipoCuenta From Cuentas WHERE  (DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY DescripcionCuentas"
'                 DtaProductos.Refresh
'                 If Lectura = "%" Then
'                  Respuesta = ""
'                  Lectura = ""
'                 End If
'              Else
'                Respuesta = "" & Respuesta & Lectura & ""
'                DtaProductos.RecordSource = "SELECT DescripcionCuentas, CodCuentas, TipoCuenta From Cuentas WHERE  (DescripcionCuentas LIKE '" & Respuesta & "%') ORDER BY DescripcionCuentas"
'                DtaProductos.Refresh
'              End If
'                Me.DbgrProducto.Columns(0).Width = 6000
''                Me.DbgrProducto.Columns(4).NumberFormat = "##,##0.00"
'                Me.DbgrProducto.Columns(1).Width = 1500
'                Me.DbgrProducto.Columns(2).Width = 1500
'                Orden = True
'
'          End If
'
'
'         Me.DbgrProducto.MarqueeStyle = dbgHighlightCell
'
'     End Select
'     Me.Caption = "Buscar:"
''Else
'    Origen = Origen & Lectura
'    If Lectura = " " Then
'      Origen = Mid$(Origen, 1, Len(Origen) - 1)
'      Origen = Origen & "_"
'    End If
'    Me.Caption = "Buscar: " & Origen
'
'  End If
'
'
'Exit Sub
'TipoErrs:
'  ControlErrores
'  Unload Me
'
'
'
'
'End Sub
