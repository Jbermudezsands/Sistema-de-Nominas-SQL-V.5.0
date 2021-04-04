VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmPuntosAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complementos Salariales"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   6930
   Begin VB.CommandButton cmdPuntos 
      Caption         =   "&Solicitud"
      Height          =   375
      Left            =   1320
      TabIndex        =   26
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtnumero 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   480
         Width           =   1605
      End
      Begin VB.TextBox txtempleado 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   840
         Width           =   4965
      End
      Begin VB.TextBox txtDpto 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1200
         Width           =   4965
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosAdmin.frx":0000
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosAdmin.frx":006A
         TabIndex        =   19
         Top             =   840
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosAdmin.frx":00D8
         TabIndex        =   20
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Frame fracomplementos 
      Caption         =   "Salario Ordinario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   6615
      Begin VB.TextBox txtPrecioPts 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtPorcentaje 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtCantPts 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtValPts 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtValPorcentaje 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtSalario 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosAdmin.frx":014E
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosAdmin.frx":01C8
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmPuntosAdmin.frx":023A
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosAdmin.frx":02A4
         TabIndex        =   13
         Top             =   1920
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "frmPuntosAdmin.frx":030C
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "frmPuntosAdmin.frx":0375
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   2880
         OleObjectBlob   =   "frmPuntosAdmin.frx":03D8
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   240
         X2              =   5880
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   240
         X2              =   5880
         Y1              =   1800
         Y2              =   1800
      End
   End
   Begin VB.Frame fraPuntos 
      Caption         =   "Puntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   6615
      Begin VB.TextBox txtAprobados 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2280
         Width           =   765
      End
      Begin VB.TextBox txtSolicitados 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2280
         Width           =   765
      End
      Begin TrueOleDBGrid80.TDBGrid tdbgPts 
         Height          =   1815
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3201
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
         AllowArrows     =   0   'False
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmPuntosAdmin.frx":043B
         TabIndex        =   28
         Top             =   2280
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "frmPuntosAdmin.frx":04BD
         TabIndex        =   30
         Top             =   2280
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   7440
      Width           =   975
   End
End
Attribute VB_Name = "frmPuntosAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New ADODB.Recordset
Private cnx As New ADODB.Connection
Private Sql As String
Private modal As Boolean                    'Para abrir otros formularios en modal
Private Id As Integer                       'Guarda el id del empleado

Private Sub cmdCerrar_Click()

Unload Me
End Sub
Private Sub CmdGrabar_Click()
On Error GoTo errgra
If val(Me.txtPorcentaje) < val(Me.txtPorcentaje.Tag) Then
    MsgBox "No puede definir un porcentaje menor al ya existente", vbInformation
    Me.txtPorcentaje = Me.txtPorcentaje.Tag
    Me.txtPorcentaje.SetFocus
    Exit Sub
End If
If MsgBox("Esta seguro de guardar los cambios", vbYesNo) = vbYes Then
    Sql = "UPDATE [dbo].[Empleado] SET [SalPorcentaje] = " & val(Me.txtPorcentaje) & ", [CantPts] = " & val(Me.txtCantPts) & ", [SueldoPeriodo] = " & val(0) & _
            " Where [CodEmpleado] = " & Id
            
    cnx.Execute Sql
    MsgBox "Actualización completada", vbInformation

End If
Exit Sub
errgra:
    MsgBox Err.Description
End Sub

Private Sub cmdPuntos_Click()
On Error GoTo erragr
modal = True
frmPuntosAprobar.prId = Id
frmPuntosAprobar.Show

Exit Sub
erragr:
    MsgBox Err.Description

End Sub

Private Sub Form_Activate()
On Error Resume Next
   
If modal Then frmPuntosAprobar.SetFocus

rs.Requery
Me.tdbgPts.ReBind
Me.tdbgPts.Refresh
Me.tdbgPts.Columns(4).Visible = False

Me.txtAprobados = "0"
Me.txtSolicitados = "0"
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        If rs!aprobado Then Me.txtAprobados.Text = "" & val(Me.txtAprobados.Text) + rs!Puntos
        If Not rs!aprobado Then Me.txtSolicitados.Text = "" & val(Me.txtSolicitados.Text) + rs!Puntos
        rs.MoveNext
    Loop
    Me.tdbgPts.Columns(3).ValueItems.Presentation = dbgCheckBox
End If

Me.txtCantPts.Text = Me.txtAprobados.Text
Me.txtValPorcentaje.Text = "" & Round(val(Me.txtSalario.Text) * val(Me.txtPorcentaje.Text) / 100, 2)
Me.txtValPts.Text = "" & val(Me.txtPrecioPts.Text) * val(Me.txtCantPts.Text)
Me.txtTotal = "" & val(Me.txtSalario.Text) + val(Me.txtValPorcentaje.Text) + val(Me.txtValPts.Text)


Exit Sub
errAct:
    MsgBox Err.Description
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo errload

MDIPrimero.Skin1.ApplySkin hWnd

If cnx.State = adStateClosed Then
'    sql = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PRUEBA;Data Source=WEBMASTER\SQL2005"
    cnx.ConnectionString = Conexion
    cnx.Open
End If

With rs
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    
    'DATOS DEL EMPLEADO
    Sql = "SELECT CodEmpleado, CodEmpleado1, 'nombre' = " & _
            "       case " & _
            "           when Nombre2 is null then Nombre1 + ' ' + Apellido1 + ' ' + Apellido2 " & _
            "           else Nombre1 + ' ' + Nombre2 + ' ' + Apellido1 + ' ' + Apellido2 " & _
            "       end, " & _
            "       departamento, salporcentaje " & _
            "FROM   Empleado e inner join Departamento d on e.coddepartamento = d.coddepartamento " & _
            "WHERE   CodEmpleado = " & Id
            
    .Open Sql, cnx, adOpenDynamic, adLockOptimistic
           
    Me.txtnumero.Text = IIf(IsNull(rs!CodEmpleado1), "", rs!CodEmpleado1)
    Me.txtempleado.Text = IIf(IsNull(rs!Nombre), "", rs!Nombre)
    Me.txtDpto.Text = IIf(IsNull(rs!departamento), "", rs!departamento)
    Me.txtPorcentaje.Text = IIf(IsNull(rs!SalPorcentaje), 0, rs!SalPorcentaje)
    Me.txtPorcentaje.Tag = Me.txtPorcentaje.Text
    
    .Close
    
    'DATOS DEL SALARIO MINIMO Y EL VALOR X PUNTOS
    Sql = "SELECT [SalarioMinimo], [ValorPts] FROM [dbo].[DatosEmpresa]"
    .Open Sql, cnx, adOpenDynamic, adLockOptimistic
    
    Me.txtSalario = IIf(IsNull(rs!SalarioMinimo), 0, rs!SalarioMinimo)
    Me.txtPrecioPts = IIf(IsNull(rs!valorpts), 0, rs!valorpts)
    
    .Close
    
    'DATOS DE TODOS LOS PUNTOS ADQUIRIDOS X EMPLEADO
    Sql = "SELECT  G.GRUPO, P.DESCRIPCION, P.CANTPTS AS PUNTOS, EP.APROBADO " & _
            "FROM    EMPLEADO E INNER JOIN PUNTOSEMPLEADO EP ON E.CODEMPLEADO = EP.EMPLEADO " & _
            "INNER JOIN PUNTOS P ON EP.PUNTOS = P.ID " & _
            "INNER JOIN PUNTOSGRUPO G ON P.GRUPO = G.ID " & _
            "WHERE   E.CODEMPLEADO = " & Id & " " & _
            "ORDER BY EP.APROBADO DESC, G.GRUPO, P.DESCRIPCION"
            
    .Open Sql, cnx, adOpenDynamic, adLockOptimistic
        
    Me.tdbgPts.DataSource = rs

End With

Me.Top = (MDIPrimero.ScaleHeight / 2) - (Me.Height / 2)
Me.Left = (MDIPrimero.ScaleWidth / 2) - (Me.Width / 2)

Exit Sub
errload:
    MsgBox Err.Description
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set cnx = Nothing
Set rs = Nothing

End Sub

Public Property Let prModal(ByVal val As Boolean)
modal = val
End Property

Public Property Let prId(ByVal val As Integer)
Id = val
End Property

Public Property Get prId() As Integer
prId = Id
End Property

Public Property Let prNumero(ByVal val As String)
Id = val
Me.txtnumero.Text = val
End Property

Private Sub txtPorcentaje_Change()
Call Form_Activate
End Sub
