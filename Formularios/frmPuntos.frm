VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmPuntos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cat�logo: Puntos"
   ClientHeight    =   4845
   ClientLeft      =   -75
   ClientTop       =   345
   ClientWidth     =   6660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6660
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   2040
      TabIndex        =   14
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      Begin VB.CommandButton cmdGrupo 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5760
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtGrupo 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   3645
      End
      Begin VB.TextBox txtIdG 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   675
      End
      Begin TrueOleDBGrid80.TDBGrid tdbgPuntos 
         Height          =   1815
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   6015
         _ExtentX        =   10610
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
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42,.ellipsis=0"
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
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12,.ellipsis=0"
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
      Begin VB.TextBox txtDescrip 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   4365
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "frmPuntos.frx":0000
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "frmPuntos.frx":0062
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "frmPuntos.frx":00D6
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmPuntos.frx":0140
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtId 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   675
      End
      Begin VB.TextBox txtCantPts 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1320
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmPuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnx As New ADODB.Connection
Private rs As New ADODB.Recordset
Private sql As String
Private modal As Boolean
Private getVal As Boolean
Private Id As Integer

Private Sub cmdAgregar_Click()
On Error GoTo erragr
If Me.cmdAgregar.Caption = "&Agregar" Then
    Me.tdbgPuntos.Enabled = False
    Me.cmdAgregar.Caption = "&Grabar"
    Me.cmdBorrar.Caption = "C&ancelar"
    Me.cmdEditar.Caption = "Save..."
    Me.cmdEditar.Enabled = False
    Me.cmdGrupo.Enabled = True
    Me.txtDescrip.Locked = False
    Me.txtCantPts.Locked = False
    Me.txtId.Text = ""
    Me.txtIdG.Text = ""
    Me.txtGrupo.Text = ""
    Me.txtDescrip.Text = ""
    Me.txtCantPts.Text = ""
    Me.txtDescrip.SetFocus
Else        'CODIGO PARA GRABAR DATOS
    If Trim(Me.txtIdG.Text) = "" Then
        MsgBox "Seleccione un grupo de puntos", vbInformation
        Me.cmdGrupo.SetFocus
        Exit Sub
    ElseIf Trim(UCase(Me.txtGrupo.Text)) = "ANTIG�EDAD" Then
        MsgBox "Operaci�n Cancelada, no puede agregar mas puntos a este grupo", vbInformation
        Me.cmdGrupo.SetFocus
        Exit Sub
    ElseIf Me.txtDescrip = "" Then
        MsgBox "La descripci�n del punto es requerida", vbInformation
        Me.txtDescrip.SetFocus
        Exit Sub
    ElseIf Me.txtCantPts = "" Then
        MsgBox "Digite la cantidad de puntos", vbInformation
        Me.txtCantPts.SetFocus
        Exit Sub
    End If
    
    Dim filt As String
    filt = "Id = " & val(Me.txtId)
    
    If Me.cmdEditar.Caption = "Save..." Then
        sql = "INSERT INTO PUNTOS(GRUPO, DESCRIPCION, CANTPTS) " & _
                "VALUES(" & val(Me.txtIdG) & ", '" & Me.txtDescrip.Text & "', '" & val(Me.txtCantPts) & "')"
    Else
        sql = "UPDATE PUNTOS SET GRUPO = '" & val(Me.txtIdG) & "', DESCRIPCION = '" & Me.txtDescrip & "', CANTPTS = " & val(Me.txtCantPts) & _
                " WHERE ID = " & val(rs!Id)
    End If
        
    cnx.Execute sql
    
    Call rsUpdate
    Call CmdBorrar_Click
    If Not filt = "" Then rs.Find filt
    
End If

Exit Sub
erragr:
    If Err.Number = -2147217873 Then
        MsgBox "Operaci�n cancelada, El identificador ya esta siendo usado", vbInformation
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub CmdBorrar_Click()
On Error GoTo errbor
If Me.cmdBorrar.Caption = "&Borrar" Then
    If rs.eof Then MsgBox "Operaci�n Cancelada, no existen registros", vbInformation: Exit Sub
    If UCase(rs!grupo) = "ANTIG�EDAD" Then MsgBox "Operaci�n Cancelada, el registro no puede ser eliminado", vbInformation: Exit Sub
    If MsgBox("�Desea eliminar el registro?", vbYesNo) = vbYes Then
        sql = "DELETE FROM PUNTOS WHERE ID = " & val(rs!Id)
        cnx.Execute sql
        Call rsUpdate
        MsgBox "Registro eliminado", vbInformation
    End If
Else        'CODIGO PARA CANCELAR AGREGACION
    Me.tdbgPuntos.Enabled = True
    Me.cmdAgregar.Caption = "&Agregar"
    Me.cmdBorrar.Caption = "&Borrar"
    Me.cmdEditar.Caption = "&Editar"
    Me.cmdEditar.Enabled = True
    Me.cmdGrupo.Enabled = False
    Me.txtCantPts.Locked = True
    Me.txtDescrip.Locked = True
    Call tdbgPuntos_RowColChange(Me.tdbgPuntos.Row, Me.tdbgPuntos.col)
End If

Exit Sub
errbor:
    If Err.Number = -2147217873 Then
        MsgBox "Operaci�n cancelada, Antes elimine los registros relacionados", vbInformation
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub CmdCerrar_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cmdGrupo_Click()
On Error Resume Next
modal = True
If Me.cmdEditar.Caption = "Update..." Then frmPuntosGrupo.prId = val(rs!idg)
frmPuntosGrupo.prGetVal = True
frmPuntosGrupo.Show

End Sub

Private Sub cmdEditar_Click()
On Error GoTo erredit
If rs.eof Then MsgBox "Operaci�n Cancelada, no existen registros", vbInformation: Exit Sub
If UCase(rs!grupo) = "ANTIG�EDAD" Then MsgBox "Operaci�n Cancelada, el registro no puede ser editado", vbInformation: Exit Sub
Me.tdbgPuntos.Enabled = False
Me.cmdAgregar.Caption = "&Grabar"
Me.cmdBorrar.Caption = "C&ancelar"
Me.cmdEditar.Caption = "Update..."
Me.cmdEditar.Enabled = False
Me.cmdGrupo.Enabled = True
Me.txtCantPts.Locked = False
Me.txtDescrip.Locked = False
Me.txtDescrip.SetFocus
Exit Sub
erredit:
    MsgBox Err.Description
End Sub

Private Sub Form_Activate()
On Error GoTo errAct
   
If modal Then frmPuntosGrupo.SetFocus
   
MDIPrimero.Skin1.ApplySkin hWnd

Exit Sub
errAct:
    MsgBox Err.Description

End Sub

Private Sub Form_Load()
On Error GoTo errload

If cnx.State = adStateClosed Then
'    sql = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PRUEBA;Data Source=WEBMASTER\SQL2005"
    cnx.ConnectionString = Conexion
    cnx.Open
End If

With rs
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    sql = "SELECT P.Id, G.GRUPO, P.DESCRIPCION, P.CANTPTS, G.ID AS IDG " & _
            "FROM PUNTOS P INNER JOIN PUNTOSGRUPO G ON P.GRUPO = G.ID"
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
    
    Me.tdbgPuntos.DataSource = rs
End With

Me.cmdGrupo.Enabled = False
Call rsUpdate
sql = "Id = " & Id
rs.Find sql

Me.Top = (MDIPrimero.ScaleHeight / 2) - (Me.Height / 2)
Me.Left = (MDIPrimero.ScaleWidth / 2) - (Me.Width / 2)
    
Exit Sub
errload:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set cnx = Nothing
Set rs = Nothing

End Sub

Private Sub tdbgPuntos_FilterChange()
On Error GoTo errfilt
Dim cols As TrueOleDBGrid80.Columns
Dim X As Integer
Set cols = Me.tdbgPuntos.Columns
X = Me.tdbgPuntos.col
Me.tdbgPuntos.HoldFields
rs.Filter = getFilter(cols, rs)
Me.tdbgPuntos.col = X

Exit Sub
errfilt:
    MsgBox Err.Description
End Sub

Private Sub tdbgPuntos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo errrow
If Not rs.eof Then
    Me.txtId = IIf(IsNull(rs!Id), "", rs!Id)
    Me.txtIdG = IIf(IsNull(rs!idg), "", rs!idg)
    Me.txtGrupo = IIf(IsNull(rs!grupo), "", rs!grupo)
    Me.txtDescrip = IIf(IsNull(rs!descripcion), "", rs!descripcion)
    Me.txtCantPts = IIf(IsNull(rs!cantpts), "", rs!cantpts)
    
End If

If Not getVal Then Exit Sub
Exit Sub
errrow:
    MsgBox Err.Description
End Sub

Public Property Let prGetVal(val As Boolean)
getVal = val
End Property

Public Property Let prModal(val As Boolean)
modal = val
End Property

Private Sub rsUpdate()
rs.Requery
Me.tdbgPuntos.ReBind
Me.tdbgPuntos.Refresh
Me.tdbgPuntos.Columns(4).Visible = False
Me.tdbgPuntos.Columns(0).Width = 500
End Sub

Public Property Let prId(ByVal val As Integer)
Id = val
End Property


