VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmPuntosAprobar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Puntos: Solicitud / Aprobación"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   6990
   Begin VB.CommandButton cmdAprobar 
      Caption         =   "&Aprobar"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   7080
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   16
      Top             =   4800
      Width           =   6615
      Begin MSComCtl2.DTPicker dtpSolicitud 
         Height          =   300
         Left            =   2040
         TabIndex        =   24
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   70451201
         CurrentDate     =   40729
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   23
         Top             =   960
         Width           =   330
      End
      Begin VB.TextBox txtJustificacion 
         Height          =   525
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   4965
      End
      Begin VB.TextBox txtDocumento 
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   960
         Width           =   4485
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosAprobar.frx":0000
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosAprobar.frx":0078
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosAprobar.frx":00E8
         TabIndex        =   21
         Top             =   1320
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosAprobar.frx":016A
         TabIndex        =   22
         Top             =   1680
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpAprobacion 
         Height          =   300
         Left            =   2040
         TabIndex        =   25
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         Format          =   70451201
         CurrentDate     =   40729
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   7080
      Width           =   975
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
      TabIndex        =   7
      Top             =   1920
      Width           =   6615
      Begin VB.TextBox txtSolicitados 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2280
         Width           =   765
      End
      Begin VB.TextBox txtAprobados 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2280
         Width           =   765
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   9
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   8
         Tag             =   "1"
         Top             =   360
         Width           =   375
      End
      Begin TrueOleDBGrid80.TDBGrid tdbgPts 
         Height          =   1815
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   5775
         _ExtentX        =   10186
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
         OleObjectBlob   =   "frmPuntosAprobar.frx":01EE
         TabIndex        =   14
         Top             =   2280
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "frmPuntosAprobar.frx":0270
         TabIndex        =   15
         Top             =   2280
         Width           =   1335
      End
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
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtDpto 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   4965
      End
      Begin VB.TextBox txtempleado 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   4965
      End
      Begin VB.TextBox txtId 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   1605
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosAprobar.frx":02EE
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosAprobar.frx":0358
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "frmPuntosAprobar.frx":03C6
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmPuntosAprobar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Id As Integer
Private rs As New ADODB.Recordset
Private cnx As New ADODB.Connection
Private Sql As String
Private modal As Boolean                    'Para abrir otros formularios en modal
Private Pts As Integer                       'Guarda el id del punto a agregar
Private agregar As Boolean
'Para guardar el punto
Public Property Let prId(ByVal val As Integer)
Id = val
End Property


Private Sub cmdAgregar_Click()
On Error GoTo erragr
'modal = True
'agregar = True
'frmPuntos.prGetVal = True
'frmPuntos.Show

FrmConsultarPuntos.txtempleado.Text = Me.txtempleado.Text
FrmConsultarPuntos.txtDpto.Text = Me.txtDpto.Text
FrmConsultarPuntos.txtId.Text = Me.txtId.Text
FrmConsultarPuntos.prId = Id
FrmConsultarPuntos.Show 1

rs.Requery
Me.tdbgPts.ReBind
Me.tdbgPts.Refresh

Me.txtAprobados = 0
Me.txtSolicitados = 0
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        If rs!aprobado Then Me.txtAprobados.Text = "" & val(Me.txtAprobados.Text) + rs!Puntos
        If Not rs!aprobado Then Me.txtSolicitados.Text = "" & val(Me.txtSolicitados.Text) + rs!Puntos
        rs.MoveNext
    Loop
    rs.MoveFirst
    
    Dim X As Integer
    For X = 5 To Me.tdbgPts.Columns.Count - 1
        Me.tdbgPts.Columns(X).Visible = False
    Next
    Me.tdbgPts.Columns(0).Width = 500
    Me.tdbgPts.Columns(4).ValueItems.Presentation = dbgCheckBox
    
End If

Exit Sub
erragr:
    MsgBox Err.Description
End Sub

Private Sub cmdAprobar_Click()
On Error GoTo errgra
If rs.EOF Then MsgBox "Operación Cancelada, no existen registros", vbInformation: Exit Sub
If Me.txtJustificacion.Text = "" Then
    MsgBox "La justificación es requerida para aprobar los puntos", vbInformation
    Me.txtJustificacion.SetFocus
    Exit Sub
End If


Sql = "Update [dbo].[PuntosEmpleado] SET [Aprobado] = 1, [Justificacion] = '" & Me.txtJustificacion & "' " & _
        ", [Documento] = '" & Me.txtDocumento & "', [FechaAprobado] = '" & Format$(Me.dtpAprobacion.Value, "yyyyMMdd") & "' " & _
        "WHERE [Empleado] = " & frmPuntosAdmin.prId & " and [Puntos] = " & rs!Id
cnx.Execute Sql
MsgBox "Solicitud aprobada", vbInformation
Call Form_Activate
Exit Sub
errgra:
    MsgBox Err.Description
End Sub

Private Sub cmdCerrar_Click()
On Error Resume Next
Set frmPuntos = Nothing
Unload Me
End Sub

Private Sub cmdEliminar_Click()
On Error GoTo erreli
If MsgBox("Esta seguro que desea eliminar la solicitud de puntos", vbYesNo) = vbYes Then
    Sql = "DELETE FROM PUNTOSEMPLEADO WHERE empleado = " & frmPuntosAdmin.prId & " and puntos = " & rs!Id
    cnx.Execute Sql

    rs.Requery
    Me.tdbgPts.ReBind
    Me.tdbgPts.Refresh
        
    If Not rs.EOF Then
        rs.MoveFirst
        Me.txtAprobados = "0"
        Me.txtSolicitados = "0"
        Do While Not rs.EOF
            If rs!aprobado Then Me.txtAprobados.Text = "" & val(Me.txtAprobados.Text) + rs!Puntos
            If Not rs!aprobado Then Me.txtSolicitados.Text = "" & val(Me.txtSolicitados.Text) + rs!Puntos
            rs.MoveNext
        Loop
        
        Dim X As Integer
        For X = 4 To Me.tdbgPts.Columns.Count - 1
            Me.tdbgPts.Columns(X).Visible = False
        Next
    End If
    
End If
Exit Sub
erreli:
    MsgBox Err.Description
End Sub

Private Sub dtpAprobacion_Change()
If Me.dtpAprobacion < Me.dtpSolicitud Then
    MsgBox "La fecha de aprobacion debe ser mayor o igual a la fecha de solicitud", vbInformation
    Me.dtpAprobacion.Value = Now
    Me.dtpAprobacion.SetFocus
End If
End Sub

Private Sub Form_Activate()
On Error GoTo errAct
   
If modal Then
    frmPuntos.SetFocus
ElseIf agregar Then
    Sql = "INSERT INTO PUNTOSEMPLEADO(EMPLEADO, PUNTOS, APROBADO, FECHASOLICITUD) VALUES(" & frmPuntosAdmin.prId & ", " & Pts & ", 0, '" & Format$(Me.dtpSolicitud.Value, "yyyymmdd") & "')"
    cnx.Execute Sql
    agregar = False
End If

rs.Requery
Me.tdbgPts.ReBind
Me.tdbgPts.Refresh

Me.txtAprobados = 0
Me.txtSolicitados = 0
If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
        If rs!aprobado Then Me.txtAprobados.Text = "" & val(Me.txtAprobados.Text) + rs!Puntos
        If Not rs!aprobado Then Me.txtSolicitados.Text = "" & val(Me.txtSolicitados.Text) + rs!Puntos
        rs.MoveNext
    Loop
    rs.MoveFirst
    
    Dim X As Integer
    For X = 5 To Me.tdbgPts.Columns.Count - 1
        Me.tdbgPts.Columns(X).Visible = False
    Next
    Me.tdbgPts.Columns(0).Width = 500
    Me.tdbgPts.Columns(4).ValueItems.Presentation = dbgCheckBox
    
End If

MDIPrimero.Skin1.ApplySkin hWnd
Exit Sub
errAct:
    If Err.Number = -2147217873 Then
        MsgBox "Operación cancelada, El identificador ya esta siendo usado", vbInformation
        agregar = False
    Else
        MsgBox Err.Description
        Unload Me
    End If
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
    
    'DATOS DEL EMPLEADO
    Sql = "SELECT CodEmpleado, CodEmpleado1, 'nombre' = " & _
            "       case " & _
            "           when Nombre2 is null then Nombre1 + ' ' + Apellido1 + ' ' + Apellido2 " & _
            "           else Nombre1 + ' ' + Nombre2 + ' ' + Apellido1 + ' ' + Apellido2 " & _
            "       end, " & _
            "       departamento " & _
            "FROM   Empleado e inner join Departamento d on e.coddepartamento = d.coddepartamento " & _
            "WHERE   CodEmpleado = " & frmPuntosAdmin.prId
            
    .Open Sql, cnx, adOpenDynamic, adLockOptimistic
    
    Me.txtId.Text = rs!CodEmpleado1
    Me.txtempleado.Text = rs!Nombre
    Me.txtDpto.Text = rs!departamento

    .Close
    
    'DATOS DE LOS PUNTOS DEL EMPLEADO
    Sql = "SELECT  P.ID, G.GRUPO, P.DESCRIPCION, P.CANTPTS AS PUNTOS, EP.APROBADO, EP.JUSTIFICACION, " & _
            "EP.DOCUMENTO, EP.DIRDOCUMENTO, EP.FECHASOLICITUD, EP.FECHAAPROBADO " & _
            "FROM    EMPLEADO E INNER JOIN PUNTOSEMPLEADO EP ON E.CODEMPLEADO = EP.EMPLEADO " & _
            "INNER JOIN PUNTOS P ON EP.PUNTOS = P.ID " & _
            "INNER JOIN PUNTOSGRUPO G ON P.GRUPO = G.ID " & _
            "WHERE   E.CODEMPLEADO = " & frmPuntosAdmin.prId & " " & _
            "ORDER BY EP.APROBADO DESC, G.GRUPO, P.DESCRIPCION"
            
    .Open Sql, cnx, adOpenDynamic, adLockOptimistic

    Me.tdbgPts.DataSource = rs
End With

dtpSolicitud = Now
dtpAprobacion = Now
Me.Top = (MDIPrimero.ScaleHeight / 2) - (Me.Height / 2)
Me.Left = (MDIPrimero.ScaleWidth / 2) - (Me.Width / 2)

Exit Sub
errload:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmPuntosAdmin.prModal = False
Set cnx = Nothing
Set rs = Nothing
Set rs = Nothing

End Sub

Public Property Let prModal(ByVal val As Boolean)
modal = val
End Property

Private Sub tdbgPts_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo errtdg
If Not rs.EOF Then
    Me.txtJustificacion = IIf(IsNull(rs!JUSTIFICACION), "", rs!JUSTIFICACION)
    Me.txtDocumento = IIf(IsNull(rs!documento), "", rs!documento)
    Me.dtpSolicitud = IIf(IsNull(rs!fechasolicitud), Now, rs!fechasolicitud)
    Me.dtpAprobacion = IIf(IsNull(rs!fechaaprobado), Now, rs!fechaaprobado)
    If rs!aprobado Then Me.cmdAprobar.Enabled = False
    If Not rs!aprobado Then Me.cmdAprobar.Enabled = True
    If rs!aprobado Then Me.cmdEliminar.Enabled = False
    If Not rs!aprobado Then Me.cmdEliminar.Enabled = True
End If
Exit Sub
errtdg:
    MsgBox Err.Description
End Sub

Public Property Let prAgregar(ByVal val As Boolean)
agregar = val
End Property

Public Property Let prPts(ByVal val As Integer)
Pts = val
End Property
