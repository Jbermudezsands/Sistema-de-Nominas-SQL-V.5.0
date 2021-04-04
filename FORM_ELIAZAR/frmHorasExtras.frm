VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form frmHorasExtras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorización de Horas Extras"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   7560
      Width           =   11055
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   9960
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraSolicitud 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   11055
      Begin VB.CommandButton cmdBorrarCta 
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
         Left            =   10560
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdAgregarCta 
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
         Left            =   10560
         TabIndex        =   1
         Tag             =   "1"
         Top             =   240
         Width           =   375
      End
      Begin TrueOleDBGrid80.TDBGrid tdbgPuntos 
         Height          =   3615
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   6376
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
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
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmHorasExtras.frx":0000
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ilnode 
      Left            =   3840
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorasExtras.frx":006C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorasExtras.frx":0830
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvNominas 
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4895
      _Version        =   393217
      Indentation     =   353
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ilnode"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmHorasExtras.frx":1640
      TabIndex        =   9
      Top             =   3240
      Width           =   1575
   End
   Begin MSComCtl2.MonthView mvFecha 
      Height          =   2820
      Left            =   6960
      TabIndex        =   10
      Top             =   360
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   -2147483647
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   75169793
      TitleBackColor  =   -2147483629
      CurrentDate     =   40749
   End
End
Attribute VB_Name = "frmHorasExtras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnx As New ADODB.Connection
Private rs As New ADODB.Recordset
Private rs3 As New ADODB.Recordset
Private rs4 As New ADODB.Recordset
Private sql As String
Private modal As Boolean

Private Sub cmdAddArea_Click()
On Error Resume Next
frmPuntosGrupo.Show
End Sub

Private Sub cmdAddPA_Click()
On Error Resume Next
frmPuntos.Show
End Sub

Private Sub cmdAgregarCta_Click()
On Error GoTo erradd
If rs3.EOF Then MsgBox "Operación Cancelada, no existen registros", vbInformation: Exit Sub

rs3.MoveFirst
Do While Not rs3.EOF
    If Format$(Me.mvFecha.Value, "DD/MM/YYYY") < Format$(rs3!Fecha, "DD/MM/YYYY") Then
        MsgBox "La fecha de aprobacion debe ser mayor o igual a la fecha de solicitud", vbInformation
        Exit Sub
    Else
        Me.tdbgPuntos.Columns(6).Value = True
        Me.tdbgPuntos.Columns(8).Value = Format$(Me.mvFecha.Value, "DD/MM/YYYY")
    End If
    rs3.MoveNext
Loop
rs3.MoveFirst

Exit Sub
erradd:
    MsgBox Err.Description

End Sub

Private Sub cmdBorrarCta_Click()
On Error GoTo erradd
If rs3.EOF Then MsgBox "Operación Cancelada, no existen registros", vbInformation: Exit Sub

rs3.MoveFirst
Do While Not rs3.EOF
    Me.tdbgPuntos.Columns(6).Value = False
    Me.tdbgPuntos.Columns(8).Value = Null
    rs3.MoveNext
Loop
rs3.MoveFirst

Exit Sub
erradd:
    MsgBox Err.Description
End Sub

Private Sub CmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo errgrb
Dim filt As String
Dim rs2 As New ADODB.Recordset

rs3.MoveFirst
rs4.MoveFirst
Do While Not rs3.EOF
    If rs3!aprobar <> rs4!aprobar Then
        sql = "select count(*) as horas " & _
                "from empleado e inner join nomina n on e.codtiponomina = n.codtiponomina " & _
                    "inner join horasextras h on (n.numnomina = h.numnomina and e.codempleado = h.codempleado) " & _
                "where n.activa = 1 and n.codtiponomina = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' and e.codempleado = " & rs3!CodEmpleado
        
        With rs2
            If .State = adStateOpen Then .Close
            .CursorLocation = adUseClient
            .Open sql, cnx, adOpenDynamic, adLockOptimistic
        End With
        
        If rs2!Horas = 0 Then
            sql = "INSERT INTO HORASEXTRAS " & _
                    " select id = (select max(id)+1 from horasextras), " & _
                    " codempleado = " & rs3!CodEmpleado & ", " & _
                    " numnomina = (select numnomina from nomina where activa = 1 and codtiponomina = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "'), " & _
                    " canthoras = (SELECT  SUM(ROUND(DATEDIFF(MI, CAST(SUBSTRING(CAST(AP.HSALIDA AS NVARCHAR(19)),1,11) + SUBSTRING(CAST(N.HSALIDA AS NVARCHAR(19)),12,8) AS DATETIME),  AP.HSALIDA) / 60.0,2)) " & _
                                    "FROM   EMPLEADO E INNER JOIN _ACTIVIDADESPRODUCCION AP ON E.CODEMPLEADO = AP.CODEMPLEADO " & _
                                            " INNER JOIN TIPONOMINA N ON E.CODTIPONOMINA = N.CODTIPONOMINA " & _
                                    "WHERE   E.CODEMPLEADO = " & rs3!CodEmpleado & " AND N.CodTipoNomina = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' AND AP.Eliminar = 1 AND AP.HEXTRAS = 1 AND " & _
                                            "AP.FECHA BETWEEN ISNULL((select FECHANOMINAINI from nomina where codtiponomina = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' AND FECHANOMINA = (SELECT DATEADD(d, -1, FECHANOMINAINI) FROM NOMINA WHERE CODTIPONOMINA = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' AND ACTIVA = 1)), (SELECT FECHANOMINAINI FROM NOMINA WHERE CODTIPONOMINA = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' AND ACTIVA = 1)) AND (SELECT FECHANOMINA FROM NOMINA WHERE CODTIPONOMINA = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' AND ACTIVA = 1) " & _
                                            "AND ROUND(DATEDIFF(MI, CAST(SUBSTRING(CAST(AP.HSALIDA AS NVARCHAR(19)),1,11) + SUBSTRING(CAST(N.HSALIDA AS NVARCHAR(19)),12,8) AS DATETIME),  AP.HSALIDA) / 60.0,2) > 0), " & _
                    "pagada = 0 "
        Else
            sql = "UPDATE HORASEXTRAS " & _
                    " SET canthoras = (SELECT  SUM(ROUND(DATEDIFF(MI, CAST(SUBSTRING(CAST(AP.HSALIDA AS NVARCHAR(19)),1,11) + SUBSTRING(CAST(N.HSALIDA AS NVARCHAR(19)),12,8) AS DATETIME),  AP.HSALIDA) / 60.0,2)) " & _
                                    "FROM   EMPLEADO E INNER JOIN _ACTIVIDADESPRODUCCION AP ON E.CODEMPLEADO = AP.EMPLEADO " & _
                                            " INNER JOIN TIPONOMINA N ON E.CODTIPONOMINA = N.CODTIPONOMINA " & _
                                    "WHERE   E.CODEMPLEADO = " & rs3!CodEmpleado & " AND N.CodTipoNomina = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' AND AP.Eliminar = 1 AND AP.HEXTRAS = 1 AND " & _
                                            "AP.FECHA BETWEEN ISNULL((select FECHANOMINAINI from nomina where codtiponomina = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' AND FECHANOMINA = (SELECT DATEADD(d, -1, FECHANOMINAINI) FROM NOMINA WHERE CODTIPONOMINA = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' AND ACTIVA = 1)), (SELECT FECHANOMINAINI FROM NOMINA WHERE CODTIPONOMINA = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' AND ACTIVA = 1)) AND (SELECT FECHANOMINA FROM NOMINA WHERE CODTIPONOMINA = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' AND ACTIVA = 1) " & _
                                            "AND ROUND(DATEDIFF(MI, CAST(SUBSTRING(CAST(AP.HSALIDA AS NVARCHAR(19)),1,11) + SUBSTRING(CAST(N.HSALIDA AS NVARCHAR(19)),12,8) AS DATETIME),  AP.HSALIDA) / 60.0,2) > 0)" & _
                    "WHERE codempleado = " & rs3!CodEmpleado & " AND " & _
                    "       numnomina = (select numnomina from nomina where activa = 1 and codtiponomina = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "')"
        End If

        cnx.Execute sql
    End If
    rs3.MoveNext
    rs4.MoveNext
Loop
Call tvNominas_NodeClick(tvNominas.SelectedItem)
MsgBox "Actualización completada", vbInformation
Set rs2 = Nothing
Exit Sub
errgrb:
    MsgBox Err.Description
End Sub

Private Sub Form_Activate()
On Error Resume Next
MDIPrimero.Skin1.ApplySkin hWnd
End Sub

Private Sub Form_Load()
On Error GoTo errload
Dim rs2 As New ADODB.Recordset

Me.Top = (MDIPrimero.ScaleHeight / 2) - (Me.Height / 2)
Me.Left = (MDIPrimero.ScaleWidth / 2) - (Me.Width / 2)

Me.mvFecha.Value = Now

If cnx.State = adStateClosed Then
'    sql = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PRUEBA;Data Source=WEBMASTER\SQL2005"
    cnx.ConnectionString = Conexion
    cnx.Open
End If

'NOMINAS
sql = "SELECT  [Nomina], TN.[CodTipoNomina], N.FECHANOMINAINI, N.FECHANOMINA From [dbo].[TipoNomina] TN inner join Nomina N ON TN.CODTIPONOMINA = N.CODTIPONOMINA Where TN.[Activa] = 1 AND N.ACTIVA = 1 "
With rs2
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

If Not rs2.EOF Then
    rs2.MoveFirst
    Do While Not rs2.EOF
        tvNominas.Nodes.Add , tvwLast, "A" & Trim(rs2!CodTipoNomina), Trim(rs2!nomina) & ": " & Format$(rs2!fechanominaini, "dd/mm/yyyy") & " - " & Format$(rs2!FechaNomina, "dd/mm/yyyy"), 1, 2
        rs2.MoveNext
    Loop
End If

If tvNominas.Nodes.Count > 0 Then tvNominas.Nodes(1).Selected = True

sql = "SELECT * FROM PUNTOSGRUPO"
        
With rs2
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

If Not rs2.EOF Then
    rs2.MoveFirst
    Do While Not rs2.EOF
'        tvActividades.Nodes.Add , tvwLast, "A" & Trim(Str(rs2!Id)), Trim(rs2!grupo), 1, 2
        rs2.MoveNext
    Loop
End If
Set rs2 = Nothing

sql = "SELECT P.Id, G.GRUPO, P.DESCRIPCION, P.CANTPTS, G.ID AS IDG " & _
            "FROM PUNTOS P INNER JOIN PUNTOSGRUPO G ON P.GRUPO = G.ID"
With rs
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
'        tvActividades.Nodes.Add "A" & Trim(Str(rs!idg)), tvwChild, "B" & Trim(Str(rs!Id)), Trim(rs!descripcion), 1, 2
        rs.MoveNext
    Loop
End If

'If tvActividades.Nodes.Count > 0 Then tvActividades.Nodes(1).Selected = True
Call tvNominas_NodeClick(tvNominas.SelectedItem)

Set rs2 = Nothing
Exit Sub
errload:
    MsgBox Err.Description
End Sub

Private Sub addchild(llave As Integer)
Dim rs2 As New ADODB.Recordset

sql = "SELECT [Llave], [Raiz], [Codigo], [Actividad] " & _
        "From [dbo].[Actividades] " & _
        "Where Raiz = " & llave & _
        " order by codigo"

With rs2
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

If Not rs2.EOF Then
    rs2.MoveFirst
    Do While Not rs2.EOF
'        tvActividades.Nodes.Add "A" & Trim(Str(rs2!raiz)), tvwChild, "A" & Trim(Str(rs2!llave)), Trim(rs2!codigo) & ".- " & Trim(rs2!actividad), 1, 2
        Call addchild(rs2!llave)
        rs2.MoveNext
    Loop
End If
Set rs2 = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errgrb
If rs3.BOF = rs3.EOF And rs3.EOF = False Then
    rs3.MoveFirst
    rs4.MoveFirst
    Do While Not rs3.EOF
        If rs3!aprobar <> rs4!aprobar Then
            Me.tdbgPuntos.Columns(6).Value = rs4!aprobar
            Me.tdbgPuntos.Columns(8).Value = IIf(IsNull(rs4!aprobado), Null, Format$(rs4!aprobado, "DD/MM/YYYY"))
        End If
        rs3.MoveNext
        rs4.MoveNext
    Loop
End If

Set cnx = Nothing
Set rs = Nothing
Set rs3 = Nothing
Set rs4 = Nothing

Exit Sub
errgrb:
    MsgBox Err.Description
End Sub

Public Property Let prModal(ByVal val As Boolean)
modal = val
End Property

Private Sub tdbgPuntos_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
If ColIndex <> 6 Then Cancel = 1
If ColIndex = 6 Then
    If Not Me.tdbgPuntos.Columns(6).Value Then
        Me.tdbgPuntos.Columns(8).Value = Null
    Else
        If Format$(Me.mvFecha.Value, "DD/MM/YYYY") < Format$(rs3!Fecha, "DD/MM/YYYY") Then
            MsgBox "La fecha de aprobacion debe ser mayor o igual a la fecha de solicitud", vbInformation
            Cancel = 1
        Else
            Me.tdbgPuntos.Columns(8).Value = Format$(Me.mvFecha.Value, "DD/MM/YYYY")
        End If
    End If
End If
End Sub

Private Sub tdbgPuntos_FilterChange()
'Gets called when an action is performed on the filter bar
Dim col As TrueOleDBGrid80.Column
Dim cols As TrueOleDBGrid80.Columns

'On Error GoTo errHandler
On Error Resume Next
Set cols = tdbgPuntos.Columns
Dim c As Integer

c = tdbgPuntos.col
tdbgPuntos.HoldFields
sql = rs3.Filter
rs3.Filter = getFilter(col, cols)

tdbgPuntos.col = c
tdbgPuntos.EditActive = True
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
        Select Case rs3.Fields(X).Type
        Case adVarWChar, adVarChar: tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
        Case adInteger, adNumeric: tmp = tmp & "[" & col.DataField & "] = " & col.FilterText
        Case adDBTimeStamp: tmp = tmp & "[" & col.DataField & "] = #" & col.FilterText & "#"
        End Select
    End If
    X = X + 1
Next col
getFilter = tmp

End Function

Private Sub tvNominas_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo errtv

Dim Fecha As Date
Fecha = CDate(Mid(tvNominas.SelectedItem.Text, Len(tvNominas.SelectedItem.Text) - 22, 10))
If Fecha > Me.mvFecha.MaxDate Then Me.mvFecha.MaxDate = Fecha + 1
Me.mvFecha.MinDate = Fecha
Me.mvFecha.MaxDate = CDate(Right(tvNominas.SelectedItem.Text, 10))
Me.mvFecha.Value = Me.mvFecha.MinDate

'LAS HORAS EXTRAS MOSTRADAS SON LAS DEL PERIODO ANTERIOR Y LAS DEL PERIODO ACTUAL
sql = "SELECT   Empleado.CodEmpleado1 AS Número, CASE WHEN Nombre2 IS NULL " & _
      "               THEN Nombre1 + ' ' + Apellido1 + ' ' + Apellido2 WHEN Nombre2 = '' THEN Nombre1 + ' ' + Apellido1 + ' ' + Apellido2 " & _
      "                 ELSE Nombre1 + ' ' + Nombre2 + ' ' + Apellido1 + ' ' + Apellido2 END AS Empleado, " & _
      "         _ActividadesProduccion.Fecha, " & _
      "         CAST(SUBSTRING(CAST(_ActividadesProduccion.HSalida AS NVARCHAR(19)), 1, 11) + SUBSTRING(CAST(TipoNomina.HSalida AS NVARCHAR(19)), 12, 8) AS DATETIME) AS Parámetro, " & _
      "         _ActividadesProduccion.HSalida AS [H. Salida], " & _
      "         ISNULL(ROUND(DATEDIFF(MI, CAST(SUBSTRING(CAST(_ActividadesProduccion.HSalida AS NVARCHAR(19)), 1, 11) + SUBSTRING(CAST(TipoNomina.HSalida AS NVARCHAR(19)), 12, 8) " & _
      "                 AS DATETIME), _ActividadesProduccion.HSalida) / 60.0, 2), case when _ActividadesProduccion.CantidadHoras - DATEDIFF(hh, TipoNomina.HEntrada, TipoNomina.HSalida) <= 0 then 0 " & _
      "         else _ActividadesProduccion.CantidadHoras - DATEDIFF(hh, TipoNomina.HEntrada, TipoNomina.HSalida) end) AS HE, _ActividadesProduccion.HExtras AS Aprobar, Empleado.CodEmpleado, " & _
      "                _ActividadesProduccion.FechaAprobado " & _
      "FROM         _ActividadesProduccion INNER JOIN " & _
      "                Empleado ON _ActividadesProduccion.CodEmpleado = Empleado.CodEmpleado INNER JOIN " & _
      "                TipoNomina ON Empleado.CodTipoNomina = TipoNomina.CodTipoNomina " & _
        "WHERE  TipoNomina.CodTipoNomina = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' AND _ActividadesProduccion.Eliminar = 1 AND " & _
        "       (_ActividadesProduccion.FECHA BETWEEN ISNULL((select FECHANOMINAINI from nomina where codtiponomina = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' AND FECHANOMINA = (SELECT DATEADD(d, -1, FECHANOMINAINI) FROM NOMINA WHERE CODTIPONOMINA = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' AND ACTIVA = 1)), " & _
        "                               (SELECT FECHANOMINAINI FROM NOMINA WHERE CODTIPONOMINA = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' AND ACTIVA = 1)) " & _
        "                        AND (SELECT FECHANOMINA FROM NOMINA WHERE CODTIPONOMINA = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' AND ACTIVA = 1) " & _
        "       AND ROUND(DATEDIFF(MI, CAST(SUBSTRING(CAST(_ActividadesProduccion.HSALIDA AS NVARCHAR(19)),1,11) + SUBSTRING(CAST(TipoNomina.HSALIDA AS NVARCHAR(19)),12,8) AS DATETIME),  _ActividadesProduccion.HSALIDA) / 60.0,2) > 0 " & _
        "       OR _ActividadesProduccion.CantidadHoras - DATEDIFF(hh, TipoNomina.HEntrada, TipoNomina.HSalida) >= 0) " & _
        "ORDER BY Empleado, _ActividadesProduccion.Fecha"

With rs3
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
    
    Set rs4 = rs3.Clone(adLockReadOnly)
    rs3.Requery
End With

Me.tdbgPuntos.DataSource = rs3
Me.tdbgPuntos.Columns(0).Width = 1000
Me.tdbgPuntos.Columns(1).Width = 3000
Me.tdbgPuntos.Columns(2).Width = 1100
Me.tdbgPuntos.Columns(3).Width = 1100
Me.tdbgPuntos.Columns(4).Width = 1100
Me.tdbgPuntos.Columns(5).Width = 700
Me.tdbgPuntos.Columns(6).Width = 700
Me.tdbgPuntos.Columns(6).ValueItems.Presentation = dbgCheckBox
Me.tdbgPuntos.Columns(7).Visible = False
Me.tdbgPuntos.Columns(8).Width = 1100

Me.tdbgPuntos.Columns(3).ValueItems.Translate = True
Me.tdbgPuntos.Columns(4).ValueItems.Translate = True

Dim item As New TrueOleDBGrid80.ValueItem

If rs3.BOF = rs3.EOF And rs3.EOF = True Then Exit Sub
rs3.MoveFirst
Do While Not rs3.EOF
    If Not IsNull(rs3!parámetro) Then
        item.Value = rs3!parámetro
        item.DisplayValue = Format$(rs3!parámetro, "HH:MM:SS AMPM")
        Me.tdbgPuntos.Columns(3).ValueItems.Add item
    End If
    If Not IsNull(rs3.Fields("H. Salida").Value) Then
        item.Value = rs3.Fields("H. Salida").Value
        item.DisplayValue = Format$(rs3.Fields("H. Salida").Value, "HH:MM:SS AMPM")
        Me.tdbgPuntos.Columns(4).ValueItems.Add item
    End If
    rs3.MoveNext
Loop
rs3.MoveFirst

Exit Sub
errtv:
    MsgBox Err.Description
End Sub


