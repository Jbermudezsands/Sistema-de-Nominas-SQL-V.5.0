VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProduccionAct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrador de Horas Laborales"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   15165
   Begin VB.Frame frahoras 
      Height          =   975
      Left            =   6120
      TabIndex        =   14
      Top             =   6960
      Width           =   8895
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
         Left            =   6360
         TabIndex        =   7
         Top             =   240
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
         Left            =   5880
         TabIndex        =   6
         Tag             =   "1"
         Top             =   240
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmProduccionAct.frx":0000
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpEntrada 
         Height          =   495
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16842754
         CurrentDate     =   40750
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   3240
         OleObjectBlob   =   "frmProduccionAct.frx":006C
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpSalida 
         Height          =   495
         Left            =   3960
         TabIndex        =   5
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16842754
         CurrentDate     =   40750
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   6120
      TabIndex        =   13
      Top             =   8160
      Width           =   8895
      Begin VB.CommandButton Command1 
         Caption         =   "&Programado vs Real"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   7800
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComCtl2.MonthView mvFecha 
      Height          =   2820
      Left            =   10680
      TabIndex        =   2
      Top             =   480
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
      StartOfWeek     =   16842753
      TitleBackColor  =   -2147483629
      CurrentDate     =   40559
      MaxDate         =   40574
      MinDate         =   40559
   End
   Begin VB.Frame Frame4 
      Height          =   3255
      Left            =   6120
      TabIndex        =   12
      Top             =   3600
      Width           =   8895
      Begin TrueOleDBGrid80.TDBGrid tdbgHoras 
         Height          =   2535
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4471
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
      Left            =   6120
      OleObjectBlob   =   "frmProduccionAct.frx":00D6
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ilnode 
      Left            =   1080
      Top             =   4320
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
            Picture         =   "frmProduccionAct.frx":0142
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProduccionAct.frx":0906
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvActividades 
      Height          =   8415
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   14843
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
   Begin MSComctlLib.TreeView tvNominas 
      Height          =   2775
      Left            =   6120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
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
      Height          =   2895
      Left            =   240
      OleObjectBlob   =   "frmProduccionAct.frx":1716
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmProduccionAct"
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

Private Sub cmdAgregarCta_Click()
On Error GoTo erradd

If Me.dtpSalida.Value <= Me.dtpEntrada.Value Then
    MsgBox "La hora de salida debe ser mayor a la hora de entrada", vbInformation
    Me.dtpSalida = DateAdd("h", 1, Me.dtpEntrada)
    Exit Sub
End If

If IsNull(rs4!hentrada) Then
    sql = "INSERT INTO [dbo].[ActividadesProduccion]([Codigo], [Empleado], [Fecha], [HEntrada], [HSalida], HExtras, Eliminar, [Raiz]) " & _
            "Values (" & val(Left(tvActividades.SelectedItem.Text, InStr(1, tvActividades.SelectedItem.Text, ".") - 1)) & _
            "   ,'" & rs4!CodEmpleado & "', '" & Format$(Me.mvFecha.Value, "yyyyMMdd") & "', '" & Format$(Me.dtpEntrada.Value, "yyyyMMdd hh:mm:ss") & "', '" & Format$(Me.dtpSalida.Value, "yyyyMMdd hh:mm:ss") & "', 0, 1, "
    
    If tvActividades.SelectedItem.FullPath = tvActividades.SelectedItem.Text Then
        sql = sql & "0)"
    Else
        sql = sql & val(Mid(tvActividades.SelectedItem.Parent.Key, 2)) & ")"
    End If
Else
    sql = "UPDATE ActividadesProduccion SET HEntrada = '" & Format$(Me.dtpEntrada.Value, "yyyyMMdd hh:mm:ss") & "', HSalida = '" & Format$(Me.dtpSalida.Value, "yyyyMMdd hh:mm:ss") & _
            "' WHERE Codigo = " & val(Left(tvActividades.SelectedItem.Text, InStr(1, tvActividades.SelectedItem.Text, ".") - 1)) & _
            " AND Empleado = '" & rs4!CodEmpleado & "' AND Fecha = '" & Format$(Me.mvFecha.Value, "yyyyMMdd") & "'"
    If tvActividades.SelectedItem.FullPath = tvActividades.SelectedItem.Text Then
        sql = sql & " AND Raiz = 0"
    Else
        sql = sql & " AND Raiz = " & val(Mid(tvActividades.SelectedItem.Parent.Key, 2))
    End If
End If

cnx.Execute sql
Call mvFecha_DateClick(Me.mvFecha.Value)
Exit Sub
erradd:
    MsgBox Err.Description
End Sub

Private Sub cmdBorrarCta_Click()
On Error GoTo erradd

sql = "DELETE FROM ActividadesProduccion " & _
        " WHERE Codigo = " & val(Left(tvActividades.SelectedItem.Text, InStr(1, tvActividades.SelectedItem.Text, ".") - 1)) & _
        " AND Empleado = '" & rs4!CodEmpleado & "' AND Fecha = '" & Format$(Me.mvFecha.Value, "yyyyMMdd") & "'"
If tvActividades.SelectedItem.FullPath = tvActividades.SelectedItem.Text Then
    sql = sql & " AND Raiz = 0"
Else
    sql = sql & " AND Raiz = " & val(Mid(tvActividades.SelectedItem.Parent.Key, 2))
End If

cnx.Execute sql
Call mvFecha_DateClick(mvFecha.Value)
Exit Sub
erradd:
    MsgBox Err.Description
End Sub

Private Sub CmdCerrar_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If modal Then frmNodes.SetFocus
'If Not modal Then Call tvActividades_Click

MDIPrimero.Skin1.ApplySkin hWnd
End Sub

Private Sub Form_Load()
On Error GoTo errload
Dim rs2 As New ADODB.Recordset

Me.Top = (MDIPrimero.ScaleHeight / 2) - (Me.Height / 2)
Me.Left = (MDIPrimero.ScaleWidth / 2) - (Me.Width / 2)

'mvFecha.Value = Now
dtpEntrada.Value = Now
dtpSalida.Value = Now

If cnx.State = adStateClosed Then
'    sql = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=PRUEBA;Data Source=WEBMASTER\SQL2005"
    cnx.ConnectionString = Conexion
    cnx.Open
End If

sql = "SELECT [Llave], [Raiz], [Codigo], [Actividad], [PagaCliente] " & _
        "From [dbo].[Actividades]"
With rs
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

'NOMINAS
sql = "SELECT  [Nomina], TN.[CodTipoNomina], N.FECHANOMINAINI, N.FECHANOMINA From [dbo].[TipoNomina] TN inner join Nomina N ON TN.CODTIPONOMINA = N.CODTIPONOMINA Where TN.[Activa] = 1 AND N.ACTIVA = 1 "
With rs2
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

If Not rs2.eof Then
    rs2.MoveFirst
    Do While Not rs2.eof
        tvNominas.Nodes.Add , tvwLast, "A" & Trim(rs2!CodTipoNomina), Trim(rs2!Nomina) & ": " & Format$(rs2!fechanominaini, "dd/mm/yyyy") & " - " & Format$(rs2!FechaNomina, "dd/mm/yyyy"), 1, 2
        rs2.MoveNext
    Loop
End If

If tvNominas.Nodes.Count > 0 Then tvNominas.Nodes(1).Selected = True


'ACTIVIDADES
sql = "SELECT [Llave], [Raiz], [Codigo], [Actividad] " & _
        "From [dbo].[Actividades] " & _
        "Where Raiz = 0 " & _
        "order by codigo"
        
With rs2
    If .State = adStateOpen Then .Close
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

If Not rs2.eof Then
    rs2.MoveFirst
    Do While Not rs2.eof
        tvActividades.Nodes.Add , tvwLast, "A" & Trim(Str(rs2!llave)), Trim(rs2!codigo) & ".- " & Trim(rs2!actividad), 1, 2
        Call addchild(rs2!llave)
        rs2.MoveNext
    Loop
End If

If tvActividades.Nodes.Count > 0 Then tvActividades.Nodes(1).Selected = True
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

If Not rs2.eof Then
    rs2.MoveFirst
    Do While Not rs2.eof
        tvActividades.Nodes.Add "A" & Trim(Str(rs2!raiz)), tvwChild, "A" & Trim(Str(rs2!llave)), Trim(rs2!codigo) & ".- " & Trim(rs2!actividad), 1, 2
        Call addchild(rs2!llave)
        rs2.MoveNext
    Loop
End If
Set rs2 = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set cnx = Nothing
Set rs = Nothing
Set rs3 = Nothing
Set rs4 = Nothing
End Sub

Private Sub mvFecha_DateClick(ByVal DateClicked As Date)
On Error GoTo errsem

If tvActividades.Nodes.Count = 0 Then
    MsgBox "Agregue actividades para continuar con las operaciones", vbInformation
    Exit Sub
ElseIf tvNominas.Nodes.Count = 0 Then
    MsgBox "Agregue o active las nominas para continuar con las operaciones", vbInformation
    Exit Sub
End If

If Me.tvActividades.SelectedItem.Children > 0 Then Exit Sub

sql = "SELECT E.[CodEmpleado1] AS 'Id', 'Empleado' = " & _
            "case when Nombre2 is null then Nombre1 + ' ' + Apellido1 + ' ' + Apellido2 " & _
                    "else Nombre1 + ' ' + Nombre2 + ' ' + Apellido1 + ' ' + Apellido2 " & _
            "end, [HEntrada], A.[HSalida], E.[CodEmpleado], N.HSALIDA AS SALNOM " & _
        "FROM [dbo].[ActividadesProduccion] A INNER JOIN [dbo].[Empleado] E ON A.Empleado = E.CodEmpleado " & _
            "INNER JOIN [dbo].[TipoNomina] N ON E.CodTipoNomina = N.CodTipoNomina " & _
        "WHERE codigo = " & Left(tvActividades.SelectedItem.Text, InStr(1, tvActividades.SelectedItem.Text, ".") - 1) & " AND E.CodTipoNomina = '" & Mid(Me.tvNominas.SelectedItem.Key, 2) & "' " & _
        "   AND FECHA = '" & Format$(Me.mvFecha.Value, "yyyyMMdd") & "' "
  
If tvActividades.SelectedItem.FullPath = tvActividades.SelectedItem.Text Then
    sql = sql & " AND raiz = 0 "
Else
    sql = sql & " AND raiz = " & Mid(tvActividades.SelectedItem.Parent.Key, 2)
End If

sql = sql & " UNION " & _
            "SELECT E.[CodEmpleado1] AS 'Id', 'Empleado' = " & _
                "case " & _
                    "when Nombre2 is null then Nombre1 + ' ' + Apellido1 + ' ' + Apellido2 " & _
                    "else Nombre1 + ' ' + Nombre2 + ' ' + Apellido1 + ' ' + Apellido2 " & _
                "end, NULL AS [HEntrada], NULL AS [HSalida], E.[CodEmpleado], N.HSALIDA AS SALNOM " & _
            "FROM [dbo].[Empleado] E INNER JOIN [dbo].[TipoNomina] N ON E.CodTipoNomina = N.CodTipoNomina " & _
            "WHERE E.CodTipoNomina = '" & Mid(Me.tvNominas.SelectedItem.Key, 2)
sql = sql & "' AND E.[CodEmpleado] NOT IN (SELECT E.[CodEmpleado] " & Mid(sql, InStr(1, sql, "FROM"), InStr(InStr(1, sql, "FROM"), sql, "UNION") - InStr(1, sql, "FROM")) & ") ORDER BY Empleado "
        
With rs4
    If .State = adStateOpen Then .Close
    .CursorLocation = adUseClient
    .Open sql, cnx, adOpenDynamic, adLockOptimistic
End With

Me.tdbgHoras.DataSource = rs4

Dim X As Integer
Me.tdbgHoras.Columns(1).Width = 2000
Me.tdbgHoras.Columns(4).Visible = False
Me.tdbgHoras.Columns(5).Visible = False
Me.tdbgHoras.Columns(2).ValueItems.Translate = True
Me.tdbgHoras.Columns(3).ValueItems.Translate = True

Me.dtpEntrada.Value = Me.mvFecha.Value
Me.dtpSalida.Value = Me.mvFecha.Value

If Me.tvActividades.SelectedItem.Children <= 0 And (rs4.BOF = rs4.eof And rs4.eof = False) Then
    Me.tdbgHoras.Enabled = True
    Me.frahoras.Enabled = True
Else
    Me.tdbgHoras.Enabled = False
    Me.frahoras.Enabled = False
    Exit Sub
End If

Dim item As New TrueOleDBGrid80.ValueItem
rs4.MoveFirst
Do While Not rs4.eof
    If Not IsNull(rs4!hentrada) Then
        item.Value = rs4!hentrada
        item.DisplayValue = Format$(rs4!hentrada, "HH:MM:SS AMPM")
        Me.tdbgHoras.Columns(2).ValueItems.Add item
    End If
    If Not IsNull(rs4!hsalida) Then
        item.Value = rs4!hsalida
        item.DisplayValue = Format$(rs4!hsalida, "HH:MM:SS AMPM")
        Me.tdbgHoras.Columns(3).ValueItems.Add item
    End If
    rs4.MoveNext
Loop
rs4.MoveFirst

Exit Sub
errsem:
    MsgBox Err.Description
End Sub

Private Sub tdbgHoras_FilterChange()
'Gets called when an action is performed on the filter bar
Dim col As TrueOleDBGrid80.Column
Dim cols As TrueOleDBGrid80.Columns

'On Error GoTo errHandler
On Error Resume Next
Set cols = tdbgHoras.Columns
Dim c As Integer

c = tdbgHoras.col
tdbgHoras.HoldFields
rs4.Filter = getFilter(col, cols)

tdbgHoras.col = c
tdbgHoras.EditActive = True
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
        Select Case rs4.Fields(X).Type
        Case adVarWChar: tmp = tmp & col.DataField & " LIKE '%" & col.FilterText & "%'"
        Case adInteger, adNumeric: tmp = tmp & col.DataField & " = " & col.FilterText
        Case adDBTimeStamp: tmp = tmp & col.DataField & " = #" & col.FilterText & "#"
        End Select
    End If
    X = X + 1
Next col
getFilter = tmp

End Function

Private Sub tdbgHoras_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
If IsNull(rs4!hentrada) Then Me.cmdBorrarCta.Enabled = False
If Not IsNull(rs4!hentrada) Then Me.cmdBorrarCta.Enabled = True
Me.dtpEntrada = IIf(IsNull(rs4!hentrada), "12:00.000 AM", rs4!hentrada)
Me.dtpSalida = IIf(IsNull(rs4!hsalida), "12:00.000 AM", rs4!hsalida)
End Sub

Public Property Let prModal(ByVal val As Boolean)
modal = val
End Property

Private Sub tvActividades_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
Call tvNominas_NodeClick(tvNominas.SelectedItem)

End Sub

Private Sub tvNominas_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo errNode
Dim fecha As Date
fecha = CDate(Mid(tvNominas.SelectedItem.Text, Len(tvNominas.SelectedItem.Text) - 22, 10))
If fecha > Me.mvFecha.MaxDate Then Me.mvFecha.MaxDate = fecha + 1
Me.mvFecha.MinDate = fecha
Me.mvFecha.MaxDate = CDate(Right(tvNominas.SelectedItem.Text, 10))
Me.mvFecha.Value = Me.mvFecha.MinDate

Call mvFecha_DateClick(mvFecha.Value)

Exit Sub
errNode:
    MsgBox Err.Description, vbInformation
End Sub
