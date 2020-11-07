VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form frmAsistManual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso Manual de Asistencia"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7680
   Begin MSAdodcLib.Adodc adoEmpleado 
      Height          =   375
      Left            =   480
      Top             =   6600
      Width           =   4695
      _ExtentX        =   8281
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
      Caption         =   "adoEmpleado"
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
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      TabIndex        =   27
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   6240
      TabIndex        =   26
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      TabIndex        =   25
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   5040
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc adoAsistencia 
      Height          =   330
      Left            =   480
      Top             =   6120
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"frmAsistenciaManual.frx":0000
      OLEDBString     =   $"frmAsistenciaManual.frx":008C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Asistencia Diaria"
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
   Begin VB.Frame Frame3 
      Caption         =   "Ingreso - Modificación Manual"
      Height          =   1575
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   7455
      Begin VB.CheckBox chkSalidaManual 
         Caption         =   "Esta Laborando este dia"
         Height          =   255
         Left            =   5040
         TabIndex        =   7
         Top             =   1080
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpFInicio 
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   83230721
         CurrentDate     =   38570
      End
      Begin MSMask.MaskEdBox mskPermisoHoraInicio 
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm AM/PM"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPermisoHoraRegreso 
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm AM/PM"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         Height          =   195
         Left            =   3120
         TabIndex        =   22
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Hora:"
         Height          =   195
         Left            =   3120
         TabIndex        =   21
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   870
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Asistencia Diaria a la empresa"
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   7455
      Begin VB.CheckBox chkSalida 
         Caption         =   "Esta Laborando este dia"
         Height          =   255
         Left            =   5160
         TabIndex        =   18
         Top             =   960
         Width           =   2175
      End
      Begin MSMask.MaskEdBox mskAsistHoraEntrada 
         Height          =   255
         Left            =   3960
         TabIndex        =   15
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm AM/PM"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpFechEntrada 
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   83230721
         CurrentDate     =   38570
      End
      Begin MSMask.MaskEdBox mskAsistHoraSalida 
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm AM/PM"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpFecSalida 
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   83230721
         CurrentDate     =   38570
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Salida"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Hora Salida:"
         Height          =   195
         Left            =   2880
         TabIndex        =   16
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hora Entrada"
         Height          =   195
         Left            =   2880
         TabIndex        =   14
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Entrada"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Empleado"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   5640
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   255
         Left            =   3960
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   83230721
         CurrentDate     =   38570
      End
      Begin TrueOleDBList80.TDBCombo cboCodigo 
         Bindings        =   "frmAsistenciaManual.frx":0118
         Height          =   315
         Left            =   840
         TabIndex        =   28
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   0
         _EDITHEIGHT     =   556
         _GAPHEIGHT      =   53
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
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits.Count    =   1
         Appearance      =   1
         BorderStyle     =   1
         ComboStyle      =   0
         AutoCompletion  =   0   'False
         LimitToList     =   0   'False
         ColumnHeaders   =   -1  'True
         ColumnFooters   =   0   'False
         DataMode        =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         FootLines       =   1
         RowDividerStyle =   0
         Caption         =   ""
         EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         LayoutName      =   ""
         LayoutFileName  =   ""
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   0
         AutoSize        =   -1  'True
         ListField       =   "CodEmpleado1"
         BoundColumn     =   ""
         IntegralHeight  =   0   'False
         CellTipsWidth   =   0
         CellTipsDelay   =   1000
         AutoDropdown    =   0   'False
         RowTracking     =   -1  'True
         RightToLeft     =   0   'False
         RowMember       =   ""
         MouseIcon       =   0
         MouseIcon.vt    =   3
         MousePointer    =   0
         MatchEntryTimeout=   2000
         OLEDragMode     =   0
         OLEDropMode     =   0
         AnimateWindow   =   0
         AnimateWindowDirection=   0
         AnimateWindowTime=   200
         AnimateWindowClose=   0
         DropdownPosition=   0
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         AddItemSeparator=   ";"
         _PropDict       =   $"frmAsistenciaManual.frx":0132
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
      Begin VB.Label lblNombre 
         Caption         =   " "
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   1080
         Width           =   5295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   3480
         TabIndex        =   9
         Top             =   480
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmAsistManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sCodEmpl As String
Public dFecha As Date
Public sCodTipoNomina As String
Public dCodEmpleado As Double




Private Sub cboCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   Me.dtpFecha.SetFocus

End If

End Sub

Private Sub CMD1_Click()

CMD1.Caption = "XXXX"


End Sub

Private Sub CmdAgregar_Click()

Dim dFecha As Date
Dim sFechaEntrada As String
'Dim cnDB As New ADODB.Connection
'Dim rsDB As New ADODB.Recordset
Dim dCodigoEmpl As Double

'cnDB.ConnectionString = "Provider=SQLOLEDB.1;Password=metro;Persist Security Info=True;User ID=metro;Initial Catalog=SistemasNominas;Data Source=METRO"
'cnDB.Open


If Trim(Me.cboCodigo.Text) <> "" Then

   
   
'   If rsDB.EOF Then
'      MsgBox "El codigo de empleado No es valido", vbInformation + vbOKOnly
'      rsDB.Close
'      cnDB.Close
'      Exit Sub
'
'   End If


   If Me.CmdAgregar.Caption = "&Agregar" Then
        
      Me.cmdModificar.Enabled = False
      Me.dtpFInicio.SetFocus
      Me.CmdAgregar.Caption = "&Guardar"
      Me.CmdSalir.Caption = "&Cancelar"
     
   ElseIf Chequear_Hora(Me.mskPermisoHoraInicio.Text) Then
    
      dFecha = Me.dtpFecha.Value
      
      sFechaEntrada = Mid$(dFecha, 7, 4) & "-" & Mid$(dFecha, 4, 2) & "-" & Mid$(dFecha, 1, 2)
      
      
      Me.adoAsistencia.CommandType = adCmdText
      Me.adoAsistencia.RecordSource = "SELECT CodEmpleado, CodEmpleado1, CodTipoNomina, FechaEntrada, FechaSalida, HoraEntrada, HREntrada, HRSalida, HoraSalida, bActivo, CodTurno, HLaboradas, HExtras " & _
                                     "FROM AsistenciaEmpleado WHERE FechaEntrada = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102) AND CodEmpleado1 ='" & Me.cboCodigo.Text & "'"
      Me.adoAsistencia.Refresh
                       
     Me.adoEmpleado.CommandType = adCmdText
     Me.adoEmpleado.RecordSource = "SELECT * FROM Empleado WHERE CodEmpleado1 LIKE '" & Me.cboCodigo.Text & "' AND Activo =1"
     Me.adoEmpleado.Refresh
     
     sCodEmpl = Me.cboCodigo.Text
                       
      If Me.adoAsistencia.Recordset.EOF Then
      
         Me.adoAsistencia.Recordset.AddNew
         
         Me.adoAsistencia.Recordset.Fields("CodEmpleado") = Me.adoEmpleado.Recordset.Fields("CodEmpleado")
         Me.adoAsistencia.Recordset.Fields("CodEmpleado1") = Me.cboCodigo.Text
        
         Me.adoAsistencia.Recordset.Fields("CodTipoNomina") = sCodTipoNomina
         Me.adoAsistencia.Recordset.Fields("FechaEntrada") = Me.dtpFInicio.Value
         Me.adoAsistencia.Recordset.Fields("HoraEntrada") = Me.mskPermisoHoraInicio.Text
         Me.adoAsistencia.Recordset.Fields("HREntrada") = Me.mskPermisoHoraInicio.Text
         
         
         If Not Me.chkSalidaManual.Value And Chequear_Hora(Me.mskPermisoHoraRegreso.Text) Then
            Me.adoAsistencia.Recordset.Fields("FechaSalida") = Me.dtpFInicio.Value
            Me.adoAsistencia.Recordset.Fields("HoraSalida") = Me.mskPermisoHoraRegreso.Text
            Me.adoAsistencia.Recordset.Fields("HRSalida") = Me.mskPermisoHoraRegreso.Text
            Me.adoAsistencia.Recordset.Fields("bActivo") = 0
         ElseIf Not Me.chkSalida.Value Then
            MsgBox "Debe de digitar la hora de salida correcta, verifique"
            Me.adoAsistencia.Recordset.CancelUpdate
            Exit Sub
         Else
            Me.adoAsistencia.Recordset.Fields("bActivo") = 1
         End If
          
         
         Me.adoAsistencia.Recordset.Fields("CodTurno") = "Diurno"
         Me.adoAsistencia.Recordset.Update
         
      Else
         MsgBox "Se tiene registrada una asistencia de este empleado para este dia, modifique el registro"
         
            
      End If
      
      
     Me.cmdModificar.Enabled = True
     Me.CmdAgregar.Caption = "&Agregar"

   End If


End If

End Sub

Private Sub cmdborrar_Click()


dFecha = Me.dtpFecha.Value
sFechaEntrada = Mid$(dFecha, 7, 4) & "-" & Mid$(dFecha, 4, 2) & "-" & Mid$(dFecha, 1, 2)


If MsgBox("Esta seguro que desea borrar la asistencia del empleado " & sCodEmpl & " en la fecha: " & Me.dtpFecha.Value, vbYesNo, "Asistencia Manual") = vbYes Then
    
  Me.adoAsistencia.CommandType = adCmdText
  Me.adoAsistencia.RecordSource = "SELECT CodEmpleado, CodEmpleado1, FechaEntrada, FechaSalida, HoraEntrada, HoraSalida, bActivo FROM AsistenciaEmpleado WHERE [FechaEntrada] =CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102) AND [CodEmpleado1] ='" & sCodEmpl & "'"
  Me.adoAsistencia.Refresh

  Me.adoAsistencia.Recordset.Delete
  Me.adoAsistencia.Refresh
  
  Me.mskAsistHoraEntrada.Text = "__:__:__"
  Me.mskAsistHoraSalida.Text = "__:__:__"
        
  Me.mskPermisoHoraInicio.Text = "__:__:__"
  Me.mskPermisoHoraRegreso.Text = "__:__:__"
  
  
End If



End Sub

Private Sub cmdBuscar_Click()

Dim sFechaEntrada As String

If Trim(Me.cboCodigo.Text) <> "" Then
   
   Me.lblNombre.Caption = ""
   
   Me.adoEmpleado.Refresh
   'Me.adoEmpleado.Recordset.Find "CodEmpleado1 LIKE '" & Trim(Me.cboCodigo.text) & "'"
   Me.adoEmpleado.CommandType = adCmdText
   Me.adoEmpleado.RecordSource = "SELECT *, CodEmpleado1 From Empleado WHERE (CodEmpleado1 <> N'IS NULL') AND Activo =1 AND CodEmpleado1 LIKE '" & Trim(Me.cboCodigo.Text) & "'"
   Me.adoEmpleado.Refresh
    
   Me.dtpFInicio.Value = dtpFecha.Value
    
   If Not Me.adoEmpleado.Recordset.EOF Then
      
      Me.lblNombre.Caption = Me.adoEmpleado.Recordset.Fields("Nombre1") & " " & Me.adoEmpleado.Recordset.Fields("Nombre2") & " " & Me.adoEmpleado.Recordset.Fields("Apellido1") & " " & Me.adoEmpleado.Recordset.Fields("Apellido2")
      sCodEmpl = Me.adoEmpleado.Recordset.Fields("CodEmpleado1")
      dCodEmpleado = Me.adoEmpleado.Recordset.Fields("CodEmpleado")
      sCodTipoNomina = Me.adoEmpleado.Recordset.Fields("CodTipoNomina")
      dFecha = Me.dtpFecha.Value
      
     sFechaEntrada = Mid$(dFecha, 7, 4) & "-" & Mid$(dFecha, 4, 2) & "-" & Mid$(dFecha, 1, 2)
      
     'Me.adoAsistencia.Recordset.Find "[FechaEntrada] = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102)"
      
      Me.adoAsistencia.CommandType = adCmdText
      Me.adoAsistencia.RecordSource = "SELECT CodEmpleado, CodEmpleado1, FechaEntrada, FechaSalida, HoraEntrada, HoraSalida, bActivo FROM AsistenciaEmpleado WHERE [FechaEntrada] = CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102) AND [CodEmpleado1] ='" & sCodEmpl & "'"
      Me.adoAsistencia.Refresh
      
      'Me.adoAsistencia.Recordset.Find "[FechaEntrada] = '" & dFecha & "' AND [CodEmpleado] ='" & sCodEmpl & "'"
      
      
      If Not Me.adoAsistencia.Recordset.EOF Then
         Me.dtpFechEntrada.Value = Me.adoAsistencia.Recordset.Fields("FechaEntrada")
         Me.mskAsistHoraEntrada.Text = Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss")
         Me.mskPermisoHoraInicio.Text = Format(Me.adoAsistencia.Recordset.Fields("HoraEntrada"), "hh:mm:ss")
         Me.cmdborrar.Enabled = True
         
         If Not IsNull(Me.adoAsistencia.Recordset.Fields("FechaSalida")) Then
            Me.dtpFecSalida.Value = Me.adoAsistencia.Recordset.Fields("FechaSalida")
            Me.mskAsistHoraSalida.Text = Format(Me.adoAsistencia.Recordset.Fields("HoraSalida"), "hh:mm:ss")
         Else
            Me.lblNombre.Caption = Me.lblNombre.Caption & ", NO tiene fecha y hora de salida este dia"
            Me.mskAsistHoraSalida.Text = "__:__:__"
         End If
           
        If Not IsNull(Me.adoAsistencia.Recordset.Fields("bActivo")) Then
           
          If Me.adoAsistencia.Recordset.Fields("bActivo") Then
             Me.chkSalida.Value = 1
          Else
             Me.chkSalida.Value = 0
          End If
       End If
       
'         Me.adoPermiso.CommandType = adCmdText
'         Me.adoPermiso.RecordSource = "SELECT * FROM Permisos WHERE [Fecha] ='" & dFecha & "' AND [CodEmpleado] ='" & sCodEmpl & "'"
'         Me.adoPermiso.Refresh

         
         'Me.adoPermiso.Recordset.Find "[Fecha] ='" & dFecha & "' AND [CodEmpleado] ='" & sCodEmpl & "'"
         
'         If Not Me.adoPermiso.Recordset.EOF Then
'            Me.dtpFInicio.Value = Me.adoPermiso.Recordset.Fields("Fecha")
'            Me.mskPermisoHoraInicio.Text = Me.adoPermiso.Recordset.Fields("HoraInicio")
'            Me.txtMotivo.Text = Me.adoPermiso.Recordset.Fields("Motivo")
'
'            If Me.adoPermiso.Recordset.Fields("Justificado") Then
'               Me.chkJustificado.Value = 1
'            Else
'               Me.chkJustificado.Value = 0
'            End If
'
'            Me.cmdModificar.Enabled = False
'            Me.cmdAgregar.Enabled = False
'
'            If Me.adoPermiso.Recordset.Fields("RegresoPendiente") Then
'               Me.chkRegreso.Value = 1
'
'            Else
'
'               Me.chkRegreso.Value = 0
'
'
'            End If
            
            Me.CmdAgregar.Enabled = True
            Me.cmdModificar.Enabled = True
            
'            If Not IsNull(Me.adoPermiso.Recordset.Fields("HoraFin")) Then
'               Me.mskPermisoHoraRegreso.Text = Me.adoPermiso.Recordset.Fields("HoraFin")
'
'            Else
               Me.mskPermisoHoraRegreso.Text = "__:__:__"
            
'            End If
            
            'Me.adoPermiso.Recordset.Update
            
         'Sino tiene un permiso para este dia, limpiar valores
         Else
           Me.mskAsistHoraEntrada.Text = "__:__:__"
           Me.mskAsistHoraSalida.Text = "__:__:__"
           
           Me.CmdAgregar.Enabled = True
           Me.dtpFInicio.Value = Me.dtpFecha.Value
           Me.mskPermisoHoraInicio.Text = "__:__:__"
'           Me.txtMotivo.Text = " "
'           Me.chkJustificado.Value = False
'           Me.chkRegreso.Value = 1
            
            
         End If
         
         
      'Sino tiene asistencia ese dia, limpiamos los valores de la asistencia.
      Else
        
        Me.mskAsistHoraEntrada.Text = "__:__:__"
        Me.mskAsistHoraSalida.Text = "__:__:__"
        
        Me.mskPermisoHoraInicio.Text = "__:__:__"
        Me.mskPermisoHoraRegreso.Text = "__:__:__"
'        Me.txtMotivo.Text = " "
'        Me.chkJustificado.Value = 0
        Me.lblNombre.Caption = Me.lblNombre.Caption & ", No asistio a trabajar el " & Me.dtpFecha.Value
        Me.mskAsistHoraEntrada.Text = "__:__:__"
        Me.mskAsistHoraSalida.Text = "__:__:__"
        Me.chkSalida.Value = 0
      
         
         
      End If
      
      
      Me.adoEmpleado.Refresh
      
   Else
   
      MsgBox "El empleado No. " & Me.cboCodigo.Text & " no se encuentra registrado", vbInformation, "Permisos - Nomina"
      Me.lblNombre.Caption = "Empleado No Encontrado"
      Me.mskAsistHoraEntrada.Text = "__:__:__"
      Me.mskAsistHoraSalida.Text = "__:__:__"
      Me.adoEmpleado.Refresh
      
   End If
    





End Sub

Private Sub cmdModificar_Click()



If Chequear_Hora(Me.mskPermisoHoraInicio.Text) And Me.cmdModificar.Caption = "&Modificar" And sCodEmpl <> "" Then
  
  Me.CmdAgregar.Enabled = False
  Me.mskPermisoHoraInicio.SetFocus
  Me.cmdModificar.Caption = "&Guardar Cambios"
  
  

ElseIf Chequear_Hora(Me.mskPermisoHoraInicio.Text) And sCodEmpl <> "" Then
  
  dFecha = Me.dtpFecha.Value
      
  sFechaEntrada = Mid$(dFecha, 7, 4) & "-" & Mid$(dFecha, 4, 2) & "-" & Mid$(dFecha, 1, 2)
      
  sCodEmpl = Me.cboCodigo.Text
  
  Me.adoAsistencia.CommandType = adCmdText
  Me.adoAsistencia.RecordSource = "SELECT CodEmpleado, CodEmpleado1, FechaEntrada, HREntrada, FechaSalida, HoraEntrada, HoraSalida, HRSalida, bActivo FROM AsistenciaEmpleado WHERE [FechaEntrada] =CONVERT(DATETIME, '" & sFechaEntrada & " 00:00:00" & "', 102) AND [CodEmpleado1] ='" & sCodEmpl & "'"
  Me.adoAsistencia.Refresh
      
           
  If Not Me.adoAsistencia.Recordset.EOF Then
           
           
  'Me.adoAsistencia.Recordset.Fields ("CodEmpleado")
  Me.adoAsistencia.Recordset.Fields("FechaEntrada") = Me.dtpFechEntrada.Value
  Me.adoAsistencia.Recordset.Fields("HoraEntrada") = Me.mskPermisoHoraInicio.Text
  Me.adoAsistencia.Recordset.Fields("HREntrada") = Me.mskPermisoHoraInicio.Text
  
  If Chequear_Hora(Me.mskPermisoHoraRegreso.Text) Then
     Me.adoAsistencia.Recordset.Fields("FechaSalida") = Me.dtpFInicio.Value
     Me.adoAsistencia.Recordset.Fields("HoraSalida") = Me.mskPermisoHoraRegreso.Text
     Me.adoAsistencia.Recordset.Fields("HRSalida") = Me.mskPermisoHoraRegreso.Text
     
  End If
         
  If Me.chkSalidaManual.Value Then
     Me.adoAsistencia.Recordset.Fields("bActivo") = 1
  ElseIf Chequear_Hora(Me.mskAsistHoraSalida.Text) Then
     Me.adoAsistencia.Recordset.Fields("bActivo") = 0
  End If
              
  Me.adoAsistencia.Recordset.Update
                      
  End If
  
  
  
  Me.CmdAgregar.Enabled = True
  Me.cmdModificar.Caption = "&Modificar"
  

End If


End Sub

Private Sub CmdSalir_Click()

 If Me.CmdSalir.Caption = "&Salir" Then
    Unload Me
 Else
    Me.CmdSalir.Caption = "&Salir"
       
    Me.cmdModificar.Enabled = True
    Me.CmdAgregar.Caption = "&Agregar"
    Me.cboCodigo.SetFocus
    
 End If
 
 


End Sub


Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
  Me.cmdBuscar.SetFocus

End If

End Sub

Private Sub Form_Activate()

Me.dtpFecha.Value = Mid$(Now, 1, 10)
Me.dtpFInicio.Value = Mid$(Now, 1, 10)
Me.cboCodigo.SetFocus



End Sub

Private Sub Form_Load()

 Dim RutaServer As String
 Dim Server As String
' Dim Conexion As String

'
' Dim ConexionSTR1 As String
' Dim TxtClaveEntrada As String
''abro el archivo para solo lectura de la cadena de conexion
' Dim NextLine As String
' Dim Autorizado As Boolean
'   Autorizado = False

' Open App.Path + "\SysInfo.dll" For Input As #1
'  Do Until EOF(1)
'   Line Input #1, NextLine
'        ConexionSTR1 = Trim(NextLine)
'   Loop
' Close #1
  
 
'
'  Conexion = ConexionSTR1
 
 
Me.dtpFecha.Value = Format(Now, "dd/mm/yyyy")
Me.dtpFechEntrada.Value = Format(Now, "dd/mm/yyyy")
Me.dtpFecSalida.Value = Format(Now, "dd/mm/yyyy")
Me.dtpFInicio.Value = Format(Now, "dd/mm/yyyy")



With Me.adoEmpleado
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "SELECT CodEmpleado1, Nombre1 + ' '+ Nombre2 +' '+Apellido1+' '+Apellido2 as Nombres From Empleado Where (Activo = 1) ORDER BY CodEmpleado1"
   .Refresh
End With

With Me.adoAsistencia
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "AsistenciaEmpleado"
   .Refresh
End With


'
'Me.adoEmpleado.ConnectionString = Conexion
''Me.adoEmpleado.CommandType = adCmdText
'Me.adoEmpleado.RecordSource = "SELECT *, CodEmpleado1 From Empleado WHERE (CodEmpleado1 <> N'IS NULL') AND Activo =1"
'Me.adoEmpleado.Refresh


'Me.adoAsistencia.ConnectionString = Conexion
''Me.adoAsistencia.CommandType = adCmdTable
'Me.adoAsistencia.RecordSource = "AsistenciaEmpleado"
'Me.adoAsistencia.Refresh

'Me.adoPermiso.ConnectionString = Conexion
'Me.adoPermiso.CommandType = adCmdTable
'Me.adoPermiso.RecordSource = "HorarioEmpleado"
'Me.adoPermiso.Refresh



'Do While Not Me.adoEmpleado.Recordset.EOF
'
'
'  Me.cboCodigo.AddItem Me.adoEmpleado.Recordset.Fields("CodEmpleado1")
'  Me.adoEmpleado.Recordset.MoveNext
'
'
'Loop
'
'Me.adoEmpleado.Refresh

End Sub


Public Function Chequear_Hora(Hora As String) As Boolean

Dim sMinutos As String
Dim sHora As String
Dim sSegundos As String

On Error GoTo ManejarError

sHora = Mid$(Hora, 1, 2)
sMinutos = Mid$(Hora, 4, 2)
sSegundos = Mid$(Hora, 7, 2)


If (IsNumeric(sHora) And CInt(sHora) <= 23) And (IsNumeric(sMinutos) And CInt(sMinutos) <= 59) And (IsNumeric(sSegundos) And CInt(sMinutos) <= 59) Then

   Chequear_Hora = True

Else
   
   Chequear_Hora = False

End If

Exit Function

ManejarError:

Chequear_Hora = False






End Function

