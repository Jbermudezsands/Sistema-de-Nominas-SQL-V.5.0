VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmCalHEFaltas 
   BorderStyle     =   0  'None
   Caption         =   "Calcular Horas Extras"
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6525
   Icon            =   "FrmCalHEFaltas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   235
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   435
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc DtaHrsExtras 
      Height          =   375
      Left            =   120
      Top             =   6240
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
      Caption         =   "DtaHrsExtras"
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
   Begin MSAdodcLib.Adodc DtaSalida 
      Height          =   375
      Left            =   120
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
      Caption         =   "DtaSalida"
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
      Left            =   120
      Top             =   5280
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
   Begin MSAdodcLib.Adodc DtaNominas 
      Height          =   375
      Left            =   120
      Top             =   4800
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
      Caption         =   "DtaNominas"
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
   Begin Project1.xp_canvas xp_canvas1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6165
      Caption         =   "Calcular Horas Extras"
      Fixed_Single    =   -1  'True
      Begin MSDataListLib.DataCombo DBCNominas 
         Bindings        =   "FrmCalHEFaltas.frx":0442
         Height          =   315
         Left            =   2520
         TabIndex        =   14
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Nomina"
         Text            =   ""
      End
      Begin Project1.xphelp xphelp1 
         Height          =   315
         Left            =   5760
         Top             =   80
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
      End
      Begin Project1.xptopbuttons xptopbuttons1 
         Height          =   315
         Left            =   6120
         Top             =   80
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   3720
         TabIndex        =   10
         Top             =   1920
         Width           =   2535
         Begin VB.CommandButton CmdHrsExtras 
            Caption         =   "Calcular Horas Extras"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   2295
         End
         Begin VB.CommandButton CmdDiasDescuento 
            Caption         =   "Calcular Dias de Descuento"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   2295
         End
         Begin VB.CommandButton CmdSalir 
            Caption         =   "Salir"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   2295
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   6135
         Begin VB.TextBox TxtFechaIni 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox TxtFechaFin 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox TxtNumNomina 
            Height          =   375
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Inicio"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Fin"
            Height          =   255
            Left            =   2400
            TabIndex        =   7
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "#"
            Height          =   255
            Left            =   4680
            TabIndex        =   6
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Nóminas Activas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmCalHEFaltas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdDiasDescuento_Click()
Dim CantDias As Integer
Dim DiasMenos As Integer
Dim SalSalida As String
Dim SqlEmpleados As String
Dim CodTipoNomina As String
Dim Fecha As Long
Dim Fechaini As Long
Dim Fechafin As Long
Dim TempFecha As Date



Fecha = CDate(TxtFechaIni.Text)
Fechaini = Val(Fecha)
Fecha = CDate(TxtFechaFin.Text)
Fechafin = Val(Fecha)
CantDias = Fechafin - Fechaini

CodTipoNomina = DtaNominas.Recordset("CodTipoNomina")
SqlEmpleados = "SELECT Empleado.* From Empleado WHERE Empleado.CodTipoNomina= '" & CodTipoNomina & "'"
DtaEmpleados.RecordSource = SqlEmpleados
DtaEmpleados.Refresh

Do While Not DtaEmpleados.Recordset.EOF
    CodEmpleado = DtaEmpleados.Recordset("CodEmpleado")
    DiasMenos = 0
    For I = 0 To (CantDias)
        SqlSalida = "SELECT Salida.CodEmpleado, Salida.Fecha, Salida.HoraEntra, Salida.HoraSale, Salida.HorasExtras From Salida WHERE Salida.CodEmpleado='" & CodEmpleado & "' AND Salida.Fecha= " & Fechaini + I & ""
        TempFecha = Fechaini + I
        
        DtaSalida.RecordSource = SqlSalida
        DtaSalida.Refresh
        If DtaSalida.Recordset.EOF Then
           DiasMenos = DiasMenos + 1
        End If
     Next
        DiasMenos = DiasMenos - 1
        'DtaEmpleados.Recordset.Edit
        DtaEmpleados.Recordset("DiasDescuento") = DiasMenos
        DtaEmpleados.Recordset.Update

DtaEmpleados.Recordset.MoveNext
Loop



End Sub

Private Sub CmdHrsExtras_Click()
Dim SqlEmpleados As String
Dim SQlSalidas As String
Dim SqlHrsExtras As String
Dim CodTipoNomina As String
Dim TotalHrsExtras As String
Dim Fecha As Long
Dim Fechaini As Long
Dim Fechafin As Long

CodTipoNomina = DtaNominas.Recordset("CodTipoNomina")
SqlEmpleados = "SELECT Empleado.* From Empleado WHERE Empleado.CodTipoNomina= '" & CodTipoNomina & "'"
DtaEmpleados.RecordSource = SqlEmpleados
DtaEmpleados.Refresh

Fecha = CDate(TxtFechaIni.Text)
Fechaini = Val(Fecha)

Fecha = CDate(TxtFechaFin.Text)
Fechafin = Val(Fecha)

NumNomina = Val(txtNumNomina.Text)

Do While Not DtaEmpleados.Recordset.EOF
  CodEmpleado = DtaEmpleados.Recordset("CodEmpleado")
  TotalHrsExtras = 0
  SqlSalida = "SELECT Salida.CodEmpleado, Salida.Fecha, Salida.HoraEntra, Salida.HoraSale, Salida.HorasExtras From Salida WHERE Salida.CodEmpleado='" & CodEmpleado & "' AND Salida.Fecha>= " & Fechaini & " AND Salida.Fecha<= " & Fechafin & ""
  DtaSalida.RecordSource = SqlSalida
  DtaSalida.Refresh
  Do While Not DtaSalida.Recordset.EOF
      TotalHrsExtras = TotalHrsExtras + DtaSalida.Recordset("HorasExtras")
  DtaSalida.Recordset.MoveNext
  Loop
  
  Me.DtaHrsExtras.RecordSource = "SELECT Id From HorasExtras"
  Me.DtaHrsExtras.Refresh
  If Me.DtaHrsExtras.Recordset.EOF Then
   ID = 1
  Else
   Me.DtaHrsExtras.Recordset.MoveLast
   ID = Me.DtaHrsExtras.Recordset("id") + 1
  End If

   SqlHrsExtras = "SELECT HorasExtras.id,HorasExtras.CodEmpleado, HorasExtras.NumNomina, HorasExtras.CantHoras, HorasExtras.Pagada From HorasExtras WHERE HorasExtras.CodEmpleado= '" & CodEmpleado & "' AND HorasExtras.NumNomina=" & NumNomina & ""
   DtaHrsExtras.RecordSource = SqlHrsExtras
   DtaHrsExtras.Refresh
   If DtaHrsExtras.Recordset.EOF Then
      DtaHrsExtras.Recordset.AddNew
      DtaHrsExtras.Recordset("id") = ID
      DtaHrsExtras.Recordset("CodEmpleado") = CodEmpleado
      DtaHrsExtras.Recordset("NumNomina") = Val(txtNumNomina.Text)
      DtaHrsExtras.Recordset("canthoras") = TotalHrsExtras
      DtaHrsExtras.Recordset.Update
   Else
      'DtaHrsExtras.Recordset.Edit
      DtaHrsExtras.Recordset("CodEmpleado") = CodEmpleado
      DtaHrsExtras.Recordset("NumNomina") = Val(txtNumNomina.Text)
      DtaHrsExtras.Recordset("canthoras") = TotalHrsExtras
      DtaHrsExtras.Recordset.Update
   End If
  
  
DtaEmpleados.Recordset.MoveNext
Loop

MsgBox "Horas extras Calculadas"

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DBCNominas_Change()
DtaNominas.Refresh
Do While Not DtaNominas.Recordset.EOF
  If DtaNominas.Recordset("nomina") = DBCNominas.Text Then
     TxtFechaIni.Text = DtaNominas.Recordset("FechaNominaINI")
     TxtFechaFin.Text = DtaNominas.Recordset("FechaNomina")
     txtNumNomina.Text = DtaNominas.Recordset("NumNomina")
     Exit Sub
  End If
DtaNominas.Recordset.MoveNext
Loop
End Sub

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub Form_Load()
'Me.DtaNominas '.DatabaseName = Ruta
Me.DtaNominas.ConnectionString = Conexion

'Me.DtaEmpleados '.DatabaseName = Ruta
Me.DtaEmpleados.ConnectionString = Conexion

'Me.DtaSalida '.DatabaseName = Ruta
Me.DtaSalida.ConnectionString = Conexion

'Me.DtaHrsExtras '.DatabaseName = Ruta
Me.DtaHrsExtras.ConnectionString = Conexion

SqlNominas = "SELECT Nomina.NumNomina, Nomina.CodTipoNomina, TipoNomina.Periodo, TipoNomina.Nomina, TipoNomina.Activa, Nomina.Activa, Nomina.FechaNomina, Nomina.FechaNominaIni FROM TipoNomina INNER JOIN Nomina ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina WHERE (((TipoNomina.Activa)=1) AND ((Nomina.Activa)=1))"
DtaNominas.RecordSource = SqlNominas
DtaNominas.Refresh

End Sub

Private Sub xptopbuttons1_Click()
Unload Me
End Sub
