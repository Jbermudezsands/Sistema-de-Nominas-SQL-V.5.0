VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form FrmHrsExtras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grabación de Horas Extras"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Horas Extras"
      Height          =   1815
      Left            =   3480
      TabIndex        =   9
      Top             =   360
      Width           =   3135
      Begin VB.TextBox TxtHrasExtras 
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar Horas"
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Horas Extras"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Data DtaNomina 
      Caption         =   "DtaNomina"
      Connect         =   "Access"
      DatabaseName    =   "C:\Zeus Nominas\Nominas.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Nomina"
      Top             =   3120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSDBCtls.DBCombo DbCCodEmpleado 
      Bindings        =   "FrmHrsExtras.frx":0000
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "CodEmpleado"
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin VB.Data DtaHrasExtras 
      Caption         =   "DtaHrasExtras"
      Connect         =   "Access"
      DatabaseName    =   "C:\Zeus Nominas\Nominas.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "HorasExtras"
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data DtaTipoNomina 
      Caption         =   "DtaTipoNomina"
      Connect         =   "Access"
      DatabaseName    =   "C:\Zeus Nominas\Nominas.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TipoNomina"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data DtaEmpleado 
      Caption         =   "DtaEmpleado"
      Connect         =   "Access"
      DatabaseName    =   "C:\Zeus Nominas\Nominas.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Empleado"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox TxtCodNomina 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox TxtApellidos 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox TxtNombres 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "CodNómina"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Apellidos"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "FrmHrsExtras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text3_Change()

End Sub

Private Sub CmdAgregar_Click()
On Error GoTo TipoErrs

If Not IsNumeric(TxtHrasExtras.Text) Then
   MsgBox "La Cantidad de Horas Extras es errónea"
   TxtHrasExtras.SetFocus
   Exit Sub
End If

'pregunto si la nómina de este empleado está activada
DtaTipoNomina.Refresh
Do While Not DtaTipoNomina.Recordset.EOF
If DtaTipoNomina.Recordset("CodTipoNomina") = TxtCodNomina.Text And DtaTipoNomina.Recordset("Activa") = False Then
   MsgBox "La Nómina de ese empleado no ha sido Activada"
   Exit Sub
End If
DtaTipoNomina.Recordset.MoveNext
Loop
'busco el número de la nómina activa de este empleado


'busco en Hrs Extras si ya le fue gravada una hora extra
DtaHrasExtras.Refresh
Do While Not DtaHrasExtras.Recordset.EOF
If DtaHrasExtras.Recordset.Fields("CodEmpleado") = DbCCodEmpleado.Text And DtaHrasExtras.Recordset.NumNomina = DtaNomina.Recordset("NumNomina") Then
   MsgBox "Ya le fueron agregadas horas extras a este empleado, las horas anteriores serán reemplazadas"
   DtaHrasExtras.Recordset.Edit
   DtaHrasExtras.Recordset.Fields("canthoras") = Val(TxtHrasExtras.Text)
   DtaHrasExtras.Recordset.Update
   DtaHrasExtras.Refresh
   TxtHrasExtras.Text = "0"
   DbCCodEmpleado.Text = ""
   TxtNombres.Text = ""
   TxtApellidos.Text = ""
   TxtCodNomina.Text = ""
   Exit Sub
End If

DtaHrasExtras.Recordset.MoveNext
Loop

'si no las encontro grabo las horas extras

   DtaHrasExtras.Recordset.AddNew
   DtaHrasExtras.Recordset.Fields("CodEmpleado") = DbCCodEmpleado.Text
   DtaHrasExtras.Recordset.NumNomina = DtaNomina.Recordset("NumNomina")
   DtaHrasExtras.Recordset.Fields("canthoras") = Val(TxtHrasExtras.Text)
   DtaHrasExtras.Recordset.Update
   
   TxtHrasExtras.Text = "0"
   DbCCodEmpleado.Text = ""
   TxtNombres.Text = ""
   TxtApellidos.Text = ""
   TxtCodNomina.Text = ""
   
Exit Sub

TipoErrs:

ControlErrores
Unload Me

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DbdCodEmpleado_Click(Area As Integer)

End Sub

Private Sub DBCCodEmpleado_Change()
Dim SQLNomina As String
Dim TipoNomina As String


DtaEmpleado.Refresh
'Busco el codigo del empleado para que automaticamente ubique el nombre
 'aunque no existe en la data consulta
    Do While Not DtaEmpleado.Recordset.EOF
     If DtaEmpleado.Recordset("CodEmpleado") = DbCCodEmpleado.Text Then
        TxtNombres.Text = DtaEmpleado.Recordset("Nombre1") + " " + DtaEmpleado.Recordset("Nombre2")
        TxtApellidos.Text = DtaEmpleado.Recordset("Apellido1") + " " + DtaEmpleado.Recordset("Apellido2")
        TxtCodNomina.Text = DtaEmpleado.Recordset("CodTipoNomina")
        Exit Do
     End If
       DtaEmpleado.Recordset.MoveNext
   Loop
   
TipoNomina = TxtCodNomina.Text
   
SQLNomina = "SELECT Nomina.* From Nomina Where Nomina.Activa = True And Nomina.CodTipoNomina = '" & TipoNomina & "' "
DtaNomina.RecordSource = SQLNomina
DtaNomina.Refresh
   
End Sub


