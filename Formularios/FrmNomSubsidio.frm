VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "CODEJO~2.OCX"
Begin VB.Form FrmNomSubsidio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nomina de Subsidios"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin MSDataListLib.DataCombo DBCNominas 
      Bindings        =   "FrmNomSubsidio.frx":0000
      Height          =   315
      Left            =   360
      TabIndex        =   11
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nomina"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DtaDetalleSubsidio 
      Height          =   375
      Left            =   120
      Top             =   7920
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "DtaDetalleSubsidio"
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
   Begin MSAdodcLib.Adodc DtaConsecutivos 
      Height          =   375
      Left            =   120
      Top             =   7560
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "DtaConsecutivos"
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
   Begin MSAdodcLib.Adodc DtaSubsidios 
      Height          =   375
      Left            =   120
      Top             =   7200
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "DtaSubsidios"
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
   Begin MSAdodcLib.Adodc DtaNomSubsidio 
      Height          =   375
      Left            =   120
      Top             =   6840
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "DtaNomSubsidio"
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
   Begin MSAdodcLib.Adodc DtaDetalleNomSubsidio 
      Height          =   375
      Left            =   120
      Top             =   6480
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "DtaDetalleNomSubsidio"
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
      Top             =   6120
      Width           =   3375
      _ExtentX        =   5953
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
      Top             =   5760
      Width           =   3375
      _ExtentX        =   5953
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
   Begin MSAdodcLib.Adodc DtadetalleNomina 
      Height          =   375
      Left            =   120
      Top             =   5400
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "DtadetalleNomina"
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
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3795
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin vbskfree.Skinner Skinner1 
         Left            =   3000
         Top             =   1320
         _ExtentX        =   1270
         _ExtentY        =   1270
         CloseButtonToolTipText=   "Cerrar"
         MinButtonToolTipText=   "Minimizar"
         MaxButtonToolTipText=   "Maximizar"
         RestoreButtonToolTipText=   "Restaurar"
         ChangeControlsBackColor=   0   'False
      End
      Begin VB.Frame Frame1 
         Caption         =   "Movimientos"
         Height          =   1455
         Left            =   720
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
         Begin VB.CommandButton CmdCalcular 
            DownPicture     =   "FrmNomSubsidio.frx":0019
            Height          =   375
            Left            =   120
            Picture         =   "FrmNomSubsidio.frx":1AFB
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton CmdImprimir 
            DownPicture     =   "FrmNomSubsidio.frx":363D
            Height          =   375
            Left            =   120
            Picture         =   "FrmNomSubsidio.frx":511F
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton CmdCerrar 
            DownPicture     =   "FrmNomSubsidio.frx":6C01
            Height          =   375
            Left            =   120
            Picture         =   "FrmNomSubsidio.frx":86E3
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.CommandButton CmdSalir 
         DownPicture     =   "FrmNomSubsidio.frx":9FE5
         Height          =   375
         Left            =   1920
         Picture         =   "FrmNomSubsidio.frx":BAC7
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox TxtNumero 
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox TxtCodTipoNom 
         Height          =   405
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   735
      End
      Begin XtremeSuiteControls.ProgressBar PB1 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   3255
         _Version        =   786432
         _ExtentX        =   5741
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   14737632
         Scrolling       =   1
         Appearance      =   6
      End
      Begin VB.Label Label1 
         Caption         =   "Número de Nomina de Subsidio Actual"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Número:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Código:"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   960
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   120
      Top             =   5040
      Width           =   3375
      _ExtentX        =   5953
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
End
Attribute VB_Name = "FrmNomSubsidio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbNomActiva_Change()

End Sub

Private Sub CmdCalcular_Click()
On Error GoTo TipoErr
Dim Edicion As Boolean
Dim TotalSubsidio As Double
Dim Subsidio As Double
Dim SQLSubsidios As String
Dim CodEmpleado As String
Dim NumSubsidio As Integer
Dim Numero As Long
Dim ELaborada As Boolean

TotalSubsidio = 0
Subsidio = 0

ELaborada = False

If DBCNominas.Text = "" Then
   MsgBox "No ha Selecionado la nómina a ejecutar"
   Exit Sub
   DBCNominas.SetFocus
End If


MousePointer = 11
Abierto = False
'pregunto si hay una nómina de subsidios abierta

DtaNomSubsidio.Refresh
Do While Not DtaNomSubsidio.Recordset.EOF
  If DtaNomSubsidio.Recordset("Activa") = True And DtaNomSubsidio.Recordset("NumNomina") = val(TxtNumero.Text) Then
     Edicion = True
     MsgBox "la nómina de subsidio anterior será reemplazada"
     Exit Do
  End If
  
  If DtaNomSubsidio.Recordset("Activa") = False And DtaNomSubsidio.Recordset("NumNomina") = val(TxtNumero.Text) Then
     ELaborada = True
     MousePointer = 1
     MsgBox "la nómina de subsidio ya fue elaborada y cerrada"
     Exit Do
  End If
  
DtaNomSubsidio.Recordset.MoveNext
Loop

If ELaborada Then
Exit Sub
End If

If Edicion Then
   'DtaNomSubsidio.Recordset.Edit
   DtaNomSubsidio.Recordset("fechapago") = Now
   DtaNomSubsidio.Recordset("TotalnomSubsidio") = 0
   DtaNomSubsidio.Recordset.Update
   
   DtaDetalleNomSubsidio.Refresh
   Do While Not DtaDetalleNomSubsidio.Recordset.EOF
   If DtaDetalleNomSubsidio.Recordset("NumNominaSubsidio") = val(TxtNumero.Text) Then
      DtaDetalleNomSubsidio.Recordset.Delete
   End If
   DtaDetalleNomSubsidio.Recordset.MoveNext
   Loop
Else

   DtaNomSubsidio.Recordset.AddNew
   DtaNomSubsidio.Recordset("NumNomina") = val(TxtNumero.Text)
   DtaNomSubsidio.Recordset("fechapago") = Now
   DtaNomSubsidio.Recordset("TotalnomSubsidio") = 0
   DtaNomSubsidio.Recordset("Activa") = 1
   DtaNomSubsidio.Recordset("Cerrada") = 0
   DtaNomSubsidio.Recordset("Procesada") = 0
   DtaNomSubsidio.Recordset.Update

End If

          
           

DtaEmpleados.Refresh
Do While Not DtaEmpleados.Recordset.EOF
    PB1.Value = PB1.Value + 1
    CodEmpleado = DtaEmpleados.Recordset("CodEmpleado")
    SQLSubsidios = "SELECT Subsidio.NumSubsidio, Subsidio.CodEmpleado, DetalleSubsidio.Valor, DetalleSubsidio.NumVez, DetalleSubsidio.Pagado, DetalleSubsidio.NumNominaSubsidio FROM Subsidio INNER JOIN DetalleSubsidio ON Subsidio.NumSubsidio = DetalleSubsidio.NumSubsidio WHERE DetalleSubsidio.Pagado=0 And Subsidio.CodEmpleado= '" & CodEmpleado & "'"
    DtaSubsidios.RecordSource = SQLSubsidios
    
    Subsidio = 0
      'agregar subsidios
        DtaSubsidios.Refresh
        If Not DtaSubsidios.Recordset.EOF Then
            NumSubsidio = DtaSubsidios.Recordset("NumSubsidio")
            Subsidio = DtaSubsidios.Recordset("valor")
            'DtaSubsidios.Recordset.Edit
            DtaSubsidios.Recordset("NumNominaSubsidio") = val(TxtNumero.Text)
            DtaSubsidios.Recordset.Update

        End If
        
        Me.DtaDetalleSubsidio.Refresh
        
        Do While Not DtaSubsidios.Recordset.EOF
        If NumSubsidio <> DtaSubsidios.Recordset("NumSubsidio") Then
           NumSubsidio = DtaSubsidios.Recordset("NumSubsidio")
           Subsidio = Subsidio + DtaSubsidios.Recordset("valor")
            'DtaSubsidios.Recordset.Edit
            DtaSubsidios.Recordset("NumNominaSubsidio") = val(TxtNumero.Text)
            DtaSubsidios.Recordset.Update
        End If
        
        DtaSubsidios.Recordset.MoveNext
        Loop
        If Subsidio > 0 Then
            Me.DtaDetalleNomSubsidio.Refresh

            If Me.DtaDetalleNomSubsidio.Recordset.EOF Then
              Id = 1
            Else
              Me.DtaDetalleNomSubsidio.Recordset.MoveLast
              Id = DtaDetalleNomSubsidio.Recordset("id") + 1
            End If
            DtaDetalleNomSubsidio.Recordset.AddNew
            DtaDetalleNomSubsidio.Recordset("id") = Id
            DtaDetalleNomSubsidio.Recordset("NumNominaSubsidio") = val(TxtNumero.Text)
            DtaDetalleNomSubsidio.Recordset("CodEmpleado") = CodEmpleado
            DtaDetalleNomSubsidio.Recordset("Subsidio") = Subsidio
            DtaDetalleNomSubsidio.Recordset.Update
        End If
'recorro el detalle de nominas y edito el total de los subsidios
'hago el sql de la nomina actual activa seleccionada
'Numero = Val(TxtNumero.Text)
'
SqlDetalleNomina = "SELECT DetalleNomina.*, DetalleNomina.NumNomina From DetalleNomina WHERE DetalleNomina.NumNomina= " & Numero & " and DetalleNomina.CODEmpleado=  '" & CodEmpleado & "'"
DtaDetalleNomina.RecordSource = SqlDetalleNomina
DtaDetalleNomina.Refresh

If DtaDetalleNomina.Recordset.EOF Then
'   A = MsgBox("No ha sido calculada la nómina correspondiente, no se continuará la ejecución", vbCritical)
'   MousePointer = 1
'   Exit Sub
Else
'    DtaDetalleNomina.Recordset.Edit
    DtaDetalleNomina.Recordset("TotalSubsidio") = Subsidio
    DtaDetalleNomina.Recordset.Update
End If


        
        
TotalSubsidio = TotalSubsidio + Subsidio
DtaEmpleados.Recordset.MoveNext
Loop
  
DtaNomSubsidio.Refresh
Do While Not DtaNomSubsidio.Recordset.EOF
If DtaNomSubsidio.Recordset("NumNomina") = val(TxtNumero.Text) Then
   
   'DtaNomSubsidio.Recordset.Edit
   DtaNomSubsidio.Recordset("TotalnomSubsidio") = TotalSubsidio
   DtaNomSubsidio.Recordset.Update
   Exit Do
End If

DtaNomSubsidio.Recordset.MoveNext
Loop

MousePointer = 1
MsgBox ("La Nómina de Subsidio " + TxtNumero.Text + " fue creada con éxito")

Exit Sub

TipoErr:
ControlErrores
End Sub

Private Sub cmdCerrar_Click()
'On Error GoTo TipoErr
Dim Cerrar As Boolean
Dim SQLSubsidios As String
Dim Letra As String
Dim SqlEmpleados As String
Dim rs As New ADODB.Recordset

Letra = "n"

If DBCNominas.Text = "" Then
   MsgBox "No ha Selecionado la nómina a ejecutar"
   Exit Sub
   DBCNominas.SetFocus
End If

SqlEmpleados = "SELECT Empleado.*, Empleado.CodTipoNomina From Empleado WHERE Empleado.CodTipoNomina= '" & TxtCodTipoNom.Text & "'"
DtaEmpleados.RecordSource = SqlEmpleados
DtaEmpleados.Refresh
DtaEmpleados.Recordset.MoveLast
CantEmpleados = DtaEmpleados.Recordset.RecordCount
DtaEmpleados.Recordset.MoveFirst
With PB1
 .Min = 0
 .Max = CantEmpleados
 .Value = 0
 i = 1
Do While Not DtaEmpleados.Recordset.EOF

PB1.Value = PB1.Value + 1

CodEmpleado = DtaEmpleados.Recordset("CodEmpleado")
SQLSubsidios = "SELECT Subsidio.NumSubsidio, Subsidio.CodEmpleado, DetalleSubsidio.Descripcion, DetalleSubsidio.Valor, DetalleSubsidio.NumVez, DetalleSubsidio.Descripcion, DetalleSubsidio.NumNominaSubsidio, DetalleSubsidio.Pagado FROM Subsidio INNER JOIN DetalleSubsidio ON Subsidio.NumSubsidio = DetalleSubsidio.NumSubsidio WHERE Subsidio.CodEmpleado='" & CodEmpleado & "' AND DetalleSubsidio.NumVez='" & Letra & "'"
DtaSubsidios.RecordSource = SQLSubsidios
DtaSubsidios.Refresh
  Do While Not DtaSubsidios.Recordset.EOF
        Me.DtaDetalleSubsidio.Refresh
        Me.DtaDetalleSubsidio.Recordset.MoveLast
        If Me.DtaDetalleSubsidio.Recordset.EOF Then
          Id = 1
        Else
          Id = DtaDetalleSubsidio.Recordset("id") + 1
        End If
        DtaDetalleSubsidio.Recordset.AddNew
        DtaDetalleSubsidio.Recordset("id") = Id
        DtaDetalleSubsidio.Recordset("NumSubsidio") = DtaSubsidios.Recordset("NumSubsidio")
        DtaDetalleSubsidio.Recordset("Descripcion") = DtaSubsidios.Recordset("Descripcion")
        DtaDetalleSubsidio.Recordset("valor") = DtaSubsidios.Recordset("valor")
        DtaDetalleSubsidio.Recordset("NumVez") = 1
        DtaDetalleSubsidio.Recordset("NumNominaSubsidio") = val(TxtNumero.Text)
        DtaDetalleSubsidio.Recordset("Pagado") = 0
        DtaDetalleSubsidio.Recordset.Update
  DtaSubsidios.Recordset.MoveNext
  Loop

DtaEmpleados.Recordset.MoveNext
Loop
End With

Cerrar = False

Me.DtaNomSubsidio.Refresh
Do While Not DtaNomSubsidio.Recordset.EOF
If DtaNomSubsidio.Recordset("NumNomina") = val(TxtNumero.Text) Then
   ''DtaNomSubsidio.Recordset.Edit
   DtaNomSubsidio.Recordset("Activa") = False
   DtaNomSubsidio.Recordset("cerrada") = True
   DtaNomSubsidio.Recordset.Update
   Cerrar = True
   Exit Do
End If
DtaNomSubsidio.Recordset.MoveNext
Loop

If Not Cerrar Then
    MsgBox ("La nómina de Subsidios " & TxtNumero.Text & "  no ha sido elaborada")
    Exit Sub
End If



   
   rs.Open "UPDATE DetalleSubsidio SET DetalleSubsidio.Pagado = 1 WHERE Detallesubsidio.NumNominasubsidio= " & TxtNumero.Text & " AND Detallesubsidio.NUmvez<> '" & Letra & "'", Conexion

MsgBox "La Nomina Ha sido Cerrada"



Unload Me

Exit Sub
TipoErr:
ControlErrores
End Sub

Private Sub cmdImprimir_Click()
On Error GoTo TipoErr
If DBCNominas.Text = "" Then
   MsgBox "No ha Selccionado la nómina a ejecutar"
   Exit Sub
   DBCNominas.SetFocus
End If

NumNominaSubsidio = val(TxtNumero.Text)
ArepNomSubsidio.DataControl1.ConnectionString = ConexionReporte
ArepNomSubsidio.DataControl1.Source = "SELECT  Empleado.CodEmpleado1,Empleado.CodEmpleado, Empleado.Nombre1, Empleado.Nombre2, Empleado.Apellido1, Empleado.Apellido2, Subsidio.NumSubsidio, TipoSubsidio.Subsidio, DetalleSubsidio.Valor, DetalleSubsidio.NumVez, DetalleSubsidio.Descripcion, DetalleSubsidio.NumNominaSubsidio FROM TipoSubsidio INNER JOIN ((Empleado INNER JOIN Subsidio ON Empleado.CodEmpleado = Subsidio.CodEmpleado) INNER JOIN DetalleSubsidio ON Subsidio.NumSubsidio = DetalleSubsidio.NumSubsidio) ON TipoSubsidio.CodTipoSubsidio = Subsidio.CodTipoSubsidio Where (((DetalleSubsidio.NumNominaSubsidio) = " & NumNominaSubsidio & " )) ORDER BY Empleado.CodEmpleado"
ArepNomSubsidio.LblTitulo.Caption = Titulo
ArepNomSubsidio.LblSubtitulo.Caption = SubTitulo
ArepNomSubsidio.ImgLogo.Picture = LoadPicture(RutaLogo)
ArepNomSubsidio.LblFecha.Caption = Format(Now, "dd/mm/yyyy")
ArepNomSubsidio.Show 1
'NumReport = 2
'FrmVerReportes.Show

Exit Sub
TipoErr:
ControlErrores

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub DBCNominas_Change()
'On Error GoTo TipoErr

If DBCNominas = "" Then
   MsgBox "Error en el tipo de nómina a calcular"
   Exit Sub
End If

MousePointer = 11
Dim SqlEmpleados As String
Dim TipoNominas As String
Dim SqlDetalleNomina As String

DtaNominas.Refresh
Do While Not DtaNominas.Recordset.EOF
If DtaNominas.Recordset("nomina") = DBCNominas.Text Then
   TxtNumero = DtaNominas.Recordset("NumNomina")
  TxtCodTipoNom.Text = DtaNominas.Recordset("CodTipoNomina")
   TipoNominas = TxtCodTipoNom.Text
   Exit Do
End If
DtaNominas.Recordset.MoveNext
Loop

'hago el sql de los empleados de este tipo de nóminas

SqlEmpleados = "SELECT Empleado.*, Empleado.CodTipoNomina From Empleado WHERE Empleado.CodTipoNomina= '" & TipoNominas & "' AND Empleado.Activo=1"
DtaEmpleados.RecordSource = SqlEmpleados
DtaEmpleados.Refresh
If Not Me.DtaEmpleados.Recordset.EOF Then
  DtaEmpleados.Recordset.MoveLast
  PB1.Min = 0
  PB1.Max = DtaEmpleados.Recordset.RecordCount
  DtaEmpleados.Refresh
Else
  MsgBox "No Existe Ningun Empleado en esta Nomina", vbInformation, "Sistema de Nominas"
  Exit Sub
End If
MousePointer = 1
Exit Sub
TipoErr:
ControlErrores

End Sub

Private Sub Form_Load()
Dim SqlNominas As String

With Me.DtaConsecutivos
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaDetalleNomina
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaDetalleNomSubsidio
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "DetalleNomSubsidio"
   .Refresh
End With

With Me.DtaDetalleSubsidio
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "DetalleSubsidio"
   .Refresh
End With

With Me.DtaEmpleados
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaConsulta
   .ConnectionString = Conexion
End With

With Me.DtaNominas
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

With Me.DtaNomSubsidio
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
   .RecordSource = "NomSubsidio"
   .Refresh
End With

With Me.DtaSubsidios
   '.DatabaseName = Ruta
   .ConnectionString = Conexion
End With

SqlNominas = "SELECT Nomina.NumNomina, Nomina.CodTipoNomina, TipoNomina.Periodo, TipoNomina.Nomina, TipoNomina.Activa, Nomina.Activa FROM TipoNomina INNER JOIN Nomina ON TipoNomina.CodTipoNomina = Nomina.CodTipoNomina WHERE (((TipoNomina.Activa)=1) AND ((Nomina.Activa)=1))"
DtaNominas.RecordSource = SqlNominas
DtaNominas.Refresh

End Sub

